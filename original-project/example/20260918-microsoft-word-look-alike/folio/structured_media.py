from __future__ import annotations

import base64
import binascii
import json
import math
from dataclasses import dataclass
from pathlib import Path
import re
import struct
from urllib.parse import unquote, urlsplit
import xml.etree.ElementTree as ET

from PySide6.QtCore import QPointF, QRectF, QSize, Qt
from PySide6.QtGui import QColor, QFont, QFontMetricsF, QImage, QPainter, QPen, QPixmap, QPolygonF
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import QComboBox, QDialog, QDialogButtonBox, QLabel, QPlainTextEdit, QScrollArea, QVBoxLayout


MAX_EQUATION_LENGTH = 16000
MAX_MEDIA_BYTES = 32 * 1024 * 1024
MAX_MODEL_VERTICES = 100000
MAX_MODEL_TRIANGLES = 100000


class MediaError(ValueError):
    pass


@dataclass(frozen=True)
class EquationNode:
    kind: str
    value: str = ""
    children: tuple[EquationNode, ...] = ()


_TEX_SYMBOLS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "theta": "θ", "lambda": "λ", "mu": "μ", "pi": "π", "rho": "ρ",
    "sigma": "σ", "tau": "τ", "phi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Pi": "Π",
    "Sigma": "Σ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω", "infty": "∞",
    "times": "×", "cdot": "·", "div": "÷", "pm": "±", "mp": "∓",
    "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "ne": "≠", "neq": "≠",
    "approx": "≈", "equiv": "≡", "to": "→", "rightarrow": "→", "leftarrow": "←",
    "partial": "∂", "nabla": "∇", "in": "∈", "notin": "∉", "forall": "∀",
    "exists": "∃", "emptyset": "∅", "ldots": "…", "cdots": "⋯",
    "sin": "sin", "cos": "cos", "tan": "tan", "log": "log", "ln": "ln",
    "exp": "exp", "lim": "lim", "min": "min", "max": "max",
}


class _TexParser:
    def __init__(self, source):
        self.source = source
        self.index = 0
        self.nodes = 0

    def node(self, kind, value="", children=()):
        self.nodes += 1
        if self.nodes > 2048:
            raise MediaError("Equation contains too many elements.")
        return EquationNode(kind, value, tuple(children))

    def skip_space(self):
        while self.index < len(self.source) and self.source[self.index].isspace():
            self.index += 1

    def group(self, depth):
        self.skip_space()
        if self.index >= len(self.source) or self.source[self.index] != "{":
            raise MediaError("Expected a braced argument, such as {x + 1}.")
        self.index += 1
        return self.row(depth + 1, "}")

    def argument(self, depth):
        self.skip_space()
        if self.index >= len(self.source):
            raise MediaError("A subscript or superscript needs an argument.")
        return self.atom(depth + 1)

    def row(self, depth=0, closing=None):
        if depth > 32:
            raise MediaError("Equation nesting exceeds 32 levels.")
        children = []
        while self.index < len(self.source):
            self.skip_space()
            if self.index == len(self.source):
                break
            character = self.source[self.index]
            if closing and character == closing:
                self.index += 1
                return self.node("row", children=children)
            if character == "}":
                raise MediaError("Unmatched closing brace in equation.")
            node = self.atom(depth)
            sub = sup = None
            self.skip_space()
            while self.index < len(self.source) and self.source[self.index] in "_^":
                marker = self.source[self.index]
                self.index += 1
                value = self.argument(depth)
                if marker == "_":
                    if sub is not None:
                        raise MediaError("An element cannot have two subscripts.")
                    sub = value
                else:
                    if sup is not None:
                        raise MediaError("An element cannot have two superscripts.")
                    sup = value
                self.skip_space()
            if sub is not None or sup is not None:
                empty = self.node("row")
                node = self.node("scripts", children=(node, sub or empty, sup or empty))
            children.append(node)
        if closing:
            raise MediaError(f"Missing closing {closing} in equation.")
        return self.node("row", children=children)

    def atom(self, depth):
        if depth > 32:
            raise MediaError("Equation nesting exceeds 32 levels.")
        if self.index >= len(self.source):
            raise MediaError("Missing equation argument.")
        character = self.source[self.index]
        self.index += 1
        if character == "{":
            return self.row(depth + 1, "}")
        if character in "}_^$&#%":
            raise MediaError(f"Unexpected {character!r}; use braces for script arguments.")
        if character != "\\":
            return self.node("text", character)
        match = re.match(r"[A-Za-z]+|.", self.source[self.index:])
        if match is None:
            raise MediaError("Incomplete TeX command.")
        command = match.group()
        self.index += len(command)
        if command in _TEX_SYMBOLS:
            return self.node("text", _TEX_SYMBOLS[command])
        if command in ("sum", "prod", "int", "oint"):
            return self.node("operator", {"sum": "∑", "prod": "∏", "int": "∫", "oint": "∮"}[command])
        if command in ("frac", "dfrac", "tfrac"):
            return self.node("fraction", children=(self.group(depth), self.group(depth)))
        if command == "sqrt":
            self.skip_space()
            index = None
            if self.index < len(self.source) and self.source[self.index] == "[":
                self.index += 1
                index = self.row(depth + 1, "]")
            body = self.group(depth)
            return self.node("root", children=(body,) if index is None else (body, index))
        if command in ("text", "mathrm", "mathbf", "mathit"):
            if command == "text":
                self.skip_space()
                if self.index >= len(self.source) or self.source[self.index] != "{":
                    raise MediaError("The text command needs a braced argument.")
                end = self.source.find("}", self.index + 1)
                if end < 0 or "{" in self.source[self.index + 1:end]:
                    raise MediaError("Text arguments must have balanced, unnested braces.")
                value = self.source[self.index + 1:end]
                self.index = end + 1
                return self.node("text", value)
            return self.node("style", command, (self.group(depth),))
        if command in (",", ";", ":", " ", "quad", "qquad", "!"):
            return self.node("space", {"!": "0", ",": "0.15", ":": "0.25", ";": "0.35", "qquad": "2"}.get(command, "1"))
        if command in ("{", "}", "_", "^", "%", "#", "&", "$", "|"):
            return self.node("text", command)
        raise MediaError(f"Unsupported TeX command: \\{command}.")


def _safe_xml(source, label):
    if re.search(r"<!\s*(?:DOCTYPE|ENTITY)|<\?xml-stylesheet", source, re.IGNORECASE):
        raise MediaError(f"{label} cannot contain document types, entities, or external stylesheets.")
    try:
        return ET.fromstring(source)
    except ET.ParseError as error:
        raise MediaError(f"Invalid {label}: {error}.") from error


def _mathml_node(element, depth=0, budget=None):
    budget = [0] if budget is None else budget
    budget[0] += 1
    if depth > 32 or budget[0] > 2048:
        raise MediaError("MathML is too deeply nested or contains too many elements.")
    tag = element.tag.rsplit("}", 1)[-1]
    allowed_attributes = {"display", "xmlns"} if tag == "math" else set()
    if tag == "mfenced":
        allowed_attributes = {"open", "close", "separators"}
    for name in element.attrib:
        if name not in allowed_attributes:
            raise MediaError(f"Unsupported MathML attribute: {name}.")
    children = tuple(_mathml_node(child, depth + 1, budget) for child in element)
    if tag in ("mi", "mn", "mo", "mtext"):
        if children:
            raise MediaError(f"MathML {tag} must contain text only.")
        text = element.text or ""
        return EquationNode("operator" if tag == "mo" and text in "∑∏∫∮" else "text", text)
    if (element.text or "").strip() or any((child.tail or "").strip() for child in element):
        raise MediaError("MathML text must be inside mi, mn, mo, or mtext elements.")
    if tag in ("math", "mrow"):
        return EquationNode("row", children=children)
    if tag == "msqrt":
        if not children:
            raise MediaError("MathML msqrt needs an argument.")
        return EquationNode("root", children=(EquationNode("row", children=children),))
    counts = {"mfrac": 2, "mroot": 2, "msub": 2, "msup": 2, "msubsup": 3,
              "munder": 2, "mover": 2, "munderover": 3}
    if tag in counts:
        if len(children) != counts[tag]:
            raise MediaError(f"MathML {tag} needs {counts[tag]} arguments.")
        if tag in ("mfrac", "mroot"):
            return EquationNode("fraction" if tag == "mfrac" else "root", children=children)
        empty = EquationNode("row")
        if tag in ("msub", "munder"):
            slots = (children[0], children[1], empty)
        elif tag in ("msup", "mover"):
            slots = (children[0], empty, children[1])
        else:
            slots = children
        return EquationNode("limits" if tag in ("munder", "mover", "munderover") else "scripts", children=slots)
    if tag == "mfenced":
        opening, closing = element.get("open", "("), element.get("close", ")")
        separators = element.get("separators", ",").replace(" ", "")
        parts = [EquationNode("text", opening)]
        for index, child in enumerate(children):
            if index and separators:
                parts.append(EquationNode("text", separators[min(index - 1, len(separators) - 1)]))
            parts.append(child)
        parts.append(EquationNode("text", closing))
        return EquationNode("row", children=tuple(parts))
    raise MediaError(f"Unsupported MathML element: {tag}.")


def parse_equation(source: str, syntax: str = "tex") -> EquationNode:
    if not isinstance(source, str) or not source.strip():
        raise MediaError("Enter an equation before inserting it.")
    if len(source) > MAX_EQUATION_LENGTH:
        raise MediaError("Equation source exceeds 16,000 characters.")
    if syntax.lower() == "tex":
        return _TexParser(source).row()
    if syntax.lower() == "mathml":
        root = _safe_xml(source, "MathML")
        if root.tag.rsplit("}", 1)[-1] != "math":
            raise MediaError("MathML must have a math root element.")
        return _mathml_node(root)
    raise MediaError("Choose TeX or MathML equation syntax.")


@dataclass
class _EquationBox:
    width: float
    height: float
    baseline: float
    operations: list


def _place(operations, box, x, y):
    operations.extend((kind, data, px + x, py + y) for kind, data, px, py in box.operations)


def _equation_box(node, size, bold=False, italic=False):
    font = QFont("Cambria Math")
    font.setPointSizeF(size)
    font.setBold(bold)
    font.setItalic(italic)
    metrics = QFontMetricsF(font)
    em = metrics.height()
    gap = max(2.0, em * 0.12)
    if node.kind in ("text", "operator"):
        if node.kind == "operator":
            font.setPointSizeF(size * 1.35)
            metrics = QFontMetricsF(font)
        return _EquationBox(max(metrics.horizontalAdvance(node.value), metrics.boundingRect(node.value).width()),
                            metrics.height(), metrics.ascent(), [("text", (node.value, font), 0, metrics.ascent())])
    if node.kind == "space":
        return _EquationBox(em * float(node.value), 0, 0, [])
    if node.kind == "style":
        return _equation_box(node.children[0], size, node.value == "mathbf", node.value == "mathit")
    if node.kind == "row":
        boxes = [_equation_box(child, size, bold, italic) for child in node.children]
        baseline = max((box.baseline for box in boxes), default=0)
        descent = max((box.height - box.baseline for box in boxes), default=0)
        operations = []
        x = 0.0
        for box in boxes:
            _place(operations, box, x, baseline - box.baseline)
            x += box.width
        return _EquationBox(x, baseline + descent, baseline, operations)
    if node.kind == "fraction":
        top, bottom = [_equation_box(child, max(8, size * 0.86), bold, italic) for child in node.children]
        width = max(top.width, bottom.width) + gap * 2
        line_y = top.height + gap
        bottom_y = line_y + gap + 1
        operations = [("line", (width, 0), 0, line_y)]
        _place(operations, top, (width - top.width) / 2, 0)
        _place(operations, bottom, (width - bottom.width) / 2, bottom_y)
        return _EquationBox(width, bottom_y + bottom.height, line_y + em * 0.23, operations)
    if node.kind == "root":
        body = _equation_box(node.children[0], size, bold, italic)
        index = _equation_box(node.children[1], max(8, size * 0.55), bold, italic) if len(node.children) > 1 else None
        root_width = em * 0.55
        offset = max(0, (index.width - root_width * 0.35)) if index else 0
        top = max(gap, index.height * 0.5 if index else 0)
        body_x = offset + root_width
        height = top + body.height
        operations = [
            ("line", (root_width * 0.18, -gap), offset, top + body.height * 0.55),
            ("line", (root_width * 0.2, body.height * 0.4 + gap), offset + root_width * 0.18, top + body.height * 0.55 - gap),
            ("line", (root_width * 0.27, -body.height * 0.95 - gap), offset + root_width * 0.38, top + body.height * 0.95),
            ("line", (body.width + root_width * 0.35 + gap, 0), offset + root_width * 0.65, top - gap),
        ]
        _place(operations, body, body_x, top)
        if index:
            _place(operations, index, 0, 0)
        return _EquationBox(body_x + body.width + gap, height, top + body.baseline, operations)
    if node.kind in ("scripts", "limits"):
        body = _equation_box(node.children[0], size, bold, italic)
        sub = _equation_box(node.children[1], max(8, size * 0.65), bold, italic)
        sup = _equation_box(node.children[2], max(8, size * 0.65), bold, italic)
        operations = []
        if node.kind == "limits":
            width = max(body.width, sub.width, sup.width)
            top = sup.height + gap if sup.height else 0
            _place(operations, sup, (width - sup.width) / 2, 0)
            _place(operations, body, (width - body.width) / 2, top)
            _place(operations, sub, (width - sub.width) / 2, top + body.height + gap)
            return _EquationBox(width, top + body.height + (sub.height + gap if sub.height else 0), top + body.baseline, operations)
        top = max(0, sup.height - body.baseline * 0.5)
        sub_y = top + body.baseline - sub.baseline * 0.2
        _place(operations, body, 0, top)
        _place(operations, sup, body.width + 1, 0)
        _place(operations, sub, body.width + 1, sub_y)
        return _EquationBox(body.width + max(sub.width, sup.width) + 1,
                            max(top + body.height, sub_y + sub.height), top + body.baseline, operations)
    raise MediaError(f"Unsupported equation node: {node.kind}.")


def render_equation(source: str, syntax: str = "tex", point_size: float = 24,
                    color: QColor | str = "#182638") -> QImage:
    if not math.isfinite(point_size) or not 8 <= point_size <= 96:
        raise MediaError("Equation size must be between 8 and 96 points.")
    box = _equation_box(parse_equation(source, syntax), point_size)
    width, height = math.ceil(box.width + 16), math.ceil(box.height + 16)
    if width > 4096 or height > 4096 or width * height > 8000000:
        raise MediaError("Rendered equation is too large. Split it into smaller equations.")
    result = QImage(width, height, QImage.Format.Format_ARGB32_Premultiplied)
    result.fill(Qt.GlobalColor.transparent)
    painter = QPainter(result)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
    painter.setPen(QPen(QColor(color), max(1.0, point_size / 20)))
    for kind, data, x, y in box.operations:
        if kind == "text":
            text, font = data
            painter.setFont(font)
            painter.drawText(QPointF(x + 8, y + 8), text)
        else:
            painter.drawLine(QPointF(x + 8, y + 8), QPointF(x + data[0] + 8, y + data[1] + 8))
    painter.end()
    return result


class EquationDialog(QDialog):
    def __init__(self, parent=None, source="", syntax="tex"):
        super().__init__(parent)
        self.setWindowTitle("Equation Editor")
        self.resize(640, 460)
        self._image = None
        layout = QVBoxLayout(self)
        self.syntax_combo = QComboBox()
        self.syntax_combo.addItem("TeX", "tex")
        self.syntax_combo.addItem("MathML", "mathml")
        self.syntax_combo.setCurrentIndex(1 if syntax == "mathml" else 0)
        layout.addWidget(self.syntax_combo)
        self.source_edit = QPlainTextEdit(source or r"\frac{x^{2} + 1}{\sqrt{y}}")
        self.source_edit.setAccessibleName("Equation source")
        layout.addWidget(self.source_edit)
        hint = QLabel("Supports fractions, roots, scripts, sums, integrals, Greek letters, and grouped expressions. Unsupported commands show an error.")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        self.preview = QLabel()
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.preview)
        layout.addWidget(scroll)
        self.error_label = QLabel()
        self.error_label.setWordWrap(True)
        self.error_label.setStyleSheet("color: #a12727;")
        layout.addWidget(self.error_label)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        layout.addWidget(self.buttons)
        self.source_edit.textChanged.connect(self.refresh_preview)
        self.syntax_combo.currentIndexChanged.connect(self.refresh_preview)
        self.refresh_preview()

    def source(self):
        return self.source_edit.toPlainText()

    def syntax(self):
        return self.syntax_combo.currentData()

    def rendered_image(self):
        if self._image is None:
            raise MediaError(self.error_label.text() or "Equation is invalid.")
        return self._image.copy()

    def refresh_preview(self):
        try:
            self._image = render_equation(self.source(), self.syntax())
        except (MediaError, ValueError) as error:
            self._image = None
            self.preview.clear()
            self.error_label.setText(str(error))
        else:
            self.preview.setPixmap(QPixmap.fromImage(self._image))
            self.error_label.clear()
        self.buttons.button(QDialogButtonBox.StandardButton.Ok).setEnabled(self._image is not None)

    def accept(self):
        self.refresh_preview()
        if self._image is not None:
            super().accept()


def render_svg(data: bytes, max_size: QSize = QSize(1024, 1024)) -> QImage:
    if not isinstance(data, bytes) or len(data) > 4 * 1024 * 1024:
        raise MediaError("SVG input must be no larger than 4 MiB.")
    try:
        source = data.decode("utf-8-sig")
    except UnicodeDecodeError as error:
        raise MediaError("SVG must use UTF-8 encoding.") from error
    root = _safe_xml(source, "SVG")
    if root.tag.rsplit("}", 1)[-1] != "svg":
        raise MediaError("The selected file does not have an SVG root element.")
    elements = list(root.iter())
    if len(elements) > 10000:
        raise MediaError("SVG contains too many drawing elements.")
    for element in elements:
        tag = element.tag.rsplit("}", 1)[-1].lower()
        if tag in ("script", "foreignobject", "animate", "animatetransform", "animatemotion", "set"):
            raise MediaError("Only static SVG drawings are supported.")
        for raw_name, value in element.attrib.items():
            name = raw_name.rsplit("}", 1)[-1].lower()
            if name.startswith("on") or name == "base":
                raise MediaError("SVG event handlers and external bases are not supported.")
            if name in ("href", "src") and not value.startswith("#"):
                if tag != "image" or not re.fullmatch(r"data:image/(?:png|jpeg);base64,[A-Za-z0-9+/=\s]+", value):
                    raise MediaError("SVG may only reference internal shapes or embedded PNG/JPEG images.")
            references = re.findall(r"url\s*\(([^)]*)\)", value, re.IGNORECASE)
            if any(not reference.strip().strip("\"'").startswith("#") for reference in references) or "\\" in value:
                raise MediaError("SVG external resources are not supported.")
        if tag == "style":
            styles = element.text or ""
            references = re.findall(r"url\s*\(([^)]*)\)", styles, re.IGNORECASE)
            if "@" in styles or "\\" in styles or any(not reference.strip().strip("\"'").startswith("#") for reference in references):
                raise MediaError("SVG styles cannot load external resources.")
    source = re.sub(r"url\(\s*(['\"])(#[^)'\"]*)\1\s*\)", r"url(\2)", source)
    renderer = QSvgRenderer(source.encode("utf-8"))
    if not renderer.isValid():
        raise MediaError("The SVG drawing could not be rendered.")
    bounds = renderer.viewBoxF()
    if not bounds.isValid() or not all(math.isfinite(value) for value in (bounds.x(), bounds.y(), bounds.width(), bounds.height())):
        raise MediaError("SVG needs a finite, nonempty drawing area.")
    limit = QSize(max(1, min(2048, max_size.width())), max(1, min(2048, max_size.height())))
    size = bounds.size().toSize().scaled(limit, Qt.AspectRatioMode.KeepAspectRatio)
    size = QSize(max(1, size.width()), max(1, size.height()))
    image = QImage(size, QImage.Format.Format_ARGB32_Premultiplied)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    renderer.render(painter, QRectF(0, 0, size.width(), size.height()))
    painter.end()
    return image


@dataclass(frozen=True)
class ModelScene:
    vertices: tuple[tuple[float, float, float], ...]
    triangles: tuple[tuple[int, int, int], ...]
    colors: tuple[tuple[float, float, float, float], ...]

    def to_dict(self):
        return {"vertices": [list(vertex) for vertex in self.vertices],
                "triangles": [list(face) for face in self.triangles],
                "colors": [list(color) for color in self.colors]}

    @classmethod
    def from_dict(cls, value):
        try:
            vertices = tuple(tuple(float(item) for item in vertex) for vertex in value["vertices"])
            triangles = tuple(tuple(item for item in face) for face in value["triangles"])
            colors = tuple(tuple(float(item) for item in color) for color in value["colors"])
        except (KeyError, TypeError, ValueError, OverflowError) as error:
            raise MediaError("Invalid embedded 3D model.") from error
        if not 1 <= len(vertices) <= MAX_MODEL_VERTICES or not 1 <= len(triangles) <= MAX_MODEL_TRIANGLES:
            raise MediaError("3D model is empty or exceeds the geometry limit.")
        if any(len(vertex) != 3 or any(not math.isfinite(item) or abs(item) > 1e12 for item in vertex) for vertex in vertices):
            raise MediaError("3D model has invalid vertex coordinates.")
        if any(len(face) != 3 or any(type(item) is not int or not 0 <= item < len(vertices) for item in face) for face in triangles):
            raise MediaError("3D model has invalid triangle indices.")
        if len(colors) != len(triangles) or any(len(color) != 4 or any(not math.isfinite(item) or not 0 <= item <= 1 for item in color) for color in colors):
            raise MediaError("3D model has invalid face colors.")
        return cls(vertices, triangles, colors)


def _bounded_file(path, limit=MAX_MEDIA_BYTES):
    try:
        if path.stat().st_size > limit:
            raise MediaError(f"Media file exceeds the {limit // (1024 * 1024)} MiB size limit.")
        with path.open("rb") as stream:
            data = stream.read(limit + 1)
        if len(data) > limit:
            raise MediaError("Media file exceeds the size limit.")
        return data
    except OSError as error:
        raise MediaError(f"Cannot read model file: {error}.") from error


def _model_json(path):
    raw = _bounded_file(path)
    binary = None
    if path.suffix.lower() == ".glb":
        if len(raw) < 20 or raw[:4] != b"glTF":
            raise MediaError("Invalid GLB header.")
        _, version, length = struct.unpack_from("<4sII", raw)
        if version != 2 or length != len(raw):
            raise MediaError("Only complete GLB version 2 files are supported.")
        offset = 12
        json_chunk = None
        while offset < len(raw):
            if offset + 8 > len(raw):
                raise MediaError("Truncated GLB chunk header.")
            length, kind = struct.unpack_from("<II", raw, offset)
            offset += 8
            if length % 4 or offset + length > len(raw):
                raise MediaError("Invalid GLB chunk length.")
            chunk = raw[offset:offset + length]
            offset += length
            if kind == 0x4E4F534A:
                if json_chunk is not None:
                    raise MediaError("GLB contains multiple JSON chunks.")
                json_chunk = chunk
            elif kind == 0x004E4942:
                if binary is not None or json_chunk is None:
                    raise MediaError("Invalid GLB binary chunk order.")
                binary = chunk
        if json_chunk is None:
            raise MediaError("GLB has no JSON scene.")
        raw = json_chunk
    elif path.suffix.lower() != ".gltf":
        raise MediaError("Select a .gltf or .glb model.")
    try:
        model = json.loads(raw.decode("utf-8-sig"))
    except (UnicodeDecodeError, ValueError, RecursionError) as error:
        raise MediaError("Invalid model JSON.") from error
    if not isinstance(model, dict) or not isinstance(model.get("asset"), dict) or model["asset"].get("version") != "2.0":
        raise MediaError("Only glTF 2.0 models are supported.")
    for key in ("buffers", "bufferViews", "accessors", "nodes", "meshes", "scenes", "materials"):
        if not isinstance(model.get(key, []), list) or any(not isinstance(item, dict) for item in model.get(key, [])):
            raise MediaError(f"Invalid model {key} collection.")
    if model.get("extensionsRequired"):
        raise MediaError("Models requiring glTF extensions are not supported.")
    if model.get("skins") or model.get("animations"):
        raise MediaError("Only static models without skinning or animation are supported.")
    buffers = []
    total = 0
    for index, descriptor in enumerate(model.get("buffers", [])):
        uri = descriptor.get("uri")
        if uri is not None and not isinstance(uri, str):
            raise MediaError("Invalid model buffer URI.")
        if uri is None:
            if index != 0 or binary is None:
                raise MediaError("Missing model buffer data.")
            payload = binary
        elif uri.startswith("data:"):
            match = re.fullmatch(r"data:(?:application/octet-stream|application/gltf-buffer);base64,([A-Za-z0-9+/=]+)", uri)
            if match is None:
                raise MediaError("Only embedded base64 binary buffers are supported.")
            try:
                payload = base64.b64decode(match.group(1), validate=True)
            except (ValueError, binascii.Error) as error:
                raise MediaError("Invalid embedded model buffer.") from error
        else:
            parts = urlsplit(uri)
            local_path = Path(unquote(parts.path))
            if parts.scheme or parts.netloc or parts.query or parts.fragment or local_path.is_absolute() or "\\" in uri or ":" in uri:
                raise MediaError("Model buffers must be local files inside the model folder.")
            base = path.parent.resolve()
            target = (base / local_path).resolve()
            if not target.is_relative_to(base):
                raise MediaError("Model buffer path escapes the model folder.")
            payload = _bounded_file(target)
        expected = descriptor.get("byteLength")
        if type(expected) is not int or expected < 0 or expected > len(payload):
            raise MediaError("Model buffer has an invalid byte length.")
        total += len(payload)
        if total > MAX_MEDIA_BYTES:
            raise MediaError("Combined model buffers exceed 32 MiB.")
        buffers.append(payload[:expected])
    return model, buffers


def _accessor(model, buffers, index, positions=False):
    try:
        if type(index) is not int or not 0 <= index < len(model.get("accessors", [])):
            raise MediaError("Invalid model accessor reference.")
        accessor = model["accessors"][index]
        if "sparse" in accessor or accessor.get("normalized"):
            raise MediaError("Sparse and normalized model accessors are not supported.")
        expected = "VEC3" if positions else "SCALAR"
        component = accessor["componentType"]
        if accessor["type"] != expected or component not in ((5126,) if positions else (5121, 5123, 5125)):
            raise MediaError("Model requires float VEC3 positions and unsigned integer indices.")
        count = accessor["count"]
        limit = MAX_MODEL_VERTICES if positions else MAX_MODEL_TRIANGLES * 3
        if type(count) is not int or not 1 <= count <= limit:
            raise MediaError("Model accessor exceeds the geometry limit.")
        view_index = accessor["bufferView"]
        if type(view_index) is not int or not 0 <= view_index < len(model.get("bufferViews", [])):
            raise MediaError("Invalid model buffer view reference.")
        view = model["bufferViews"][view_index]
        buffer_index = view["buffer"]
        if type(buffer_index) is not int or not 0 <= buffer_index < len(buffers):
            raise MediaError("Invalid model buffer reference.")
        data = buffers[buffer_index]
        code = {5121: "B", 5123: "H", 5125: "I", 5126: "f"}[component]
        format_string = "<" + code * (3 if positions else 1)
        width = struct.calcsize(format_string)
        stride = view.get("byteStride", width)
        view_start, view_length = view.get("byteOffset", 0), view["byteLength"]
        start = accessor.get("byteOffset", 0)
        if any(type(item) is not int for item in (stride, view_start, view_length, start)):
            raise MediaError("Invalid model buffer offsets.")
        if stride < width or stride > 252 or min(view_start, view_length, start) < 0 or view_start + view_length > len(data) or start + (count - 1) * stride + width > view_length:
            raise MediaError("Model accessor extends beyond its buffer.")
        return [struct.unpack_from(format_string, data, view_start + start + row * stride) for row in range(count)]
    except (KeyError, IndexError, TypeError, struct.error) as error:
        raise MediaError("Invalid model accessor or buffer reference.") from error


def _matrix_multiply(left, right):
    return tuple(sum(left[row + k * 4] * right[k + column * 4] for k in range(4))
                 for column in range(4) for row in range(4))


_IDENTITY = (1., 0., 0., 0., 0., 1., 0., 0., 0., 0., 1., 0., 0., 0., 0., 1.)


def _node_matrix(node):
    try:
        if "matrix" in node:
            if any(key in node for key in ("translation", "rotation", "scale")):
                raise MediaError("A model node cannot mix matrix and TRS transforms.")
            values = tuple(float(value) for value in node["matrix"])
            if len(values) != 16 or not all(math.isfinite(value) for value in values):
                raise MediaError("Invalid model transform matrix.")
            if values[3] != 0 or values[7] != 0 or values[11] != 0 or values[15] != 1:
                raise MediaError("Model transforms must be affine.")
            return values
        tx, ty, tz = (float(value) for value in node.get("translation", (0, 0, 0)))
        sx, sy, sz = (float(value) for value in node.get("scale", (1, 1, 1)))
        x, y, z, w = (float(value) for value in node.get("rotation", (0, 0, 0, 1)))
        if not all(math.isfinite(value) for value in (tx, ty, tz, sx, sy, sz, x, y, z, w)):
            raise MediaError("Model transforms must be finite.")
        norm = math.sqrt(x*x + y*y + z*z + w*w)
        if norm < 1e-8:
            raise MediaError("Model rotation quaternion is invalid.")
        x, y, z, w = x/norm, y/norm, z/norm, w/norm
        return ((1 - 2*y*y - 2*z*z)*sx, (2*x*y + 2*z*w)*sx, (2*x*z - 2*y*w)*sx, 0,
                (2*x*y - 2*z*w)*sy, (1 - 2*x*x - 2*z*z)*sy, (2*y*z + 2*x*w)*sy, 0,
                (2*x*z + 2*y*w)*sz, (2*y*z - 2*x*w)*sz, (1 - 2*x*x - 2*y*y)*sz, 0,
                tx, ty, tz, 1)
    except (TypeError, ValueError, OverflowError) as error:
        raise MediaError("Invalid model node transform.") from error


def load_model(path: str | Path) -> ModelScene:
    path = Path(path)
    try:
        model, buffers = _model_json(path)
        vertices, triangles, colors = [], [], []
        meshes = model.get("meshes", [])

        def add_mesh(mesh_index, matrix):
            if type(mesh_index) is not int or not 0 <= mesh_index < len(meshes):
                raise MediaError("Invalid model mesh reference.")
            for primitive in meshes[mesh_index].get("primitives", []):
                if not isinstance(primitive, dict):
                    raise MediaError("Invalid model mesh primitive.")
                if primitive.get("mode", 4) != 4 or primitive.get("targets") or primitive.get("extensions"):
                    raise MediaError("Only uncompressed static triangle meshes are supported.")
                material_index = primitive.get("material")
                if material_index is not None and (type(material_index) is not int or not 0 <= material_index < len(model.get("materials", []))):
                    raise MediaError("Invalid model material reference.")
                material = model["materials"][material_index] if material_index is not None else {}
                pbr = material.get("pbrMetallicRoughness", {})
                if any("Texture" in key for key in pbr) or any("Texture" in key for key in material) or "COLOR_0" in primitive.get("attributes", {}):
                    raise MediaError("Textured and vertex-colored models are not supported; use solid material colors.")
                color = tuple(pbr.get("baseColorFactor", (0.32, 0.52, 0.78, 1.0)))
                source = _accessor(model, buffers, primitive["attributes"]["POSITION"], positions=True)
                indices = [item[0] for item in _accessor(model, buffers, primitive["indices"])] if "indices" in primitive else list(range(len(source)))
                if len(indices) % 3 or any(index >= len(source) for index in indices):
                    raise MediaError("Model triangle indices are invalid.")
                base = len(vertices)
                if base + len(source) > MAX_MODEL_VERTICES or len(triangles) + len(indices) // 3 > MAX_MODEL_TRIANGLES:
                    raise MediaError("Model exceeds 100,000 vertices or triangles.")
                for x, y, z in source:
                    vertices.append(tuple(matrix[row] * x + matrix[4 + row] * y + matrix[8 + row] * z + matrix[12 + row] for row in range(3)))
                for index in range(0, len(indices), 3):
                    triangles.append(tuple(base + item for item in indices[index:index + 3]))
                    colors.append(color)

        nodes = model.get("nodes", [])
        visited = set()

        def visit(index, matrix, ancestry):
            if type(index) is not int or not 0 <= index < len(nodes) or index in ancestry or len(ancestry) > 64:
                raise MediaError("Invalid or cyclic model node hierarchy.")
            if index in visited:
                raise MediaError("Model node is referenced by more than one parent.")
            visited.add(index)
            node = nodes[index]
            transform = _matrix_multiply(matrix, _node_matrix(node))
            if "mesh" in node:
                add_mesh(node["mesh"], transform)
            for child in node.get("children", []):
                visit(child, transform, ancestry | {index})

        if nodes:
            scenes = model.get("scenes", [])
            if scenes:
                scene_index = model.get("scene", 0)
                if type(scene_index) is not int or not 0 <= scene_index < len(scenes):
                    raise MediaError("Invalid default model scene reference.")
                roots = scenes[scene_index].get("nodes", [])
            else:
                children = {child for node in nodes for child in node.get("children", [])}
                roots = [index for index in range(len(nodes)) if index not in children]
            for index in roots:
                visit(index, _IDENTITY, set())
        else:
            for index in range(len(meshes)):
                add_mesh(index, _IDENTITY)
        return ModelScene.from_dict({"vertices": vertices, "triangles": triangles, "colors": colors})
    except MediaError:
        raise
    except (ValueError, TypeError, IndexError, KeyError, AttributeError, OverflowError, RecursionError) as error:
        raise MediaError("The model contains invalid or unsupported geometry.") from error


def render_model(scene: ModelScene, size: QSize = QSize(640, 480)) -> QImage:
    scene = ModelScene.from_dict(scene.to_dict())
    width, height = max(64, min(2048, size.width())), max(64, min(2048, size.height()))
    yaw, pitch = math.radians(-28), math.radians(18)
    projected = []
    for x, y, z in scene.vertices:
        x, z = x * math.cos(yaw) + z * math.sin(yaw), -x * math.sin(yaw) + z * math.cos(yaw)
        y, z = y * math.cos(pitch) - z * math.sin(pitch), y * math.sin(pitch) + z * math.cos(pitch)
        projected.append((x, -y, z))
    minimum = [min(vertex[axis] for vertex in projected) for axis in range(2)]
    maximum = [max(vertex[axis] for vertex in projected) for axis in range(2)]
    scale = min((width - 48) / max(maximum[0] - minimum[0], 1e-8), (height - 48) / max(maximum[1] - minimum[1], 1e-8))
    center = [(minimum[axis] + maximum[axis]) / 2 for axis in range(2)]
    screen = [QPointF((vertex[0] - center[0]) * scale + width/2, (vertex[1] - center[1]) * scale + height/2) for vertex in projected]
    image = QImage(width, height, QImage.Format.Format_ARGB32_Premultiplied)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    order = sorted(range(len(scene.triangles)), key=lambda index: sum(projected[item][2] for item in scene.triangles[index]))
    for index in order:
        face = scene.triangles[index]
        a, b, c = [projected[item] for item in face]
        u, v = [b[axis] - a[axis] for axis in range(3)], [c[axis] - a[axis] for axis in range(3)]
        normal = (u[1]*v[2] - u[2]*v[1], u[2]*v[0] - u[0]*v[2], u[0]*v[1] - u[1]*v[0])
        length = math.sqrt(sum(value*value for value in normal))
        if length < 1e-12:
            continue
        light = 0.45 + 0.55 * abs((-0.3*normal[0] - 0.4*normal[1] + 0.866*normal[2]) / length)
        red, green, blue, alpha = scene.colors[index]
        color = QColor.fromRgbF(min(1, red*light), min(1, green*light), min(1, blue*light), alpha)
        painter.setBrush(color)
        painter.setPen(QPen(color.darker(115), 0.65))
        painter.drawPolygon(QPolygonF([screen[item] for item in face]))
    painter.end()
    return image
