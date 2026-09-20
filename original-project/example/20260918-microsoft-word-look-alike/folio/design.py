from __future__ import annotations

import base64
import math
import re
from dataclasses import asdict, dataclass, fields

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import (
    QColor, QFont, QImage, QPainter, QPen, QTextCharFormat, QTextCursor,
    QTextFormat,
)


STYLE_PROPERTY = int(QTextFormat.Property.UserProperty) + 810
BASELINE_PROPERTY = int(QTextFormat.Property.UserProperty) + 811
PT_TO_PX = 96.0 / 72.0
MAX_WATERMARK_BYTES = 8 * 1024 * 1024

PALETTES = {
    "Office Blue": {"text": "#263244", "muted": "#64748b",
                    "accents": ("#17365d", "#245c9f", "#4472c4", "#70ad47")},
    "Forest": {"text": "#24362c", "muted": "#65786b",
               "accents": ("#166534", "#22844d", "#65a30d", "#ca8a04")},
    "Plum": {"text": "#36243e", "muted": "#82718c",
             "accents": ("#6b21a8", "#9333ea", "#be185d", "#7c3aed")},
    "Monochrome": {"text": "#333333", "muted": "#777777",
                   "accents": ("#222222", "#555555", "#888888", "#aaaaaa")},
}
FONT_PAIRINGS = {
    "Calibri / Calibri": ("Calibri", "Calibri"),
    "Cambria / Calibri": ("Cambria", "Calibri"),
    "Georgia / Georgia": ("Georgia", "Georgia"),
    "Segoe UI / Segoe UI": ("Segoe UI", "Segoe UI"),
}
THEMES = {
    "Office": ("Office Blue", "Calibri / Calibri", 0.86),
    "Woodland": ("Forest", "Cambria / Calibri", 0.88),
    "Editorial": ("Plum", "Georgia / Georgia", 0.91),
    "Minimal": ("Monochrome", "Segoe UI / Segoe UI", 0.93),
}


@dataclass(frozen=True)
class SemanticStyle:
    size_pt: float
    before_pt: float
    after_pt: float
    font_token: str = "body"
    color_token: str = "text"
    bold: bool = False
    italic: bool = False
    heading_level: int = 0


STYLE_SETS = {
    "Modern": {
        "Normal": SemanticStyle(11, 0, 8),
        "Title": SemanticStyle(32, 0, 18, "heading", "accent1"),
        "Subtitle": SemanticStyle(16, 0, 16, "body", "muted"),
        "Heading 1": SemanticStyle(22, 20, 10, "heading", "accent1", True, False, 1),
        "Heading 2": SemanticStyle(16, 16, 8, "heading", "accent2", True, False, 2),
        "Heading 3": SemanticStyle(12, 12, 6, "heading", "accent2", True, False, 3),
        "Quote": SemanticStyle(12, 10, 10, "body", "muted", False, True),
    },
    "Compact": {
        "Normal": SemanticStyle(10, 0, 4),
        "Title": SemanticStyle(26, 0, 12, "heading", "accent1"),
        "Subtitle": SemanticStyle(13, 0, 8, "body", "muted"),
        "Heading 1": SemanticStyle(18, 12, 6, "heading", "accent1", True, False, 1),
        "Heading 2": SemanticStyle(14, 10, 4, "heading", "accent2", True, False, 2),
        "Heading 3": SemanticStyle(12, 8, 4, "heading", "accent2", True, False, 3),
        "Quote": SemanticStyle(10, 6, 6, "body", "muted", False, True),
    },
    "Large Print": {
        "Normal": SemanticStyle(16, 0, 12),
        "Title": SemanticStyle(38, 0, 24, "heading", "accent1"),
        "Subtitle": SemanticStyle(22, 0, 20, "body", "muted"),
        "Heading 1": SemanticStyle(28, 26, 16, "heading", "accent1", True, False, 1),
        "Heading 2": SemanticStyle(22, 20, 12, "heading", "accent2", True, False, 2),
        "Heading 3": SemanticStyle(18, 16, 10, "heading", "accent2", True, False, 3),
        "Quote": SemanticStyle(17, 14, 14, "body", "muted", False, True),
    },
}
SPACING_POLICIES = {"Default": None, "None": (0, 0), "Compact": (0, 4),
                    "Relaxed": (0, 12), "Double": (0, 20)}


def _number(value, low, high) -> bool:
    return (isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(value) and low <= value <= high)


def _color(value) -> bool:
    return isinstance(value, str) and bool(re.fullmatch(r"#[0-9a-fA-F]{6}", value))


def watermark_image(data_uri: str) -> QImage:
    if not isinstance(data_uri, str) or len(data_uri) > MAX_WATERMARK_BYTES * 2:
        raise ValueError("The watermark image is too large")
    match = re.fullmatch(r"data:image/(?:png|jpeg|jpg|webp);base64,([A-Za-z0-9+/=\r\n]+)",
                         data_uri)
    if not match:
        raise ValueError("Watermarks require an embedded PNG, JPEG, or WebP image")
    try:
        payload = base64.b64decode(re.sub(r"\s", "", match.group(1)), validate=True)
    except ValueError as exc:
        raise ValueError("Invalid watermark image encoding") from exc
    if len(payload) > MAX_WATERMARK_BYTES:
        raise ValueError("The watermark image is too large")
    image = QImage.fromData(payload)
    if image.isNull() or image.width() * image.height() > 40_000_000:
        raise ValueError("The watermark image is invalid or too large")
    return image


@dataclass
class DesignSettings:
    theme: str = "Office"
    color_palette: str = ""
    font_pairing: str = ""
    style_set: str = "Modern"
    paragraph_spacing: str = "Default"
    page_color: str = "#ffffff"
    page_border: str = "none"
    page_border_color: str = "#64748b"
    page_border_width_pt: float = 1.0
    page_border_inset_mm: float = 8.0
    watermark_kind: str = "none"
    watermark_text: str = ""
    watermark_image: str = ""
    watermark_color: str = "#64748b"
    watermark_opacity: float = 0.12
    watermark_angle: float = -35.0
    watermark_layer: str = "behind"

    def validate(self) -> None:
        for field in fields(self):
            if field.type == "str" and not isinstance(getattr(self, field.name), str):
                raise ValueError(f"{field.name} must be text")
        if self.theme not in THEMES:
            raise ValueError("Unknown document theme")
        if self.color_palette and self.color_palette not in PALETTES:
            raise ValueError("Unknown theme color palette")
        if self.font_pairing and self.font_pairing not in FONT_PAIRINGS:
            raise ValueError("Unknown theme font pairing")
        if self.style_set not in STYLE_SETS or self.paragraph_spacing not in SPACING_POLICIES:
            raise ValueError("Unknown style set or paragraph spacing policy")
        for name in ("page_color", "page_border_color", "watermark_color"):
            if not _color(getattr(self, name)):
                raise ValueError(f"{name} must be a six-digit hexadecimal color")
        if self.page_border not in ("none", "box", "double", "dashed"):
            raise ValueError("Unknown page border style")
        for name, low, high in (("page_border_width_pt", 0.25, 12),
                                ("page_border_inset_mm", 0, 50),
                                ("watermark_opacity", 0, 1),
                                ("watermark_angle", -180, 180)):
            if not _number(getattr(self, name), low, high):
                raise ValueError(f"Invalid {name}")
        if self.watermark_kind not in ("none", "text", "image"):
            raise ValueError("Unknown watermark type")
        if self.watermark_layer not in ("behind", "front"):
            raise ValueError("Unknown watermark layer")
        if not isinstance(self.watermark_text, str) or len(self.watermark_text) > 500:
            raise ValueError("Watermark text must contain at most 500 characters")
        if not isinstance(self.watermark_image, str):
            raise ValueError("Watermark image must be an embedded data URI")
        if self.watermark_image:
            watermark_image(self.watermark_image)
        if self.watermark_kind == "image" and not self.watermark_image:
            raise ValueError("Select a watermark image")

    def to_dict(self) -> dict:
        self.validate()
        return asdict(self)

    @classmethod
    def from_dict(cls, data: object) -> DesignSettings:
        if not isinstance(data, dict):
            raise ValueError("Design settings must be an object")
        allowed = {field.name for field in fields(cls)}
        if any(not isinstance(key, str) for key in data):
            raise ValueError("Invalid design setting name")
        instance = cls(**{key: value for key, value in data.items() if key in allowed})
        instance.validate()
        return instance


def theme_tokens(settings: DesignSettings) -> dict:
    settings.validate()
    palette_name, pairing_name, tint = THEMES[settings.theme]
    palette = PALETTES[settings.color_palette or palette_name]
    heading, body = FONT_PAIRINGS[settings.font_pairing or pairing_name]
    return {"heading_font": heading, "body_font": body, "text": palette["text"],
            "muted": palette["muted"], "accents": palette["accents"], "shape_tint": tint}


def theme_shape_colors(settings: DesignSettings, accent: int = 0) -> dict[str, str]:
    if not isinstance(accent, int) or isinstance(accent, bool) or accent < 0:
        raise ValueError("Accent index must be a non-negative integer")
    tokens = theme_tokens(settings)
    stroke = QColor(tokens["accents"][accent % len(tokens["accents"])])
    tint = tokens["shape_tint"]
    fill = QColor(*(round(channel + (255 - channel) * tint)
                    for channel in (stroke.red(), stroke.green(), stroke.blue())))
    return {"fill": fill.name(), "stroke": stroke.name(), "text": tokens["text"]}


def resolved_style(name: str, settings: DesignSettings) -> dict:
    tokens = theme_tokens(settings)
    if name not in STYLE_SETS[settings.style_set]:
        raise ValueError(f"Unknown paragraph style: {name}")
    style = STYLE_SETS[settings.style_set][name]
    color = (tokens["accents"][int(style.color_token[-1]) - 1]
             if style.color_token.startswith("accent") else tokens[style.color_token])
    spacing = SPACING_POLICIES[settings.paragraph_spacing]
    before, after = spacing if spacing is not None else (style.before_pt, style.after_pt)
    return {"font": tokens[style.font_token + "_font"], "size": style.size_pt,
            "weight": int(QFont.Weight.Bold if style.bold else QFont.Weight.Normal),
            "italic": style.italic, "color": color, "before": before * PT_TO_PX,
            "after": after * PT_TO_PX, "heading": style.heading_level}


def style_name(block) -> str:
    name = block.blockFormat().property(STYLE_PROPERTY)
    if name in STYLE_SETS["Modern"]:
        return name
    level = block.blockFormat().headingLevel()
    return f"Heading {min(level, 3)}" if level > 0 else "Normal"


def _inline_values(fmt: QTextCharFormat) -> dict:
    return {"font": fmt.fontFamilies()[0] if fmt.fontFamilies() else None,
            "size": fmt.fontPointSize() or None,
            "weight": fmt.fontWeight() if fmt.hasProperty(QTextFormat.Property.FontWeight) else None,
            "italic": fmt.fontItalic() if fmt.hasProperty(QTextFormat.Property.FontItalic) else None,
            "color": fmt.foreground().color().name()
            if fmt.hasProperty(QTextFormat.Property.ForegroundBrush) else None}


def _inline_format(values: dict) -> QTextCharFormat:
    fmt = QTextCharFormat()
    if "font" in values:
        fmt.setFontFamilies([values["font"]])
    if "size" in values:
        fmt.setFontPointSize(values["size"])
    if "weight" in values:
        fmt.setFontWeight(values["weight"])
    if "italic" in values:
        fmt.setFontItalic(values["italic"])
    if "color" in values:
        fmt.setForeground(QColor(values["color"]))
    return fmt


def _update_block(block, name: str, target: dict, baseline: dict, force: bool) -> None:
    cursor = QTextCursor(block)
    fmt = block.blockFormat()
    fmt.setProperty(STYLE_PROPERTY, name)
    fmt.setProperty(BASELINE_PROPERTY, target)
    fmt.setHeadingLevel(target["heading"])
    if force or abs(fmt.topMargin() - baseline["before"]) < 0.01:
        fmt.setTopMargin(target["before"])
    if force or abs(fmt.bottomMargin() - baseline["after"]) < 0.01:
        fmt.setBottomMargin(target["after"])
    if force:
        fmt.setLeftMargin(24.0 if name == "Quote" else 0.0)
        fmt.setRightMargin(0.0)
        fmt.setTextIndent(0.0)
        fmt.setIndent(0)
    cursor.setBlockFormat(fmt)
    fragments = []
    iterator = block.begin()
    while not iterator.atEnd():
        fragment = iterator.fragment()
        if fragment.isValid() and not fragment.charFormat().isImageFormat():
            fragments.append((fragment.position(), fragment.length(), fragment.charFormat()))
        iterator += 1
    for position, length, char_format in fragments:
        current = _inline_values(char_format)
        updates = {key: target[key] for key, value in current.items()
                   if force or value is None or value == baseline[key]}
        if updates:
            edit = QTextCursor(block.document())
            edit.setPosition(position)
            edit.setPosition(position + length, QTextCursor.MoveMode.KeepAnchor)
            edit.mergeCharFormat(_inline_format(updates))
    current = _inline_values(cursor.blockCharFormat())
    updates = {key: target[key] for key, value in current.items()
               if force or value is None or value == baseline[key]}
    cursor.mergeBlockCharFormat(_inline_format(updates))


def bind_style(cursor: QTextCursor, name: str, settings: DesignSettings | None = None) -> None:
    settings = settings or DesignSettings()
    target = resolved_style(name, settings)
    first = cursor.document().findBlock(cursor.selectionStart())
    end = cursor.selectionEnd() - (1 if cursor.hasSelection() else 0)
    last = cursor.document().findBlock(end)
    cursor.beginEditBlock()
    block = first
    while block.isValid():
        _update_block(block, name, target, target, True)
        if block == last:
            break
        block = block.next()
    cursor.endEditBlock()


def apply_design(document, settings: DesignSettings,
                 previous_settings: DesignSettings | None = None) -> None:
    settings.validate()
    previous_settings = previous_settings or DesignSettings()
    previous_settings.validate()
    cursor = QTextCursor(document)
    cursor.beginEditBlock()
    block = document.begin()
    while block.isValid():
        name = style_name(block)
        stored = block.blockFormat().property(BASELINE_PROPERTY)
        baseline = stored if isinstance(stored, dict) else resolved_style(name, previous_settings)
        _update_block(block, name, resolved_style(name, settings), baseline, False)
        block = block.next()
    cursor.endEditBlock()


def capture_style_bindings(document) -> list[dict]:
    result = []
    block = document.begin()
    while block.isValid():
        fmt = block.blockFormat()
        name, baseline = fmt.property(STYLE_PROPERTY), fmt.property(BASELINE_PROPERTY)
        if name in STYLE_SETS["Modern"] and isinstance(baseline, dict):
            result.append({"position": block.position(), "style": name,
                           "baseline": dict(baseline)})
        block = block.next()
    return result


def _valid_baseline(baseline) -> bool:
    if not isinstance(baseline, dict):
        return False
    return (isinstance(baseline.get("font"), str) and 0 < len(baseline["font"]) <= 200
            and _number(baseline.get("size"), 1, 400)
            and _number(baseline.get("weight"), 1, 1000)
            and isinstance(baseline.get("italic"), bool)
            and _color(baseline.get("color"))
            and _number(baseline.get("before"), 0, 1000)
            and _number(baseline.get("after"), 0, 1000)
            and isinstance(baseline.get("heading"), int)
            and 0 <= baseline["heading"] <= 3)


def restore_style_bindings(document, records: object) -> None:
    if not isinstance(records, list) or len(records) > document.blockCount():
        raise ValueError("Invalid paragraph style bindings")
    valid = []
    seen = set()
    for item in records:
        if not isinstance(item, dict):
            raise ValueError("Invalid paragraph style binding")
        position, name, baseline = item.get("position"), item.get("style"), item.get("baseline")
        if (not isinstance(position, int) or isinstance(position, bool)
                or position < 0 or position >= document.characterCount()
                or position in seen or not isinstance(name, str)
                or name not in STYLE_SETS["Modern"] or not _valid_baseline(baseline)):
            raise ValueError("Invalid paragraph style binding")
        block = document.findBlock(position)
        if block.position() != position:
            raise ValueError("Paragraph style binding does not identify a block")
        valid.append((block, name, baseline))
        seen.add(position)
    for block, name, baseline in valid:
        cursor = QTextCursor(block)
        fmt = block.blockFormat()
        fmt.setProperty(STYLE_PROPERTY, name)
        fmt.setProperty(BASELINE_PROPERTY, dict(baseline))
        fmt.setHeadingLevel(STYLE_SETS["Modern"][name].heading_level)
        cursor.setBlockFormat(fmt)


def paint_page_decoration(painter: QPainter, rect: QRectF, settings: DesignSettings,
                          layer: str = "behind") -> None:
    settings.validate()
    if layer not in ("behind", "front"):
        raise ValueError("Unknown page decoration layer")
    painter.save()
    painter.setClipRect(rect, Qt.ClipOperation.IntersectClip)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    if layer == "behind":
        painter.fillRect(rect, QColor(settings.page_color))
    if settings.watermark_kind != "none" and settings.watermark_layer == layer:
        painter.save()
        painter.setOpacity(settings.watermark_opacity)
        painter.translate(rect.center())
        painter.rotate(settings.watermark_angle)
        if settings.watermark_kind == "text":
            font = QFont(theme_tokens(settings)["heading_font"])
            font.setPixelSize(max(16, round(min(rect.width(), rect.height()) / 10)))
            font.setBold(True)
            painter.setFont(font)
            painter.setPen(QColor(settings.watermark_color))
            text_rect = QRectF(-rect.width() * 0.42, -rect.height() * 0.15,
                               rect.width() * 0.84, rect.height() * 0.3)
            metrics = painter.fontMetrics()
            width = metrics.horizontalAdvance(settings.watermark_text)
            if width > text_rect.width():
                font.setPixelSize(max(8, round(font.pixelSize() * text_rect.width() / width)))
                painter.setFont(font)
            painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, settings.watermark_text)
        else:
            image = watermark_image(settings.watermark_image)
            size = image.size().scaled(round(rect.width() * 0.6), round(rect.height() * 0.5),
                                       Qt.AspectRatioMode.KeepAspectRatio)
            painter.drawImage(QRectF(-size.width() / 2, -size.height() / 2,
                                     size.width(), size.height()), image)
        painter.restore()
    if layer == "front" and settings.page_border != "none":
        inset = settings.page_border_inset_mm * 96 / 25.4
        border = rect.adjusted(inset, inset, -inset, -inset)
        pen = QPen(QColor(settings.page_border_color), settings.page_border_width_pt * PT_TO_PX)
        if settings.page_border == "dashed":
            pen.setStyle(Qt.PenStyle.DashLine)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        painter.drawRect(border)
        if settings.page_border == "double":
            gap = max(3.0, pen.widthF() * 2)
            painter.drawRect(border.adjusted(gap, gap, -gap, -gap))
    painter.restore()
