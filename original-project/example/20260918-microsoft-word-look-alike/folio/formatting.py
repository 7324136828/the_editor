from __future__ import annotations

import base64
import json
import math
import re

from PySide6.QtCore import QPointF
from PySide6.QtGui import QColor, QPen, QTextCursor, QTextFormat


FONT_SIZE_STEPS = (8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28,
                   36, 48, 72, 96, 120, 144, 192, 288, 384, 576, 768, 999)
PARAGRAPH_BORDER_PROPERTY = int(QTextFormat.Property.UserProperty) + 1800
MULTILEVEL_LIST_PROPERTY = int(QTextFormat.Property.UserProperty) + 1801
_ENVELOPE = re.compile(r"<!--folio-block-formats:([A-Za-z0-9+/=]+)-->")
_MAX_ENVELOPE = 4 * 1024 * 1024
_SIDES = ("top", "bottom", "left", "right")


def stepped_font_size(points: float, direction: int) -> float:
    if not math.isfinite(points) or points <= 0:
        raise ValueError("Font size must be a positive finite number")
    if direction > 0:
        return float(next((size for size in FONT_SIZE_STEPS if size > points),
                          FONT_SIZE_STEPS[-1]))
    if direction < 0:
        return float(next((size for size in reversed(FONT_SIZE_STEPS)
                           if size < points), FONT_SIZE_STEPS[0]))
    return points


def border_spec(kind: str, color: str = "#263244", width: float = 1.0) -> dict:
    if kind not in (*_SIDES, "box", "none"):
        raise ValueError("Unknown paragraph border configuration")
    parsed = QColor(color)
    if not parsed.isValid() or not math.isfinite(width) or not 0.25 <= width <= 12:
        raise ValueError("Borders need a valid color and a width from 0.25 to 12")
    return {"sides": list(_SIDES) if kind == "box" else ([] if kind == "none" else [kind]),
            "color": parsed.name(), "width": float(width)}


def paragraph_border(block) -> dict | None:
    raw = block.blockFormat().property(PARAGRAPH_BORDER_PROPERTY)
    if not isinstance(raw, str):
        return None
    try:
        value = json.loads(raw)
        sides = value["sides"]
        color = QColor(value["color"])
        width = float(value["width"])
        if (not isinstance(sides, list) or not sides
                or any(side not in _SIDES for side in sides)
                or not color.isValid() or not math.isfinite(width)
                or not 0.25 <= width <= 12):
            return None
        return {"sides": list(dict.fromkeys(sides)), "color": color.name(), "width": width}
    except (ValueError, TypeError, KeyError):
        return None


def serialize_block_metadata(document, source: str) -> str:
    records = []
    block = document.begin()
    while block.isValid():
        border = paragraph_border(block)
        multilevel = bool(block.textList() and block.textList().format().property(
            MULTILEVEL_LIST_PROPERTY))
        if border or multilevel:
            records.append({"block": block.blockNumber(), "border": border,
                            "multilevel": multilevel})
        block = block.next()
    from .design import capture_style_bindings
    from .document_features import capture_node_bindings, fragments

    styles = capture_style_bindings(document)
    nodes = capture_node_bindings(document)
    images = []
    for fragment in fragments(document):
        fmt = fragment.charFormat()
        description = fmt.property(QTextFormat.Property.ImageAltText)
        if fmt.isImageFormat() and isinstance(description, str) and description:
            images.append({"start": fragment.position(), "alt": description})
    if not records and not styles and not nodes and not images:
        return source
    payload = json.dumps({"blocks": records, "styles": styles, "nodes": nodes, "images": images},
                         separators=(",", ":")).encode("utf-8")
    if len(payload) > _MAX_ENVELOPE:
        raise ValueError("Paragraph formatting metadata exceeds the supported size")
    return source + "\n<!--folio-block-formats:" + base64.b64encode(payload).decode("ascii") + "-->"


def restore_block_metadata(document, source: str) -> None:
    match = _ENVELOPE.search(source)
    if match is None or len(match[1]) > _MAX_ENVELOPE * 4 // 3 + 4:
        return
    try:
        payload = json.loads(base64.b64decode(match[1], validate=True))
    except (ValueError, TypeError):
        return
    if not isinstance(payload, dict):
        return
    records = payload.get("blocks", [])
    if not isinstance(records, list) or len(records) > document.blockCount():
        return
    for record in records:
        if not isinstance(record, dict) or not isinstance(record.get("block"), int):
            continue
        block = document.findBlockByNumber(record["block"])
        if not block.isValid():
            continue
        border = record.get("border")
        if isinstance(border, dict):
            cursor = QTextCursor(block)
            fmt = cursor.blockFormat()
            fmt.setProperty(PARAGRAPH_BORDER_PROPERTY, json.dumps(border))
            cursor.setBlockFormat(fmt)
            if paragraph_border(block) is None:
                fmt.clearProperty(PARAGRAPH_BORDER_PROPERTY)
                cursor.setBlockFormat(fmt)
        if record.get("multilevel") is True and block.textList():
            text_list = block.textList()
            fmt = text_list.format()
            fmt.setProperty(MULTILEVEL_LIST_PROPERTY, True)
            text_list.setFormat(fmt)
    from .design import restore_style_bindings

    styles = payload.get("styles", [])
    if isinstance(styles, list):
        try:
            restore_style_bindings(document, styles)
        except ValueError:
            pass
    from .document_features import restore_node_bindings

    if "nodes" in payload:
        restore_node_bindings(document, payload["nodes"])
    images = payload.get("images", [])
    if isinstance(images, list):
        for image in images:
            if (not isinstance(image, dict) or type(image.get("start")) is not int
                    or not 0 <= image["start"] < document.characterCount() - 1
                    or not isinstance(image.get("alt"), str)):
                continue
            cursor = QTextCursor(document)
            cursor.setPosition(image["start"])
            cursor.setPosition(image["start"] + 1, QTextCursor.MoveMode.KeepAnchor)
            fmt = cursor.charFormat()
            if fmt.isImageFormat():
                fmt.setProperty(QTextFormat.Property.ImageAltText, image["alt"])
                cursor.setCharFormat(fmt)


def draw_paragraph_borders(painter, document) -> None:
    painter.save()
    layout = document.documentLayout()
    block = document.begin()
    while block.isValid():
        spec = paragraph_border(block)
        if spec:
            rect = layout.blockBoundingRect(block)
            inset = spec["width"] / 2.0
            rect.adjust(inset, inset, -inset, -inset)
            painter.setPen(QPen(QColor(spec["color"]), spec["width"]))
            endpoints = {
                "top": (rect.topLeft(), rect.topRight()),
                "bottom": (rect.bottomLeft(), rect.bottomRight()),
                "left": (rect.topLeft(), rect.bottomLeft()),
                "right": (rect.topRight(), rect.bottomRight()),
            }
            for side in spec["sides"]:
                start, end = endpoints[side]
                painter.drawLine(QPointF(start), QPointF(end))
        block = block.next()
    painter.restore()
