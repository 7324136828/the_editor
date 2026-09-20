from __future__ import annotations

import base64
import copy
import hashlib
import json
import math
import re
import uuid
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QByteArray, QBuffer, QIODevice, QObject, QRectF, Qt, Signal
from PySide6.QtGui import QColor, QFont, QImage, QImageReader, QPainter, QPen, QTextCharFormat, QTextCursor, QTextFormat, QTextImageFormat

from .models import PageSettings, ValidationError

ANCHOR = "folio-node:"
MARKER = "\u2063"
MAX_OBJECTS = 1000
MAX_ATTACHMENT_BYTES = 8 * 1024 * 1024
NODE_IDS_PROPERTY = int(QTextFormat.Property.UserProperty) + 1802
MAX_FEATURE_BYTES = 16 * 1024 * 1024
NODE_MIME_TYPE = "application/x-folio-nodes+json"


def _identifier(value):
    return isinstance(value, str) and re.fullmatch(r"[A-Za-z0-9_-]{1,128}", value)


def _text(value, maximum=10000):
    return isinstance(value, str) and len(value) <= maximum and "\x00" not in value


def _decode_payload(value, limit=MAX_ATTACHMENT_BYTES):
    if not isinstance(value, str) or len(value) > (limit + 2) // 3 * 4:
        raise ValidationError("Embedded payload exceeds the size limit")
    try:
        data = base64.b64decode(value, validate=True)
    except ValueError as exc:
        raise ValidationError("Invalid embedded payload encoding") from exc
    if len(data) > limit:
        raise ValidationError("Embedded payload exceeds the size limit")
    return data


def _png_payload(value):
    prefix = "data:image/png;base64,"
    if not isinstance(value, str) or not value.startswith(prefix):
        raise ValidationError("Floating images require embedded PNG data")
    data = _decode_payload(value[len(prefix):])
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        raise ValidationError("Invalid PNG image")
    buffer = QBuffer()
    buffer.setData(QByteArray(data))
    buffer.open(QIODevice.OpenModeFlag.ReadOnly)
    reader = QImageReader(buffer, b"PNG")
    size = reader.size()
    if not size.isValid() or size.width() * size.height() > 40_000_000:
        raise ValidationError("Floating image dimensions exceed the supported size")
    if reader.read().isNull():
        raise ValidationError("Invalid PNG image")
    return data


def uid() -> str:
    return uuid.uuid4().hex


def checked_features(value) -> dict:
    if value is None:
        return {"objects": [], "nodes": {}, "sections": []}
    if not isinstance(value, dict):
        raise ValidationError("Document features must be an object")
    try:
        encoded = json.dumps(value, allow_nan=False).encode("utf-8")
    except (TypeError, ValueError, RecursionError) as exc:
        raise ValidationError("Document features must contain finite JSON values") from exc
    if len(encoded) > MAX_FEATURE_BYTES:
        raise ValidationError("Document features exceed the metadata size limit")
    result = copy.deepcopy(value)
    objects = result.setdefault("objects", [])
    nodes = result.setdefault("nodes", {})
    sections = result.setdefault("sections", [])
    if not isinstance(objects, list) or not isinstance(nodes, dict) or not isinstance(sections, list):
        raise ValidationError("Invalid document feature collections")
    if len(objects) + len(nodes) + len(sections) > MAX_OBJECTS:
        raise ValidationError("Too many document objects")
    ids = set()
    for obj in objects:
        if not isinstance(obj, dict) or not _identifier(obj.get("id")) or obj["id"] in ids:
            raise ValidationError("Invalid or duplicate floating object ID")
        ids.add(obj["id"])
        if obj.get("kind") not in ("text", "image", "group"):
            raise ValidationError("Unsupported floating object type")
        for name, default in (("x", 40), ("y", 40), ("width", 200), ("height", 100), ("z", 0)):
            number = obj.setdefault(name, default)
            if isinstance(number, bool) or not isinstance(number, (float, int)) or not math.isfinite(number):
                raise ValidationError(f"Invalid object {name}")
            if abs(number) > 100000 or (name in ("width", "height") and number < 1):
                raise ValidationError(f"Object {name} out of range")
        if obj.get("wrap", "front") not in ("square", "tight", "behind", "front"):
            raise ValidationError("Invalid wrapping mode")
        page = obj.setdefault("page", 0)
        if not isinstance(page, int) or isinstance(page, bool) or not 0 <= page <= 10000:
            raise ValidationError("Invalid object page")
        if not isinstance(obj.setdefault("visible", True), bool):
            raise ValidationError("Invalid object visibility")
        for name in ("text", "image", "name", "group"):
            if name in obj and not isinstance(obj[name], str):
                raise ValidationError(f"Object {name} must be text")
        if not _text(obj.get("text", ""), 100000) or not _text(obj.get("name", ""), 500):
            raise ValidationError("Floating object text is too long")
        for name in ("fill", "stroke", "color"):
            if name in obj and (not isinstance(obj[name], str) or not QColor(obj[name]).isValid()):
                raise ValidationError("Invalid floating object color")
        if "font_size" in obj and (type(obj["font_size"]) not in (int, float)
                                   or not 1 <= obj["font_size"] <= 400):
            raise ValidationError("Invalid floating object font size")
        if "theme_bound" in obj and not isinstance(obj["theme_bound"], bool):
            raise ValidationError("Invalid theme binding")
        if obj.get("image") or obj["kind"] == "image":
            _png_payload(obj.get("image", ""))
    by_id = {obj["id"]: obj for obj in objects}
    for obj in objects:
        visited = {obj["id"]}
        parent = obj.get("group")
        while parent:
            if parent in visited or parent not in by_id or by_id[parent]["kind"] != "group":
                raise ValidationError("Invalid or cyclic object group")
            if obj["page"] != by_id[parent]["page"]:
                raise ValidationError("Grouped objects must be on the same page")
            visited.add(parent)
            parent = by_id[parent].get("group")
    for key, node in nodes.items():
        if not _identifier(key) or not isinstance(node, dict) or key in ids:
            raise ValidationError("Invalid inline node")
        ids.add(key)
        if node.get("kind") not in ("bookmark", "reference", "datetime", "equation", "model", "attachment", "signature", "diagram",
                                    "note", "citation", "caption", "index_entry", "authority_entry", "generated", "ink"):
            raise ValidationError("Unsupported inline node")
        if node.get("kind") == "attachment":
            payload = _decode_payload(node.get("data"))
            name = node.get("name")
            if not _text(name, 255) or not name or re.search(r"[/\\:\x00]", name) or name in (".", ".."):
                raise ValidationError("Invalid attachment filename")
            digest = hashlib.sha256(payload).hexdigest()
            if node.get("sha256", digest) != digest:
                raise ValidationError("The attachment digest does not match its contents")
            node["sha256"] = digest
        kind = node["kind"]
        if kind == "ink":
            from .review_services import checked_ink

            try:
                node["ink"] = checked_ink(node.get("ink"))
            except ValueError as exc:
                raise ValidationError(str(exc)) from exc
        if kind == "bookmark" and (not isinstance(node.get("name"), str)
                                    or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,63}", node["name"])):
            raise ValidationError("Invalid bookmark name")
        if kind == "reference" and (not _identifier(node.get("target"))
                                     or node.get("display", "text") not in ("text", "page")):
            raise ValidationError("Invalid cross-reference")
        if kind == "datetime" and node.get("format", "date") not in ("date", "time", "datetime"):
            raise ValidationError("Invalid date field format")
        if kind == "equation":
            from .structured_media import parse_equation

            if node.get("syntax") not in ("tex", "mathml"):
                raise ValidationError("Invalid equation syntax")
            parse_equation(node.get("source"), node["syntax"])
        if kind == "model":
            from .structured_media import ModelScene

            ModelScene.from_dict(node.get("scene"))
        if kind == "diagram":
            steps = node.get("steps")
            if not isinstance(steps, list) or not 1 <= len(steps) <= 100 or any(not _text(step, 2000) for step in steps):
                raise ValidationError("Invalid diagram steps")
        if kind == "signature":
            for name in ("signer", "name", "signed_by", "signed_at"):
                if name in node and not _text(node[name], 500):
                    raise ValidationError("Invalid signature field")
            if "signed_by" in node:
                if (not node["signed_by"].strip() or node.get("consent") is not True
                        or node.get("method") != "typed"
                        or not isinstance(node.get("body_sha256"), str)
                        or not re.fullmatch(r"[0-9a-f]{64}", node["body_sha256"])):
                    raise ValidationError("Invalid typed signature record")
                try:
                    datetime.fromisoformat(node.get("signed_at", ""))
                except ValueError as exc:
                    raise ValidationError("Invalid signature timestamp") from exc
    for section in sections:
        if (not isinstance(section, dict) or not _identifier(section.get("id"))
                or section["id"] in ids
                or section.get("start") not in ("next", "continuous", "even", "odd")):
            raise ValidationError("Invalid section")
        ids.add(section["id"])
        PageSettings.from_dict(section.get("page", {}))
    from .references import validate_reference_features

    validate_reference_features(result)
    if "mailing" in result:
        from .mailmerge import MergeError, checked_mailing

        try:
            result["mailing"] = checked_mailing(result["mailing"])
        except MergeError as exc:
            raise ValidationError(str(exc)) from exc
    return result


def fragments(document):
    block = document.begin()
    while block.isValid():
        iterator = block.begin()
        while not iterator.atEnd():
            fragment = iterator.fragment()
            if fragment.isValid():
                yield fragment
            iterator += 1
        block = block.next()


def node_ranges(document) -> dict[str, tuple[int, int]]:
    ranges = {}
    closed = set()
    for fragment in fragments(document):
        for key in _node_ids(fragment.charFormat()):
            if key in closed:
                continue
            start, end = ranges.get(key, (fragment.position(), fragment.position()))
            if fragment.position() > end:
                gap = QTextCursor(document)
                gap.setPosition(end)
                gap.setPosition(fragment.position(), QTextCursor.MoveMode.KeepAnchor)
                if any(char not in "\u2029\u2028" for char in gap.selectedText()):
                    closed.add(key)
                    continue
            ranges[key] = (min(start, fragment.position()),
                           max(end, fragment.position() + fragment.length()))
    return ranges


def _node_ids(fmt):
    value = fmt.property(NODE_IDS_PROPERTY)
    keys = list(value) if isinstance(value, list) else []
    keys.extend(name[len(ANCHOR):] for name in fmt.anchorNames() if name.startswith(ANCHOR))
    return sorted({key for key in keys if _identifier(key)})


def capture_node_bindings(document):
    result = []
    for fragment in fragments(document):
        keys = _node_ids(fragment.charFormat())
        if not keys:
            continue
        start, end = fragment.position(), fragment.position() + fragment.length()
        if result and result[-1]["end"] == start and result[-1]["keys"] == keys:
            result[-1]["end"] = end
        else:
            result.append({"start": start, "end": end, "keys": keys})
    return result


def restore_node_bindings(document, records, span=None):
    if not isinstance(records, list) or len(records) > 100000:
        return
    valid = []
    for record in records:
        if not isinstance(record, dict):
            return
        start, end, keys = record.get("start"), record.get("end"), record.get("keys")
        if (type(start) is not int or type(end) is not int
                or not 0 <= start < end < document.characterCount()
                or not isinstance(keys, list) or not 1 <= len(keys) <= MAX_OBJECTS
                or any(not _identifier(key) for key in keys)):
            return
        if valid and start < valid[-1][1]:
            return
        valid.append((start, end, keys))
    cursor = QTextCursor(document)
    cursor.beginEditBlock()
    low, high = span if span is not None else (0, document.characterCount() - 1)
    parsed = [(max(low, fragment.position()), min(high, fragment.position() + fragment.length()),
               fragment.charFormat()) for fragment in fragments(document)
              if _node_ids(fragment.charFormat()) and fragment.position() < high
              and fragment.position() + fragment.length() > low]
    for position, limit, fmt in parsed:
        cursor.setPosition(position)
        cursor.setPosition(limit, QTextCursor.MoveMode.KeepAnchor)
        cursor.setCharFormat(_outside_node_format(fmt, detach_link=False))
    for start, end, keys in valid:
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        fmt = QTextCharFormat()
        fmt.setAnchor(True)
        fmt.setAnchorNames([ANCHOR + key for key in keys])
        fmt.setProperty(NODE_IDS_PROPERTY, keys)
        cursor.mergeCharFormat(fmt)
    cursor.endEditBlock()


def _outside_node_format(fmt, detach_link=True):
    result = QTextCharFormat(fmt)
    names = [name for name in result.anchorNames() if not name.startswith(ANCHOR)]
    result.setAnchorNames(names)
    result.clearProperty(NODE_IDS_PROPERTY)
    if detach_link and result.anchorHref().startswith(("folio-note:", "#" + ANCHOR)):
        if result.anchorHref().startswith("folio-note:"):
            result.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignNormal)
        result.setAnchorHref("")
    result.setAnchor(bool(names or result.anchorHref()))
    return result


def detach_completed_node(editor):
    state = getattr(editor.document(), "folio_state", None)
    cursor = editor.textCursor()
    if state is None or cursor.hasSelection():
        return
    fmt = cursor.charFormat()
    keys = _node_ids(fmt)
    if not keys:
        return
    ranges = node_ranges(editor.document())
    removed = {key for key in keys
               if cursor.position() in ranges.get(key, (-1, -1))
               and state.features["nodes"].get(key, {}).get("kind") != "bookmark"}
    if not removed:
        return
    remaining = [key for key in keys if key not in removed]
    fmt.setAnchorNames([name for name in fmt.anchorNames()
                        if name not in [ANCHOR + key for key in removed]])
    if remaining:
        fmt.setProperty(NODE_IDS_PROPERTY, remaining)
    else:
        fmt.clearProperty(NODE_IDS_PROPERTY)
    href = fmt.anchorHref()
    if href.startswith("folio-note:") and href[len("folio-note:"):] in removed:
        fmt.setAnchorHref("")
        fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignNormal)
    if href.startswith("#" + ANCHOR) and any(
            state.features["nodes"].get(key, {}).get("kind") == "reference" for key in removed):
        fmt.setAnchorHref("")
    fmt.setAnchor(bool(fmt.anchorNames() or fmt.anchorHref()))
    editor.setCurrentCharFormat(fmt)


def clipboard_feature_payload(document, source_state):
    keys = node_ranges(document)
    values = source_state.features
    return json.dumps({"source": source_state.id, "features": {
        "nodes": {key: value for key, value in values["nodes"].items() if key in keys},
        "sections": [value for value in values["sections"] if value["id"] in keys],
    }}, allow_nan=False).encode("utf-8")


def prepare_pasted_nodes(document, target_state, payload=None):
    source = {"nodes": {}, "sections": []}
    if payload and len(payload) <= MAX_FEATURE_BYTES:
        try:
            data = json.loads(payload)
            source = checked_features(data["features"])
        except (ValueError, TypeError, KeyError):
            source = {"nodes": {}, "sections": []}
    keys = node_ranges(document)
    mapped = {key: uid() for key in keys
              if target_state is not None and (key in source["nodes"] or
                 any(section["id"] == key for section in source["sections"]))}
    nodes = {}
    sections = []
    names = {node.get("name") for node in target_state.features["nodes"].values()
             if node["kind"] == "bookmark"} if target_state else set()
    for key, replacement in mapped.items():
        if key in source["nodes"]:
            node = copy.deepcopy(source["nodes"][key])
            if node["kind"] == "reference":
                node["target"] = mapped.get(node["target"], node["target"])
            if node["kind"] == "bookmark":
                original = node["name"]
                index = 1
                while node["name"] in names:
                    suffix = "_copy" + str(index)
                    node["name"] = original[:64 - len(suffix)] + suffix
                    index += 1
                names.add(node["name"])
            if node["kind"] == "signature":
                for name in ("signed_by", "signed_at", "body_sha256", "method", "consent"):
                    node.pop(name, None)
            nodes[replacement] = node
        else:
            section = next(copy.deepcopy(item) for item in source["sections"] if item["id"] == key)
            section["id"] = replacement
            sections.append(section)
    if target_state:
        merged = copy.deepcopy(target_state.features)
        merged["nodes"].update(nodes)
        merged["sections"].extend(sections)
        checked_features(merged)
    edits = [(fragment.position(), fragment.length(), fragment.charFormat())
             for fragment in fragments(document) if _node_ids(fragment.charFormat())]
    for start, length, original in edits:
        fmt = _outside_node_format(original)
        names = [mapped[key] for key in _node_ids(original) if key in mapped]
        if names:
            fmt.setProperty(NODE_IDS_PROPERTY, names)
            fmt.setAnchorNames(fmt.anchorNames() + [ANCHOR + key for key in names])
            fmt.setAnchor(True)
            link = original.anchorHref()
            for prefix in ("folio-note:", "#" + ANCHOR):
                if link.startswith(prefix):
                    target = link[len(prefix):]
                    fmt.setAnchorHref(prefix + mapped.get(target, target))
                    if prefix == "folio-note:":
                        fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSuperScript)
        cursor = QTextCursor(document)
        cursor.setPosition(start)
        cursor.setPosition(start + length, QTextCursor.MoveMode.KeepAnchor)
        cursor.setCharFormat(fmt)
    if target_state:
        target_state.features["nodes"].update(nodes)
        target_state.features["sections"].extend(sections)


def mark_cursor(cursor, key):
    if not _identifier(key):
        raise ValueError("Invalid node identifier")
    start, end = cursor.selectionStart(), cursor.selectionEnd()
    selected = [(max(start, fragment.position()),
                 min(end, fragment.position() + fragment.length()), fragment.charFormat())
                for fragment in fragments(cursor.document())
                if fragment.position() < end and fragment.position() + fragment.length() > start]
    for lo, hi, current in selected:
        keys = sorted(set(_node_ids(current) + [key]))
        fmt = QTextCharFormat()
        fmt.setAnchor(True)
        names = [name for name in current.anchorNames() if not name.startswith(ANCHOR)]
        fmt.setAnchorNames(names + [ANCHOR + node_key for node_key in keys])
        fmt.setProperty(NODE_IDS_PROPERTY, keys)
        edit = QTextCursor(cursor.document())
        edit.setPosition(lo)
        edit.setPosition(hi, QTextCursor.MoveMode.KeepAnchor)
        edit.mergeCharFormat(fmt)


def insert_node(editor, state, node, text=None, image=None) -> str:
    key = uid()
    node = checked_features({"nodes": {key: node}})["nodes"][key]
    editor.document().folio_state = state
    if image is not None and (image.isNull() or image.width() * image.height() > 40_000_000):
        raise ValueError("Insert a valid image within the supported size")
    if text is not None and not _text(text, 100000):
        raise ValueError("Invalid node text")
    cursor = editor.textCursor()
    old_format = _outside_node_format(cursor.charFormat())
    cursor.beginEditBlock()
    if image is not None:
        from .editor import _image_to_data_uri
        fmt = QTextImageFormat()
        fmt.setName(_image_to_data_uri(image))
        width = min(float(image.width()), editor.content_width_px())
        fmt.setWidth(width)
        fmt.setHeight(width * image.height() / max(1, image.width()))
        fmt.setAnchor(True)
        fmt.setAnchorNames([ANCHOR + key])
        fmt.setProperty(NODE_IDS_PROPERTY, [key])
        cursor.insertImage(fmt)
    else:
        fmt = QTextCharFormat(old_format)
        fmt.setAnchor(True)
        fmt.setAnchorNames([ANCHOR + key])
        fmt.setProperty(NODE_IDS_PROPERTY, [key])
        fmt.setAnchorHref("#" + ANCHOR + node["target"] if node["kind"] == "reference" else "")
        cursor.insertText(text or MARKER, fmt)
    end = cursor.position()
    cursor.setCharFormat(old_format)
    cursor.endEditBlock()
    editor.setTextCursor(cursor)
    editor.setCurrentCharFormat(old_format)
    state.features["nodes"][key] = node
    return key


def bookmark(editor, state, name: str) -> str:
    name = name.strip()
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,63}", name):
        raise ValueError("Bookmark names start with a letter and contain letters, digits, or underscores.")
    if any(n.get("kind") == "bookmark" and n.get("name") == name for n in state.features["nodes"].values()):
        raise ValueError("A bookmark with this name already exists.")
    cursor = editor.textCursor()
    if not cursor.hasSelection():
        cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
    key = uid()
    state.features["nodes"][key] = {"kind": "bookmark", "name": name}
    if cursor.hasSelection():
        cursor.beginEditBlock()
        mark_cursor(cursor, key)
        cursor.endEditBlock()
        if not editor.textCursor().hasSelection():
            editor.setCurrentCharFormat(_outside_node_format(editor.currentCharFormat()))
    else:
        return insert_node(editor, state, state.features["nodes"].pop(key), MARKER)
    return key


def refresh_fields(document, features, now=None, page_for_position=None, can_edit=None) -> int:
    now = now or datetime.now()
    ranges = node_ranges(document)
    changes = []

    def field_value(key, visiting):
        if key in visiting:
            return "[Circular reference]"
        node = features["nodes"].get(key)
        if node is None or key not in ranges:
            return "[Reference missing]"
        if node["kind"] == "datetime":
            style = node.get("format", "date")
            return now.strftime({"date": "%B %d, %Y", "time": "%H:%M", "datetime": "%Y-%m-%d %H:%M"}.get(style, "%Y-%m-%d"))
        if node["kind"] == "reference":
            target = ranges.get(node.get("target"))
            if target is None:
                return "[Reference missing]"
            if node.get("display") == "page":
                if page_for_position:
                    return str(page_for_position(target[0]))
                block = document.findBlock(target[0])
                height = document.pageSize().height()
                top = document.documentLayout().blockBoundingRect(block).top()
                return str(int(top / height) + 1 if height > 0 else 1)
            start, end = ranges[key]
            if target[0] < end and start < target[1]:
                return "[Circular reference]"
            target_node = features["nodes"].get(node.get("target"), {})
            if target_node.get("kind") in ("datetime", "reference"):
                return field_value(node["target"], visiting | {key})
            cursor = QTextCursor(document)
            cursor.setPosition(target[0])
            cursor.setPosition(target[1], QTextCursor.MoveMode.KeepAnchor)
            return cursor.selectedText().replace("\u2029", " ").replace(MARKER, "")
        return None

    for key, node in features["nodes"].items():
        if key not in ranges or node["kind"] not in ("datetime", "reference"):
            continue
        value = field_value(key, set())
        if value is not None:
            start, end = ranges[key]
            if can_edit is not None and not can_edit(start, end):
                continue
            cursor = QTextCursor(document)
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            if cursor.selectedText() != value:
                changes.append((start, end, value, cursor.charFormat()))
    for index, (start, end, value, fmt) in enumerate(sorted(changes, reverse=True, key=lambda entry: entry[0])):
        cursor = QTextCursor(document)
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        if index:
            cursor.joinPreviousEditBlock()
        else:
            cursor.beginEditBlock()
        cursor.insertText(value, fmt)
        cursor.endEditBlock()
    return len(changes)


def insert_section(editor, state, kind, settings):
    settings.validate()
    if kind not in ("next", "continuous", "even", "odd"):
        raise ValueError("Unknown section break")
    cursor = editor.textCursor()
    if cursor.currentTable():
        raise ValueError("Insert section breaks outside a table.")
    cursor.beginEditBlock()
    cursor.insertBlock()
    key = uid()
    fmt = QTextCharFormat()
    fmt.setAnchor(True)
    fmt.setAnchorNames([ANCHOR + key])
    fmt.setProperty(NODE_IDS_PROPERTY, [key])
    fmt.setFontPointSize(0.1)
    cursor.insertText(MARKER, fmt)
    cursor.setCharFormat(QTextCharFormat())
    cursor.endEditBlock()
    editor.setTextCursor(cursor)
    editor.setCurrentCharFormat(QTextCharFormat())
    state.features["sections"].append({"id": key, "start": kind, "page": settings.to_dict()})
    return key


def section_ranges(document, state):
    ranges = node_ranges(document)
    entries = sorted(((ranges[s["id"]][0], s) for s in state.features["sections"]
                      if s["id"] in ranges), key=lambda item: item[0])
    result = []
    start, settings, kind = 0, state.page, "next"
    for position, section in entries:
        if position > start:
            result.append((start, position, settings, kind))
        start = position
        settings = PageSettings.from_dict(section["page"])
        kind = section["start"]
    result.append((start, document.characterCount() - 1, settings, kind))
    return result


def attachment_node(path: Path):
    path = Path(path)
    if path.stat().st_size > MAX_ATTACHMENT_BYTES:
        raise ValueError("Attachments are limited to 8 MiB.")
    with path.open("rb") as stream:
        data = stream.read(MAX_ATTACHMENT_BYTES + 1)
    if len(data) > MAX_ATTACHMENT_BYTES:
        raise ValueError("Attachments are limited to 8 MiB.")
    return {"kind": "attachment", "name": path.name,
            "data": base64.b64encode(data).decode("ascii"),
            "sha256": hashlib.sha256(data).hexdigest()}


def sign_node(node: dict, name: str, consent: bool, document) -> None:
    if node.get("kind") != "signature" or not _text(name, 500) or not name.strip() or consent is not True:
        raise ValueError("Enter your name and confirm that you intend to sign.")
    node.update({"signed_by": name.strip(), "signed_at": datetime.now().astimezone().isoformat(),
                 "body_sha256": hashlib.sha256(document.toPlainText().encode("utf-8")).hexdigest(),
                 "method": "typed", "consent": True})


def visible_objects(objects):
    by_id = {obj["id"]: obj for obj in objects}
    for obj in sorted(objects, key=lambda item: item.get("z", 0)):
        current = obj
        visible = True
        visited = set()
        while current:
            if current["id"] in visited:
                raise ValueError("Cyclic object group")
            visited.add(current["id"])
            visible = visible and current.get("visible", True)
            current = by_id.get(current.get("group"))
        if visible and obj["kind"] != "group":
            yield obj


def object_rect(obj):
    return QRectF(obj["x"], obj["y"], obj["width"], obj["height"])


def object_image(obj):
    value = obj.get("image", "")
    if not value.startswith("data:image/png;base64,"):
        return QImage()
    try:
        return QImage.fromData(_png_payload(value))
    except ValidationError:
        return QImage()


def paint_objects(painter, objects, page, layer="front"):
    for obj in visible_objects(objects):
        if obj.get("page", 0) != page or (obj.get("wrap") == "behind") != (layer == "behind"):
            continue
        rect = object_rect(obj)
        if obj["kind"] == "image":
            painter.drawImage(rect, object_image(obj))
        else:
            painter.save()
            painter.setClipRect(rect)
            painter.fillRect(rect, QColor(obj.get("fill", "#e0ecff")))
            painter.setPen(QColor(obj.get("stroke", "#185abd")))
            painter.drawRect(rect)
            font = QFont("Calibri")
            font.setPointSizeF(obj.get("font_size", 14))
            painter.setFont(font)
            painter.drawText(rect.adjusted(8, 6, -8, -6),
                             Qt.TextFlag.TextWordWrap | Qt.AlignmentFlag.AlignTop,
                             obj.get("text", ""))
            painter.restore()


def align_objects(objects, selected, operation):
    if operation not in ("left", "center", "top", "distribute"):
        raise ValueError("Unknown alignment operation")
    items = [obj for obj in objects if obj["id"] in selected]
    if len(items) < 2:
        raise ValueError("Select at least two objects.")
    if len({obj.get("page", 0) for obj in items}) != 1:
        raise ValueError("Align objects on the same page.")
    left = min(obj["x"] for obj in items)
    right = max(obj["x"] + obj["width"] for obj in items)
    top = min(obj["y"] for obj in items)
    for obj in items:
        old_x, old_y = obj["x"], obj["y"]
        if operation == "left":
            obj["x"] = left
        elif operation == "center":
            obj["x"] = (left + right - obj["width"]) / 2
        elif operation == "top":
            obj["y"] = top
        if obj["kind"] == "group":
            move_group_children(objects, obj["id"], obj["x"] - old_x, obj["y"] - old_y)
    if operation == "distribute":
        items.sort(key=lambda obj: obj["x"])
        gap = (right - left - sum(obj["width"] for obj in items)) / (len(items) - 1)
        x = left
        for obj in items:
            delta = x - obj["x"]
            obj["x"] = x
            if obj["kind"] == "group":
                move_group_children(objects, obj["id"], delta, 0)
            x += obj["width"] + gap


def move_group_children(objects, group, dx, dy):
    pending = [group]
    visited = set()
    descendants = []
    while pending:
        current = pending.pop()
        if current in visited:
            raise ValueError("Cyclic object group")
        visited.add(current)
        for child in objects:
            if child.get("group") == current:
                descendants.append(child)
                if child["kind"] == "group":
                    pending.append(child["id"])
    for child in descendants:
        child["x"] += dx
        child["y"] += dy


def group_objects(objects, selected):
    items = [obj for obj in objects if obj["id"] in selected]
    if len(items) < 2 or len({obj.get("page", 0) for obj in items}) != 1:
        raise ValueError("Select at least two objects on the same page.")
    if any(obj.get("group") for obj in items):
        raise ValueError("Ungroup the selected objects before regrouping them.")
    bounds = object_rect(items[0])
    for obj in items[1:]:
        bounds = bounds.united(object_rect(obj))
    key = uid()
    group = {"id": key, "kind": "group", "name": "Group", "x": bounds.x(), "y": bounds.y(),
             "width": bounds.width(), "height": bounds.height(), "page": items[0].get("page", 0),
             "z": max(obj.get("z", 0) for obj in items), "visible": True}
    objects.append(group)
    for obj in items:
        obj["group"] = key
    return key


class CollaborationHooks(QObject):
    commentDispatched = Signal(dict)
    dictationStateChanged = Signal(str)
    transcriptReady = Signal(str)
    modeChanged = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.mode = "Editing"
        self.dictation_state = "idle"

    def set_mode(self, mode):
        if mode not in ("Editing", "Reviewing", "Viewing"):
            raise ValueError("Unknown permission view state")
        self.mode = mode
        if mode == "Viewing":
            self.set_dictation_state("idle")
        self.modeChanged.emit(mode)

    def set_dictation_state(self, state):
        if state not in ("idle", "listening", "processing", "error"):
            raise ValueError("Unknown dictation state")
        if self.mode == "Viewing" and state == "listening":
            raise ValueError("Dictation is unavailable in Viewing mode")
        self.dictation_state = state
        self.dictationStateChanged.emit(state)

    def submit_transcript(self, text):
        if self.dictation_state not in ("listening", "processing") or self.mode == "Viewing":
            raise ValueError("Start dictation before submitting a transcript")
        if not isinstance(text, str):
            raise ValueError("Transcript must be text")
        self.transcriptReady.emit(text)

    def dispatch_comment(self, comment):
        if self.mode == "Viewing":
            raise ValueError("Comments cannot be dispatched in Viewing mode")
        self.commentDispatched.emit(copy.deepcopy(comment.to_dict()))
