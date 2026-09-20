from __future__ import annotations

import copy
import re

from PySide6.QtGui import QTextCharFormat, QTextCursor

from .document_features import ANCHOR, MARKER, insert_node, mark_cursor, node_ranges, uid
from .models import ValidationError

STYLES = ("APA", "MLA", "Chicago", "IEEE")
LABELS = ("Figure", "Table", "Equation")
GENERATED = ("bibliography", "figures", "index", "authorities", "endnotes")
REFERENCE_KINDS = ("note", "citation", "caption", "index_entry", "authority_entry", "generated")


def _text(value, limit=2000):
    return isinstance(value, str) and len(value) <= limit and "\x00" not in value


def checked_source(value):
    if not isinstance(value, dict):
        raise ValidationError("A citation source must be an object")
    result = copy.deepcopy(value)
    for name in ("author", "title", "year", "publisher", "journal", "volume", "issue", "pages", "url", "doi"):
        if not _text(result.setdefault(name, "")):
            raise ValidationError(f"Invalid source {name}")
        result[name] = result[name].strip()
    if not result["title"]:
        raise ValidationError("A citation source needs a title")
    if result.setdefault("type", "book") not in ("book", "article", "website"):
        raise ValidationError("Unsupported citation source type")
    return result


def validate_reference_features(features):
    references = features.get("references", {})
    if not isinstance(references, dict) or references.get("style", "APA") not in STYLES:
        raise ValidationError("Invalid citation settings")
    sources = references.get("sources", {})
    if not isinstance(sources, dict) or len(sources) > 1000:
        raise ValidationError("Invalid source registry")
    for key, value in sources.items():
        if not isinstance(key, str) or not re.fullmatch(r"[A-Za-z0-9_-]{1,128}", key):
            raise ValidationError("Invalid source identifier")
        sources[key] = checked_source(value)
    for node in features["nodes"].values():
        kind = node["kind"]
        if kind == "note":
            if node.get("type") not in ("footnote", "endnote") or not _text(node.get("text"), 10000):
                raise ValidationError("Invalid note")
        elif kind == "citation":
            if not isinstance(node.get("source"), str) or not re.fullmatch(r"[A-Za-z0-9_-]{1,128}", node["source"]):
                raise ValidationError("Invalid citation source")
            if not _text(node.get("locator", ""), 200):
                raise ValidationError("Invalid citation locator")
            if "source_data" in node:
                node["source_data"] = checked_source(node["source_data"])
        elif kind == "caption":
            if node.get("label") not in LABELS or not _text(node.get("text")):
                raise ValidationError("Invalid caption")
        elif kind in ("index_entry", "authority_entry"):
            if not _text(node.get("term"), 500) or not node["term"].strip():
                raise ValidationError("Enter an index term or legal authority")
            if kind == "authority_entry" and (not _text(node.get("category"), 200) or not node["category"].strip()):
                raise ValidationError("Enter an authority category")
        elif kind == "generated":
            if node.get("collection") not in GENERATED or node.get("label", "All") not in ("All", *LABELS):
                raise ValidationError("Invalid generated reference list")


def registry(state):
    return state.features.setdefault("references", {"style": "APA", "sources": {}})


def add_source(state, source, key=None):
    source = checked_source(source)
    key = key or uid()
    registry(state).setdefault("sources", {})[key] = source
    return key


def source_authors(source):
    authors = [value.strip() for value in source.get("author", "").split(";") if value.strip()]
    return authors or [source["title"]]


def _surname(author):
    return author.split(",", 1)[0].strip() if "," in author else author.split()[-1]


def citation_text(source, style="APA", number=1, locator=""):
    source = checked_source(source)
    authors = source_authors(source)
    name = _surname(authors[0]) if source["author"] else source["title"]
    if len(authors) > 2:
        name += " et al."
    elif len(authors) == 2:
        name += (" & " if style == "APA" else " and ") + _surname(authors[1])
    year = source["year"] or "n.d."
    if style == "IEEE":
        return f"[{number}" + (f", p. {locator}" if locator else "") + "]"
    if style == "MLA":
        return f"({name}" + (f" {locator}" if locator else "") + ")"
    if style == "Chicago":
        return f"({name} {year}" + (f", {locator}" if locator else "") + ")"
    return f"({name}, {year}" + (f", p. {locator}" if locator else "") + ")"


def bibliography_text(source, style="APA", number=1):
    source = checked_source(source)
    authors = "; ".join(source_authors(source)) if source["author"] else ""
    title, year = source["title"].rstrip("."), source["year"] or "n.d."
    venue = source["journal"] or source["publisher"]
    if source["type"] == "article":
        if source["volume"]:
            venue += f", {source['volume']}"
        if source["issue"]:
            venue += f"({source['issue']})"
        if source["pages"]:
            venue += f", {source['pages']}"
    suffix = (" https://doi.org/" + source["doi"].removeprefix("https://doi.org/")
              if source["doi"] else " " + source["url"] if source["url"] else "")
    if style == "IEEE":
        body = f"[{number}] " + (authors + ", " if authors else "") + f'"{title},"'
        return body + (f" {venue}," if venue else "") + f" {year}." + suffix
    if style == "MLA":
        return (authors + ". " if authors else "") + title + ". " + (venue + ", " if venue else "") + year + "." + suffix
    if style == "Chicago":
        return (authors + ". " if authors else "") + year + ". " + title + "." + (" " + venue + "." if venue else "") + suffix
    return (authors + " " if authors else "") + f"({year}). {title}." + (" " + venue + "." if venue else "") + suffix


def ordered_nodes(document, state, kind=None):
    ranges = node_ranges(document)
    return sorted(((ranges[key][0], key, node) for key, node in state.features["nodes"].items()
                   if key in ranges and (kind is None or node["kind"] == kind)), key=lambda item: item[0])


def note_entries(document, state, note_type=None):
    counts = {"footnote": 0, "endnote": 0}
    result = []
    for position, key, node in ordered_nodes(document, state, "note"):
        counts[node["type"]] += 1
        if note_type is None or node["type"] == note_type:
            result.append((position, key, counts[node["type"]], node))
    return result


def insert_note(editor, state, text, note_type="footnote"):
    key = insert_node(editor, state, {"kind": "note", "type": note_type, "text": text}, "1")
    start, end = node_ranges(editor.document())[key]
    cursor = QTextCursor(editor.document())
    cursor.setPosition(start)
    cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
    fmt = QTextCharFormat()
    fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSuperScript)
    fmt.setAnchorHref("folio-note:" + key)
    cursor.mergeCharFormat(fmt)
    if note_type == "endnote":
        ensure_collection(editor, state, "endnotes", terminal=True)
    return key


def insert_citation(editor, state, source, locator=""):
    data = registry(state)["sources"].get(source)
    if data is None:
        raise ValueError("Choose a source from this document")
    return insert_node(editor, state, {"kind": "citation", "source": source,
                                      "source_data": copy.deepcopy(data), "locator": locator},
                       citation_text(data, registry(state).get("style", "APA"), locator=locator))


def tag_selection(editor, state, term, category=None):
    cursor = editor.textCursor()
    if not cursor.hasSelection():
        raise ValueError("Select the text to mark first")
    node = {"kind": "authority_entry" if category is not None else "index_entry", "term": term.strip()}
    if category is not None:
        node["category"] = category.strip()
    validate_reference_features({"nodes": {"tag": node}})
    key = uid()
    mark_cursor(cursor, key)
    state.features["nodes"][key] = node
    return key


def ensure_collection(editor, state, collection, label="All", terminal=False):
    for _, key, node in ordered_nodes(editor.document(), state, "generated"):
        if node["collection"] == collection and node.get("label", "All") == label:
            return key
    saved = editor.textCursor()
    saved.setKeepPositionOnInsert(True)
    cursor = editor.textCursor()
    if terminal:
        cursor.movePosition(QTextCursor.MoveOperation.End)
    if cursor.block().text():
        cursor.insertBlock()
    cursor.setCharFormat(QTextCharFormat())
    editor.setTextCursor(cursor)
    key = insert_node(editor, state, {"kind": "generated", "collection": collection, "label": label},
                      {"bibliography": "Bibliography", "figures": "Table of Figures", "index": "Index",
                       "authorities": "Table of Authorities", "endnotes": "Endnotes"}[collection])
    cursor = editor.textCursor()
    cursor.insertBlock()
    saved.setKeepPositionOnInsert(False)
    editor.setTextCursor(saved if terminal else cursor)
    return key


def reference_values(document, state, page_for_position):
    nodes = ordered_nodes(document, state)
    config = registry(state)
    style = config.get("style", "APA")
    sources = config.setdefault("sources", {})
    values, counters, cited, captions = {}, {}, [], []
    indexes, authorities = {}, {}
    notes = note_entries(document, state)
    for _, key, number, node in notes:
        values[key] = str(number)
    for position, key, node in nodes:
        kind = node["kind"]
        if kind == "citation":
            source_key = node["source"]
            if source_key not in cited:
                cited.append(source_key)
            source = sources.get(source_key) or node.get("source_data")
            if source:
                node["source_data"] = copy.deepcopy(source)
                values[key] = citation_text(source, style, cited.index(source_key) + 1, node.get("locator", ""))
            else:
                values[key] = "[Source missing]"
        elif kind == "caption":
            label = node["label"]
            counters[label] = counters.get(label, 0) + 1
            values[key] = f"{label} {counters[label]}: {node['text']}"
            captions.append((position, key, node))
        elif kind == "index_entry":
            indexes.setdefault(node["term"], set()).add(page_for_position(position))
        elif kind == "authority_entry":
            authorities.setdefault(node["category"], {}).setdefault(node["term"], set()).add(page_for_position(position))
    all_sources = {key: sources.get(key) or next((node.get("source_data") for _, _, node in nodes
                                               if node["kind"] == "citation" and node["source"] == key), None)
                   for key in cited}
    for _, key, node in nodes:
        if node["kind"] != "generated":
            continue
        collection = node["collection"]
        if collection == "bibliography":
            order = cited if style == "IEEE" else sorted(cited, key=lambda item: (
                str((all_sources[item] or {}).get("author", "")).casefold(),
                str((all_sources[item] or {}).get("year", "")),
                str((all_sources[item] or {}).get("title", "")).casefold()))
            lines = ["Bibliography"] + [bibliography_text(all_sources[item], style, cited.index(item) + 1)
                                         if all_sources[item] else "[Source missing]" for item in order]
        elif collection == "figures":
            label = node.get("label", "All")
            lines = ["Table of Figures" if label == "All" else "List of " + label + "s"]
            lines += [f"{values[target]}  ·  p. {page_for_position(position)}" for position, target, caption in captions
                      if label == "All" or caption["label"] == label]
        elif collection == "index":
            lines = ["Index"] + [f"{term}  {', '.join(map(str, sorted(indexes[term])))}"
                                   for term in sorted(indexes, key=str.casefold)]
        elif collection == "authorities":
            lines = ["Table of Authorities"]
            for category in sorted(authorities, key=str.casefold):
                lines.append(category)
                lines += [f"{term}  {', '.join(map(str, sorted(authorities[category][term])))}"
                          for term in sorted(authorities[category], key=str.casefold)]
        else:
            lines = ["Endnotes"] + [f"{number}. {note['text']}" for _, _, number, note in notes if note["type"] == "endnote"]
        values[key] = "\n".join(lines)
    return values


def refresh_references(document, state, page_for_position=lambda position: 1, can_edit=None):
    values = reference_values(document, state, page_for_position)
    ranges = node_ranges(document)
    changes = []
    for key, value in values.items():
        start, end = ranges[key]
        if can_edit is not None and not can_edit(start, end):
            continue
        cursor = QTextCursor(document)
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        if cursor.selectedText().replace("\u2029", "\n").replace("\u2028", "\n") != value:
            changes.append((start, end, value, cursor.charFormat()))
    edit = QTextCursor(document)
    if changes:
        if can_edit is None:
            edit.beginEditBlock()
        for start, end, value, fmt in sorted(changes, reverse=True, key=lambda item: item[0]):
            if can_edit is not None:
                edit.beginEditBlock()
            edit.setPosition(start)
            edit.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            edit.insertText(value, fmt)
            if can_edit is not None:
                edit.endEditBlock()
        if can_edit is None:
            edit.endEditBlock()
    return len(changes)
