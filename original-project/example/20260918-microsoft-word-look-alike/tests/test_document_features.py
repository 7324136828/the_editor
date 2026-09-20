import base64
import copy
import hashlib
from datetime import datetime

import pytest
from PySide6.QtGui import QColor, QImage, QTextCursor

from folio.document_features import (
    attachment_node, bookmark, checked_features, group_objects, insert_node,
    insert_section, mark_cursor, move_group_children, node_ranges, refresh_fields,
    section_ranges, sign_node, visible_objects,
)
from folio.editor import SafeDocument
from folio.models import DocumentState, PageSettings, ValidationError
from folio.review import ReviewTracker


def select(editor, start, end):
    cursor = editor.textCursor()
    cursor.setPosition(start)
    cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
    editor.setTextCursor(cursor)
    return cursor


def end(editor):
    cursor = editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    cursor.clearSelection()
    editor.setTextCursor(cursor)
    return cursor


def test_bookmark_full_unicode_range_survives_html_and_interior_edits(editor):
    editor.setHtml("<p>\U0001f600 <b>alpha</b> beta</p><p>omega</p>")
    state = DocumentState()
    select(editor, 3, 19)
    key = bookmark(editor, state, "Span")
    assert node_ranges(editor.document())[key] == (3, 19)
    editor.document().setHtml(editor.document().toHtml())
    assert node_ranges(editor.document())[key] == (3, 19)
    select(editor, 5, 5).insertText("XX")
    assert node_ranges(editor.document())[key] == (3, 21)
    select(editor, 0, 0).insertText("\U0001f600")
    assert node_ranges(editor.document())[key] == (5, 23)


def test_nested_bookmarks_preserve_both_identities(editor):
    editor.setPlainText("outer inner outer")
    state = DocumentState()
    select(editor, 6, 11)
    inner = bookmark(editor, state, "Inner")
    editor.selectAll()
    outer = bookmark(editor, state, "Outer")
    restored = SafeDocument()
    restored.setHtml(editor.document().toHtml())
    assert node_ranges(restored) == {inner: (6, 11), outer: (0, 17)}


def test_image_node_and_following_typing_keep_separate_ranges(editor):
    state = DocumentState()
    image = QImage(30, 20, QImage.Format.Format_RGB32)
    image.fill(QColor("red"))
    key = insert_node(editor, state, {"kind": "diagram", "steps": ["A", "B"]}, image=image)
    editor.textCursor().insertText("following")
    assert node_ranges(editor.document())[key] == (0, 1)
    restored = SafeDocument()
    restored.setHtml(editor.document().toHtml())
    assert node_ranges(restored)[key] == (0, 1)
    assert restored.toPlainText() == "\ufffcfollowing"


def test_datetime_reload_refresh_does_not_duplicate_suffix_or_consume_following_text(editor):
    state = DocumentState()
    key = insert_node(editor, state, {"kind": "datetime", "format": "date"}, "Date")
    editor.textCursor().insertText(" | footer")
    refresh_fields(editor.document(), state.features, datetime(2026, 9, 19))
    assert editor.toPlainText() == "September 19, 2026 | footer"
    editor.document().setHtml(editor.document().toHtml())
    refresh_fields(editor.document(), state.features, datetime(2027, 1, 1))
    assert editor.toPlainText() == "January 01, 2027 | footer"
    assert node_ranges(editor.document())[key] == (0, 16)
    assert refresh_fields(editor.document(), state.features, datetime(2027, 1, 1)) == 0


def test_reference_tracks_renamed_bookmark_and_reports_deleted_target(editor):
    editor.setPlainText("Alpha")
    state = DocumentState()
    editor.selectAll()
    target = bookmark(editor, state, "Target")
    cursor = end(editor)
    cursor.insertBlock()
    editor.setTextCursor(cursor)
    reference = insert_node(editor, state, {"kind": "reference", "target": target}, "Reference")
    refresh_fields(editor.document(), state.features)
    assert editor.toPlainText() == "Alpha\nAlpha"
    select(editor, 1, 4).insertText("str")
    refresh_fields(editor.document(), state.features)
    assert editor.toPlainText() == "Astra\nAstra"
    editor.document().setHtml(editor.document().toHtml())
    select(editor, 0, 5).removeSelectedText()
    refresh_fields(editor.document(), state.features)
    assert editor.toPlainText() == "\n[Reference missing]"
    assert reference in node_ranges(editor.document())


def test_refresh_rebases_comments_between_multiple_changed_fields(editor):
    state = DocumentState()
    insert_node(editor, state, {"kind": "datetime", "format": "time"}, "x")
    editor.textCursor().insertText(" unchanged ")
    insert_node(editor, state, {"kind": "datetime", "format": "time"}, "y")
    tracker = ReviewTracker(editor.document(), state)
    comment = tracker.add_comment(2, 11, "note")
    refresh_fields(editor.document(), state.features, datetime(2026, 9, 19, 13, 15))
    assert editor.toPlainText() == "13:15 unchanged 13:15"
    assert (comment.start, comment.end, comment.orphaned) == (6, 15, False)


def test_references_in_cycles_have_stable_error_text(editor):
    state = DocumentState()
    first = insert_node(editor, state, {"kind": "reference", "target": "missing"}, "one")
    editor.textCursor().insertText(" | ")
    second = insert_node(editor, state, {"kind": "reference", "target": first}, "two")
    state.features["nodes"][first]["target"] = second
    refresh_fields(editor.document(), state.features)
    assert editor.toPlainText() == "[Circular reference] | [Circular reference]"
    assert refresh_fields(editor.document(), state.features) == 0


def test_copy_paste_clones_fields_and_preserves_intervening_text(editor):
    from PySide6.QtTest import QTest

    state = DocumentState()
    editor.document().folio_state = state
    first = insert_node(editor, state, {"kind": "datetime", "format": "time"}, "13:15")
    select(editor, 0, 5)
    mime = editor.createMimeDataFromSelection()
    assert "folio-block-formats" in mime.html(), mime.html()
    copied = SafeDocument()
    copied.setHtml(mime.html())
    assert node_ranges(copied)[first] == (0, 5)
    end(editor)
    QTest.keyClicks(editor, " safe text ")
    editor.insert_mime_data(mime)
    assert len(node_ranges(editor.document())) == 2
    assert len(state.features["nodes"]) == 2
    assert sorted(limit - start for start, limit in node_ranges(editor.document()).values()) == [5, 5]
    refresh_fields(editor.document(), state.features, datetime(2026, 9, 19, 9, 5))
    assert editor.toPlainText() == "09:05 safe text 09:05"
    assert node_ranges(editor.document())[first] == (0, 5)
    editor.document().setHtml(editor.document().toHtml())
    refresh_fields(editor.document(), state.features, datetime(2026, 9, 19, 10, 30))
    assert editor.toPlainText() == "10:30 safe text 10:30"


def test_disjoint_legacy_node_ids_never_consume_gap(editor):
    editor.setPlainText("one PROTECTED two")
    first = select(editor, 0, 3)
    mark_cursor(first, "same")
    second = select(editor, 14, 17)
    mark_cursor(second, "same")
    state = DocumentState(features={"nodes": {"same": {"kind": "datetime", "format": "time"}}})
    assert node_ranges(editor.document())["same"] == (0, 3)
    refresh_fields(editor.document(), state.features, datetime(2026, 9, 19, 9, 5))
    assert editor.toPlainText() == "09:05 PROTECTED two"


def test_section_marker_survives_roundtrip_and_prefix_edits(editor):
    editor.setPlainText("first")
    state = DocumentState()
    end(editor)
    key = insert_section(editor, state, "odd", PageSettings(landscape=True))
    editor.textCursor().insertText("second")
    assert node_ranges(editor.document())[key] == (6, 7)
    editor.document().setHtml(editor.document().toHtml())
    select(editor, 0, 0).insertText("\U0001f600")
    ranges = section_ranges(editor.document(), state)
    assert ranges[0][:2] == (0, 8)
    assert ranges[1][0] == 8 and ranges[1][2].landscape and ranges[1][3] == "odd"


def test_signature_requires_explicit_consent_and_records_fingerprint(editor):
    node = {"kind": "signature", "signer": "Ada"}
    editor.setPlainText("Contract")
    for name, consent in (("", True), ("Ada", False), ("Ada", "yes")):
        with pytest.raises(ValueError):
            sign_node(node, name, consent, editor.document())
    sign_node(node, " Ada ", True, editor.document())
    assert node["signed_by"] == "Ada" and node["method"] == "typed"
    assert node["body_sha256"] == hashlib.sha256(b"Contract").hexdigest()
    assert checked_features({"nodes": {"sig": node}})["nodes"]["sig"] == node


@pytest.mark.parametrize("features", [
    {"nodes": {"x": {"kind": "datetime", "format": "unknown"}}},
    {"nodes": {"x": {"kind": "reference", "target": {}}}},
    {"nodes": {"x": {"kind": "attachment", "name": "../secret", "data": ""}}},
    {"nodes": {"x": {"kind": "attachment", "name": "file", "data": "????"}}},
    {"nodes": {"x": {"kind": "attachment", "name": "file", "data": "eA==", "sha256": "0" * 64}}},
    {"objects": [{"id": "x", "kind": "text", "font_size": 100000}]},
    {"objects": [{"id": "x", "kind": "text", "fill": "invalid"}]},
    {"objects": [{"id": "x", "kind": "image", "image": "data:image/png;base64,eA=="}]},
    {"objects": [{"id": "x", "kind": "group", "group": "x"}]},
    {"sections": [{"id": "x", "start": "next"}, {"id": "x", "start": "odd"}]},
])
def test_invalid_feature_payloads_rejected(features):
    with pytest.raises(ValidationError):
        checked_features(features)


def test_attachments_validate_digest_and_roundtrip(tmp_path):
    path = tmp_path / "sample.bin"
    path.write_bytes(b"\x00payload\xff")
    node = attachment_node(path)
    assert base64.b64decode(node["data"]) == path.read_bytes()
    assert checked_features({"nodes": {"attachment": node}})["nodes"]["attachment"] == node


def test_group_movement_visibility_and_cycle_detection_are_bounded():
    objects = checked_features({"objects": [
        {"id": "a", "kind": "text", "x": 10, "y": 20},
        {"id": "b", "kind": "text", "x": 60, "y": 50},
    ]})["objects"]
    key = group_objects(objects, {"a", "b"})
    move_group_children(objects, key, 5, -10)
    assert objects[0]["x"] == 15 and objects[1]["y"] == 40
    objects[-1]["visible"] = False
    assert list(visible_objects(objects)) == []
    objects[-1]["group"] = key
    before = copy.deepcopy(objects)
    with pytest.raises(ValueError):
        move_group_children(objects, key, 1, 1)
    assert objects == before
    with pytest.raises(ValueError):
        list(visible_objects(objects))
