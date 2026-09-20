import json

import pytest
from PySide6.QtGui import QColor, QFont, QImage, QTextCharFormat, QTextCursor

from folio.models import DocumentState, ValidationError
from folio.review import ReviewError, ReviewTracker
from folio.review_services import (
    SpeechReader, accessibility_issues, checked_ink, compare_three_way, contrast_ratio,
    load_glossary, review_document, translate_glossary,
)


def select(document, start, end):
    cursor = QTextCursor(document)
    cursor.setPosition(start)
    cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
    return cursor


def test_thread_replies_roundtrip_and_author_filter(editor):
    editor.setPlainText("😀 first and second")
    state = DocumentState()
    tracker = ReviewTracker(editor.document(), state)
    first = tracker.add_comment(3, 8, "Check first", "Alex")
    second = tracker.add_comment(13, 19, "Check second", "Sam")
    first.created, second.created = "2026-01-01", "2026-01-02"
    tracker.reply_to_comment(first.id, "Confirmed", "Sam")
    assert first.quote == "first"
    assert tracker.chronological_comments("Sam") == [first, second]
    tracker.resolve_comment(first.id)
    assert tracker.chronological_comments(include_resolved=False) == [second]
    restored = DocumentState.from_dict(state.to_dict())
    assert restored.comments[0].replies[0].text == "Confirmed"
    assert restored.comments[0].replies[0].author == "Sam"
    tracker.delete_comment(first.id)
    assert state.comments == [second]


def test_format_delta_reverts_without_moving_comments(editor):
    editor.setPlainText("😀 alpha beta")
    state = DocumentState(track_changes=True, review={"track_formatting": True})
    tracker = ReviewTracker(editor.document(), state)
    comment = tracker.add_comment(3, 8, "alpha")
    cursor = select(editor.document(), 3, 8)
    fmt = QTextCharFormat()
    fmt.setFontWeight(QFont.Weight.Bold)
    cursor.mergeCharFormat(fmt)
    pending = tracker.pending_revisions()
    assert len(pending) == 1 and pending[0].kind == "format"
    assert pending[0].before_text == pending[0].after_text == "alpha"
    tracker.reject_revision(pending[0].id)
    assert editor.toPlainText() == "😀 alpha beta"
    assert select(editor.document(), 3, 8).charFormat().fontWeight() < QFont.Weight.Bold
    assert (comment.start, comment.end, comment.orphaned) == (3, 8, False)


def test_later_format_edit_refuses_stale_rejection(editor):
    editor.setPlainText("alpha")
    state = DocumentState(track_changes=True, review={"track_formatting": True})
    tracker = ReviewTracker(editor.document(), state)
    cursor = select(editor.document(), 0, 5)
    fmt = QTextCharFormat()
    fmt.setFontItalic(True)
    cursor.mergeCharFormat(fmt)
    revision = tracker.pending_revisions()[0]
    tracker.set_tracking(False)
    fmt.setFontWeight(QFont.Weight.Bold)
    cursor.mergeCharFormat(fmt)
    with pytest.raises(ReviewError):
        tracker.reject_revision(revision.id)


def test_bulk_reject_format_on_inserted_text(editor):
    editor.setPlainText("alpha")
    state = DocumentState(track_changes=True, review={"track_formatting": True})
    tracker = ReviewTracker(editor.document(), state)
    select(editor.document(), 0, 0).insertText("X")
    fmt = QTextCharFormat()
    fmt.setFontItalic(True)
    select(editor.document(), 0, 1).mergeCharFormat(fmt)
    assert tracker.reject_all() == 2
    assert editor.toPlainText() == "alpha"


def test_paragraph_border_and_list_format_rejection(editor):
    from folio.formatting import paragraph_border

    editor.setPlainText("alpha")
    state = DocumentState(track_changes=True, review={"track_formatting": True})
    tracker = ReviewTracker(editor.document(), state)
    editor.set_paragraph_border("box", "#123456", 2)
    assert paragraph_border(editor.document().begin()) is not None
    assert tracker.reject_all() == 1
    assert paragraph_border(editor.document().begin()) is None
    editor.set_list(True)
    assert editor.document().begin().textList() is not None
    assert tracker.reject_all() >= 1
    assert editor.document().begin().textList() is None


def test_table_cell_format_rejection_preserves_structure(editor):
    table = editor.insert_table(2, 2)
    editor.setTextCursor(table.cellAt(0, 0).firstCursorPosition())
    editor.textCursor().insertText("alpha")
    state = DocumentState(track_changes=True, review={"track_formatting": True})
    tracker = ReviewTracker(editor.document(), state)
    start = table.cellAt(0, 0).firstCursorPosition().position()
    editor.setTextCursor(select(editor.document(), start, start + 5))
    editor.toggle_bold()
    assert tracker.reject_all() == 1
    assert not select(editor.document(), start, start + 5).charFormat().fontWeight() >= QFont.Weight.Bold
    assert table.rows() == 2 and table.columns() == 2
    assert table.cellAt(0, 0).firstCursorPosition().block().text() == "alpha"


def test_markup_snapshots_leave_source_untouched(editor):
    editor.setPlainText("alpha beta")
    state = DocumentState(track_changes=True)
    tracker = ReviewTracker(editor.document(), state)
    select(editor.document(), 6, 10).insertText("gamma")
    source = editor.document().toHtml()
    assert review_document(editor.document(), state, "Original").toPlainText() == "alpha beta"
    assert review_document(editor.document(), state, "No Markup").toPlainText() == "alpha gamma"
    assert "beta" in review_document(editor.document(), state, "All Markup").toPlainText()
    assert review_document(editor.document(), state, "Simple Markup").toPlainText() == "alpha gamma"
    assert editor.document().toHtml() == source
    assert len(tracker.pending_revisions()) == 1


def test_range_locks_reject_cursor_mutations_and_rebase(editor):
    editor.setPlainText("prefix locked suffix")
    state = DocumentState(review={"locks": [{"id": "lock", "name": "Section", "start": 7, "end": 13}]})
    tracker = ReviewTracker(editor.document(), state)
    blocked = []
    tracker.edit_blocked.connect(blocked.append)
    select(editor.document(), 9, 10).insertText("X")
    assert editor.toPlainText() == "prefix locked suffix"
    assert blocked
    select(editor.document(), 0, 0).insertText("😀 ")
    assert state.review["locks"][0]["start"] == 10
    assert state.review["locks"][0]["end"] == 16
    select(editor.document(), 12, 12).insertText("NO")
    assert editor.toPlainText() == "😀 prefix locked suffix"
    assert not editor.document().isRedoAvailable()


def test_lock_does_not_undo_previous_allowed_typing(editor):
    editor.setPlainText("open locked")
    state = DocumentState(review={"locks": [{"id": "lock", "name": "Section", "start": 5, "end": 11}]})
    tracker = ReviewTracker(editor.document(), state)
    cursor = select(editor.document(), 4, 4)
    cursor.insertText("A")
    cursor.insertText("B")
    before = editor.toPlainText()
    select(editor.document(), 8, 8).insertText("X")
    assert editor.toPlainText() == before


def test_formatting_lock_keeps_text_editable(editor):
    editor.setPlainText("alpha")
    state = DocumentState(review={"formatting_locked": True})
    tracker = ReviewTracker(editor.document(), state)
    cursor = select(editor.document(), 0, 5)
    fmt = QTextCharFormat()
    fmt.setFontWeight(QFont.Weight.Bold)
    cursor.mergeCharFormat(fmt)
    assert select(editor.document(), 0, 5).charFormat().fontWeight() < QFont.Weight.Bold
    select(editor.document(), 5, 5).insertText(" beta")
    assert editor.toPlainText() == "alpha beta"


def test_invalid_locks_rejected():
    payload = DocumentState().to_dict()
    payload["review"] = {"locks": [{"start": True, "end": 3}]}
    with pytest.raises(ValidationError):
        DocumentState.from_dict(payload)


def test_review_metadata_docx_roundtrip_and_rejection(editor, tmp_path):
    from folio.docx_io import read_docx, write_docx

    editor.setPlainText("alpha beta")
    state = DocumentState(track_changes=True, review={"track_formatting": True})
    tracker = ReviewTracker(editor.document(), state)
    comment = tracker.add_comment(0, 5, "Check", "Alex")
    tracker.reply_to_comment(comment.id, "Yes", "Sam")
    fmt = QTextCharFormat()
    fmt.setFontItalic(True)
    select(editor.document(), 0, 5).mergeCharFormat(fmt)
    state.review["locks"] = [{"id": "lock", "name": "Beta", "start": 6, "end": 10}]
    path = tmp_path / "review.docx"
    write_docx(path, state, editor.document())
    restored, warnings = read_docx(path)
    assert not warnings
    assert restored.comments[0].replies[0].text == "Yes"
    assert restored.review["locks"][0]["name"] == "Beta"
    assert restored.revisions[0].kind == "format"
    with tracker.loading():
        editor.setHtml(restored.html)
    tracker.reset(restored)
    tracker.reject_revision(restored.revisions[0].id)
    assert not select(editor.document(), 0, 5).charFormat().fontItalic()


def test_review_window_integration(qapp, store):
    from folio.window import FolioWindow

    window = FolioWindow(store=store, restore=False)
    try:
        state = DocumentState(html="<p>alpha beta</p>", track_changes=True,
                              review={"track_formatting": True})
        window.load_state(state, clean=True)
        assert window.format_tracking.isChecked()
        assert window.speech_reader.engine is None
        window.editor.setTextCursor(select(window.editor.document(), 0, 5))
        window.editor.toggle_bold()
        assert any(revision.kind == "format" for revision in state.revisions)
        state.review["locks"] = [{"id": "lock", "name": "Alpha", "start": 0, "end": 5}]
        window.tracker.refresh_protection()
        select(window.editor.document(), 1, 2).insertText("X")
        assert window.editor.toPlainText() == "alpha beta"
        assert "protected" in window.statusBar().currentMessage()
    finally:
        window._dirty = False
        window.close()


def test_accessibility_checks_and_alt_remediation(editor):
    editor.setHtml('<h3>Skipped levels</h3><p><a href="https://example.com">click here</a></p>')
    cursor = editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    editor.setTextCursor(cursor)
    image = QImage(10, 10, QImage.Format.Format_RGB32)
    image.fill(QColor("red"))
    editor.insert_image(image)
    issues = accessibility_issues(editor.document())
    assert {item.rule for item in issues} >= {"headings", "link_text", "image_alt"}
    issue = next(item for item in issues if item.rule == "image_alt")
    cursor = select(editor.document(), issue.start, issue.start + issue.length)
    fmt = QTextCharFormat()
    fmt.setProperty(fmt.Property.ImageAltText, "A red square")
    cursor.mergeCharFormat(fmt)
    assert not any(item.rule == "image_alt" for item in accessibility_issues(editor.document()))
    assert contrast_ratio(QColor("black"), QColor("white")) == pytest.approx(21)


def test_plain_black_text_has_no_contrast_warning(editor):
    editor.setHtml("<p>Readable black text</p>")
    assert not any(item.rule == "contrast" for item in accessibility_issues(editor.document()))


def test_three_way_merge_independent_changes_and_conflicts():
    result = compare_three_way("one\ntwo\nthree\n", "ONE\ntwo\nthree\n", "one\ntwo\nTHREE\n")
    assert result.text == "ONE\ntwo\nTHREE\n"
    assert not result.conflicts
    result = compare_three_way("one\ntwo\n", "one\nTWO\n", "one\nSECOND\n")
    assert len(result.conflicts) == 1
    assert result.conflicts[0].base == "two\n"
    assert "<<<<<<< CURRENT" in result.text
    assert compare_three_way("a", "same", "same").text == "same"
    assert compare_three_way("", "new\n", "").text == "new\n"


def test_glossary_import_translation_and_validation(tmp_path):
    path = tmp_path / "glossary.json"
    path.write_text(json.dumps({"source_language": "English", "target_language": "Spanish",
                                "entries": {"good morning": "buenos días", "hello": "hola"}}), encoding="utf-8")
    source, target, entries = load_glossary(path)
    assert (source, target) == ("English", "Spanish")
    result, matches = translate_glossary("Hello! Good morning, unknown.", entries)
    assert result == "hola! buenos días, unknown." and matches == 2
    assert translate_glossary("shelloworld", entries) == ("shelloworld", 0)
    path.write_text('{"entries": []}', encoding="utf-8")
    with pytest.raises(ValueError):
        load_glossary(path)


def test_speech_word_offsets_and_unavailable_engine(qapp, monkeypatch):
    from PySide6.QtTextToSpeech import QTextToSpeech
    monkeypatch.setattr(QTextToSpeech, "availableEngines", lambda: [])
    reader = SpeechReader()
    words, messages = [], []
    reader.word_range.connect(lambda start, length: words.append((start, length)))
    reader.message.connect(messages.append)
    assert not reader.start("Hello")
    assert "No supported local speech engine" in messages[-1]
    reader._offset = 3
    reader._word("word", 0, 5, 4)
    assert words == [(8, 4)]


def test_editable_ink_docx_roundtrip(editor, tmp_path):
    from folio.docx_io import read_docx, write_docx
    from folio.document_features import insert_node, node_ranges
    from folio.review_dialogs import InkCanvas

    data = {"width": 720, "height": 360, "strokes": [
        {"color": "#17365d", "width": 3, "points": [[20, 20], [100, 80]]}]}
    canvas = InkCanvas(ink=data)
    state = DocumentState()
    key = insert_node(editor, state, {"kind": "ink", "ink": canvas.ink()}, image=canvas.image())
    path = tmp_path / "ink.docx"
    write_docx(path, state, editor.document())
    restored, warnings = read_docx(path)
    assert not warnings
    assert restored.features["nodes"][key]["ink"] == checked_ink(data)
    editor.document().setHtml(restored.html)
    assert key in node_ranges(editor.document())
    reopened = InkCanvas(ink=restored.features["nodes"][key]["ink"])
    assert reopened.ink() == canvas.ink()
    assert reopened.image().pixelColor(20, 20).alpha() > 0
    reopened.undo()
    assert reopened.ink()["strokes"] == []
    canvas.deleteLater()
    reopened.deleteLater()


def test_draw_ink_reopens_selected_drawing(qapp, store, monkeypatch):
    from PySide6.QtCore import QPointF
    from PySide6.QtWidgets import QDialog
    from folio.document_features import node_ranges
    from folio.review_dialogs import InkDialog
    from folio.window import FolioWindow

    opened_strokes = []
    def edit(dialog):
        opened_strokes.append(len(dialog.canvas.strokes))
        dialog.canvas.strokes.append(("#17365d", 3, [QPointF(20, 20), QPointF(100, 80)]))
        return QDialog.DialogCode.Accepted
    monkeypatch.setattr(InkDialog, "exec", edit)
    window = FolioWindow(store=store, restore=False)
    try:
        window.load_state(DocumentState(), clean=True)
        window._draw_ink()
        window._draw_ink()
        assert opened_strokes == [0, 1]
        ranges = node_ranges(window.editor.document())
        assert len(ranges) == 1
        drawing = window.state.features["nodes"][next(iter(ranges))]["ink"]
        assert len(drawing["strokes"]) == 2
    finally:
        window._dirty = False
        window.close()


@pytest.mark.parametrize("change", [
    {"width": 5000}, {"height": True},
    {"strokes": [{"color": "#17365d", "width": 3, "points": [[float("nan"), 0]]}]},
    {"strokes": [{"color": "#17365d", "width": 3, "points": [[999, 0]]}]},
])
def test_invalid_ink_metadata_rejected(change):
    data = {"width": 720, "height": 360, "strokes": []}
    data.update(change)
    with pytest.raises(ValueError):
        checked_ink(data)
