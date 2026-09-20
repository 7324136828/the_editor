from types import SimpleNamespace

import pytest
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QKeyEvent, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import QScrollArea

from folio.models import DocumentState
from folio.review import ReviewTracker
from folio.view_tools import DocumentViewport, LiveTextView
from folio.workspace_tools import (HELP_TOPICS, HelpDialog, apply_macro,
                                   feedback_record, plan_macro, search_help)


@pytest.fixture
def owner(editor):
    editor.setPlainText("Alpha 😀 Beta")
    editor.document().clearUndoRedoStacks()
    state = DocumentState()
    tracker = ReviewTracker(editor.document(), state)
    canvas = QScrollArea()
    result = SimpleNamespace(editor=editor, state=state, tracker=tracker,
        _title="View test", canvas=canvas, _can_edit=lambda: True, raise_=lambda: None)
    yield result
    canvas.deleteLater()


def select(document, start, end):
    cursor = QTextCursor(document)
    cursor.setPosition(start)
    cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
    return cursor


@pytest.mark.parametrize("draft", [False, True])
def test_live_view_insert_delete_utf16_and_reverse_sync(owner, draft):
    view = LiveTextView(owner, draft=draft)
    try:
        select(view.document(), 8, 8).insertText(" New")
        assert owner.editor.toPlainText() == "Alpha 😀 New Beta"
        select(view.document(), 6, 8).removeSelectedText()
        assert owner.editor.toPlainText() == "Alpha  New Beta"
        select(owner.editor.document(), 0, 5).insertText("Start")
        assert view.toPlainText() == "Start  New Beta"
    finally:
        view.deleteLater()


def test_web_view_format_changes_reach_source(owner):
    view = LiveTextView(owner)
    try:
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Weight.Bold)
        select(view.document(), 0, 5).mergeCharFormat(fmt)
        assert select(owner.editor.document(), 0, 5).charFormat().fontWeight() == QFont.Weight.Bold
        fmt.setFontItalic(True)
        select(owner.editor.document(), 9, 13).mergeCharFormat(fmt)
        assert select(view.document(), 9, 13).charFormat().fontItalic()
    finally:
        view.deleteLater()


def test_web_view_typing_beside_caption_survives_refresh(owner):
    from folio.document_features import insert_node
    from folio.references import refresh_references

    owner.editor.clear()
    owner.editor.document().folio_state = owner.state
    insert_node(owner.editor, owner.state, {"kind": "caption", "label": "Figure", "text": "A"}, "Figure 1: A")
    view = LiveTextView(owner)
    try:
        view.moveCursor(QTextCursor.MoveOperation.End)
        view.insertPlainText(" Additional prose")
        refresh_references(owner.editor.document(), owner.state, lambda _position: 1)
        assert owner.editor.toPlainText().endswith(" Additional prose")
        view.moveCursor(QTextCursor.MoveOperation.Start)
        event = QKeyEvent(QKeyEvent.Type.KeyPress, Qt.Key.Key_X, Qt.KeyboardModifier.NoModifier, "x")
        view.keyPressEvent(event)
        refresh_references(owner.editor.document(), owner.state, lambda _position: 1)
        assert owner.editor.toPlainText().startswith("xFigure")
    finally:
        view.deleteLater()


def test_web_view_clipboard_preserves_distinct_note_nodes(owner):
    from folio.document_features import node_ranges
    from folio.references import insert_note

    owner.editor.clear()
    owner.editor.document().folio_state = owner.state
    insert_note(owner.editor, owner.state, "Retained note")
    view = LiveTextView(owner)
    try:
        view.setTextCursor(select(view.document(), 0, 1))
        mime = view.createMimeDataFromSelection()
        view.moveCursor(QTextCursor.MoveOperation.End)
        view.insertFromMimeData(mime)
        assert len(node_ranges(owner.editor.document())) == 2
        assert [node["text"] for node in owner.state.features["nodes"].values()] == ["Retained note"] * 2
    finally:
        view.deleteLater()


def test_draft_edit_preserves_existing_rich_text(owner):
    fmt = QTextCharFormat()
    fmt.setFontWeight(QFont.Weight.Bold)
    select(owner.editor.document(), 0, 5).mergeCharFormat(fmt)
    view = LiveTextView(owner, draft=True)
    try:
        select(view.document(), 13, 13).insertText("!")
        assert owner.editor.toPlainText().endswith("Beta!")
        assert select(owner.editor.document(), 0, 5).charFormat().fontWeight() == QFont.Weight.Bold
    finally:
        view.deleteLater()


def test_live_view_keyboard_undo_uses_shared_history(owner):
    view = LiveTextView(owner)
    try:
        select(view.document(), 0, 5).insertText("Changed")
        event = QKeyEvent(QKeyEvent.Type.KeyPress, Qt.Key.Key_Z, Qt.KeyboardModifier.ControlModifier)
        view.keyPressEvent(event)
        assert owner.editor.toPlainText() == "Alpha 😀 Beta"
        assert view.toPlainText() == owner.editor.toPlainText()
        owner.editor.redo()
        assert view.toPlainText() == "Changed 😀 Beta"
    finally:
        view.deleteLater()


def test_live_view_rejects_readonly_and_protected_edits(owner):
    view = LiveTextView(owner)
    try:
        owner.editor.setReadOnly(True)
        select(view.document(), 0, 5).insertText("Changed")
        assert owner.editor.toPlainText() == "Alpha 😀 Beta"
        assert view.toPlainText() == owner.editor.toPlainText()
        owner.editor.setReadOnly(False)
        owner.state.review["locks"] = [{"start": 0, "end": 5}]
        owner.tracker.refresh_protection()
        select(view.document(), 0, 5).insertText("Changed")
        assert owner.editor.toPlainText() == "Alpha 😀 Beta"
        assert view.toPlainText() == owner.editor.toPlainText()
    finally:
        view.deleteLater()


def test_view_modes_preserve_source_and_multifit_places_two_pages(owner, qapp):
    owner.editor.setPlainText("First")
    owner.editor.moveCursor(QTextCursor.MoveOperation.End)
    owner.editor.insert_page_break()
    owner.editor.insertPlainText("Second")
    owner.editor.document().setModified(False)
    before = owner.editor.document().toHtml()
    viewport = DocumentViewport(owner)
    try:
        viewport.resize(1100, 800)
        viewport.show()
        qapp.processEvents()
        for mode in ("Web Layout", "Draft Mode", "Read Mode", "Outline Mode", "Print Layout"):
            viewport.mode.setCurrentText(mode)
            assert owner.editor.document().toHtml() == before
            assert not owner.editor.document().isModified()
        viewport.zoom.setCurrentText("Multiple Pages")
        assert viewport.pages.horizontal
        pages = sorted(viewport.pages.scene().items(), key=lambda item: item.index)
        assert len(pages) == 2
        assert pages[1].x() > pages[0].x()
        assert pages[1].y() == pages[0].y()
        total = pages[1].sceneBoundingRect().right() - pages[0].sceneBoundingRect().left()
        assert total * viewport.pages.transform().m11() <= viewport.pages.viewport().width()
        viewport.zoom.setCurrentText("Percent")
        viewport.percent.setValue(125)
        assert viewport.pages.transform().m11() == 1.25
    finally:
        viewport.close()
        viewport.deleteLater()


def test_macro_plan_and_atomic_undo(owner):
    operations = plan_macro('folio.replaceAll("Alpha", "Start"); folio.appendText("!");',
                            owner.editor.toPlainText())
    assert operations == [["replace", "Alpha", "Start"], ["append", "!"]]
    apply_macro(owner, operations)
    assert owner.editor.toPlainText() == "Start 😀 Beta!"
    owner.editor.undo()
    assert owner.editor.toPlainText() == "Alpha 😀 Beta"


def test_macro_errors_and_timeout_do_not_mutate(owner):
    before = owner.editor.toPlainText()
    with pytest.raises(ValueError, match="Error"):
        plan_macro('folio.appendText("uncommitted"); throw new Error("failure");', before)
    with pytest.raises(ValueError, match="time limit"):
        plan_macro("while (true) {}", before, timeout=0.05)
    with pytest.raises(ValueError, match="nonempty"):
        plan_macro('folio.replaceAll("", "bad");', before)
    assert owner.editor.toPlainText() == before


def test_macro_restrictions_prevent_any_operation(owner):
    owner.state.review["locks"] = [{"start": 0, "end": 5}]
    owner.tracker.refresh_protection()
    with pytest.raises(ValueError, match="restrictions"):
        apply_macro(owner, [["append", "forbidden"]])
    assert owner.editor.toPlainText() == "Alpha 😀 Beta"


def test_macro_has_no_host_file_or_network_api(qapp):
    operations = plan_macro("folio.appendText([typeof require, typeof process, typeof fetch, "
                            "typeof XMLHttpRequest].join(','));", "")
    assert operations == [["append", "undefined,undefined,undefined,undefined"]]


def test_macro_unicode_script_survives_worker_transport(qapp):
    assert plan_macro('folio.appendText("こんにちは 🧬");', "") == [["append", "こんにちは 🧬"]]


def test_macro_replace_all_is_literal_case_sensitive(owner):
    owner.editor.setPlainText("alpha Alpha alpha")
    owner.tracker.reset(owner.state)
    apply_macro(owner, plan_macro('folio.replaceAll("alpha", "X");', owner.editor.toPlainText()))
    assert owner.editor.toPlainText() == "X Alpha X"


def test_help_search_and_feedback_validation(qapp):
    assert search_help("") == HELP_TOPICS
    results = search_help("MaIl merge")
    assert any(title == "Mail merge" for title, body in results)
    assert search_help("no-such-feature-word") == []
    dialog = HelpDialog()
    dialog.search.setText("no-such-feature-word")
    assert dialog.topics.count() == 0
    assert "No topics match" in dialog.body.toPlainText()
    dialog.close()
    record = feedback_record("Bug report", " Preview broke ", " Open then preview ")
    assert record["title"] == "Preview broke"
    assert record["delivery"] == "local export"
    assert "document" not in record
    with pytest.raises(ValueError, match="category"):
        feedback_record("Unknown", "Title", "Detail")
    with pytest.raises(ValueError, match="description"):
        feedback_record("Bug report", "", "Detail")
    with pytest.raises(ValueError, match="length"):
        feedback_record("Bug report", "Title", "x" * 100001)
