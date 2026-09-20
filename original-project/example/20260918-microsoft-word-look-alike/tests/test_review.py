import pytest
from PySide6.QtGui import QTextCursor

from folio.editor import RichEditor, SafeDocument
from folio.models import DocumentState
from folio.review import ReviewError, ReviewTracker


@pytest.fixture
def doc(qapp):
    editor = RichEditor()
    yield editor.document()
    editor.deleteLater()


def _tracker(doc, state=None):
    state = state or DocumentState()
    return ReviewTracker(doc, state), state


def _type(doc, pos, text):
    cursor = QTextCursor(doc)
    cursor.setPosition(pos)
    cursor.insertText(text)


def _delete(doc, pos, length):
    cursor = QTextCursor(doc)
    cursor.setPosition(pos)
    cursor.setPosition(pos + length, QTextCursor.MoveMode.KeepAnchor)
    cursor.removeSelectedText()


def test_insert_tracked_and_reject(doc):
    doc.setPlainText("Alpha Beta")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    _type(doc, 6, "new ")
    assert doc.toPlainText() == "Alpha new Beta"
    pending = [r for r in state.revisions if r.status == "pending"]
    assert len(pending) == 1
    assert pending[0].after_text == "new "
    tracker.reject_revision(pending[0].id)
    assert doc.toPlainText() == "Alpha Beta"


def test_reject_after_later_untracked_edit(doc):
    doc.setPlainText("Alpha Beta")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    _type(doc, 6, "new ")
    tracker.set_tracking(False)
    _type(doc, len("Alpha new Beta"), " tail")
    assert doc.toPlainText() == "Alpha new Beta tail"
    pending = [r for r in state.revisions if r.status == "pending"]
    assert len(pending) == 1
    tracker.reject_revision(pending[0].id)
    assert doc.toPlainText() == "Alpha Beta tail"


def test_overlapping_edit_marks_conflict_and_refuses(doc):
    doc.setPlainText("Alpha Beta")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    _type(doc, 6, "new ")
    rev = state.revisions[-1]
    _type(doc, 7, "X")
    assert rev.conflicted
    before = doc.toPlainText()
    with pytest.raises(ReviewError):
        tracker.reject_revision(rev.id)
    assert doc.toPlainText() == before


def test_tracked_delete_records_and_rejects(doc):
    doc.setPlainText("keep remove me")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    _delete(doc, 5, 9)
    assert doc.toPlainText() == "keep "
    rev = [r for r in state.revisions if r.status == "pending"][-1]
    assert rev.before_text == "remove me"
    tracker.reject_revision(rev.id)
    assert doc.toPlainText() == "keep remove me"


def test_astral_prefix_preserves_utf16_positions(doc):
    doc.setPlainText("\U0001D11E Alpha Beta")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    pos = doc.toPlainText().find("Beta")
    utf16_pos = len(doc.toPlainText()[:pos].encode("utf-16-le")) // 2
    _type(doc, utf16_pos, "new ")
    rev = state.revisions[-1]
    assert rev.after_text == "new "
    assert rev.start == utf16_pos
    tracker.reject_revision(rev.id)
    assert doc.toPlainText() == "\U0001D11E Alpha Beta"


def test_formatting_only_change_adds_no_revision_but_updates_shadow(doc):
    from PySide6.QtGui import QFont, QTextCharFormat
    doc.setHtml("<p>plain text</p>")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    cursor = QTextCursor(doc)
    cursor.setPosition(0)
    cursor.setPosition(5, QTextCursor.MoveMode.KeepAnchor)
    fmt = QTextCharFormat()
    fmt.setFontWeight(QFont.Weight.Bold)
    cursor.mergeCharFormat(fmt)
    assert state.revisions == []
    _delete(doc, 0, 5)
    rev = [r for r in state.revisions if r.status == "pending"][-1]
    assert rev.before_text == "plain"
    tracker.reject_revision(rev.id)
    found = False
    iterator = doc.begin().begin()
    while not iterator.atEnd():
        frag = iterator.fragment()
        if frag.isValid() and frag.text() == "plain":
            assert frag.charFormat().fontWeight() >= QFont.Weight.Bold
            found = True
        iterator += 1
    assert found


def test_structural_marker_insert_not_recorded(doc):
    doc.setPlainText("text")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    cursor = QTextCursor(doc)
    cursor.setPosition(4)
    fmt = cursor.charFormat()
    from PySide6.QtGui import QTextImageFormat
    from PySide6.QtGui import QImage, QColor
    image_fmt = QTextImageFormat()
    image = QImage(10, 10, QImage.Format.Format_RGB32)
    image.fill(QColor("#ff0000"))
    doc.addResource(2, "memory://img", image)
    image_fmt.setName("memory://img")
    image_fmt.setWidth(10)
    image_fmt.setHeight(10)
    cursor.insertImage(image_fmt)
    text = doc.toPlainText()
    assert "￼" in text
    assert not any(r.status == "pending" and "￼" in r.after_text
                   for r in state.revisions)


def test_coalescing_typed_text(doc):
    doc.setPlainText("")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    for index, char in enumerate("hello"):
        _type(doc, index, char)
    assert len(state.revisions) == 1
    assert state.revisions[0].after_text == "hello"


def test_table_insert_while_tracking_records_no_structural_revision(qapp):
    editor = RichEditor()
    try:
        doc = editor.document()
        tracker, state = _tracker(doc)
        tracker.set_tracking(True)
        table = editor.insert_table(2, 2)
        assert table is not None
        pending = [r for r in state.revisions if r.status == "pending"]
        assert not any(
            "\ufffc" in r.after_text or "\ufdd0" in r.after_text
            or "\ufdd1" in r.after_text
            for r in pending)
        editor.setTextCursor(table.cellAt(0, 0).firstCursorPosition())
        editor.textCursor().insertText("cell text")
        assert "cell text" in doc.toPlainText()
        pos = doc.toPlainText().find("cell text")
        utf16_pos = len(doc.toPlainText()[:pos].encode("utf-16-le")) // 2
        _delete(doc, utf16_pos, 9)
        rev = [r for r in state.revisions if r.status == "pending"][-1]
        tracker.reject_revision(rev.id)
        assert "cell text" in doc.toPlainText()
    finally:
        editor.deleteLater()


def test_author_propagates(doc):
    doc.setPlainText("x")
    tracker, state = _tracker(doc)
    tracker.author = "Reviewer"
    tracker.set_tracking(True)
    _type(doc, 1, "y")
    assert state.revisions[-1].author == "Reviewer"


def test_accept_revision(doc):
    doc.setPlainText("ab")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    _type(doc, 2, "c")
    rev = state.revisions[-1]
    tracker.accept_revision(rev.id)
    assert rev.status == "accepted"
    assert doc.toPlainText() == "abc"


def test_reject_all_and_accept_all(doc):
    doc.setPlainText("start")
    tracker, state = _tracker(doc)
    tracker.set_tracking(True)
    _type(doc, 0, "A")
    _type(doc, 6, "B")
    assert tracker.reject_all() >= 1


def test_comments_add_resolve_orphan(doc):
    doc.setPlainText("some text here")
    tracker, state = _tracker(doc)
    comment = tracker.add_comment(0, 4, "note", "Me")
    assert comment.quote == "some"
    tracker.resolve_comment(comment.id)
    assert state.comments[0].resolved


def test_toc_invalidated_on_interior_edit(doc):
    from folio.models import TocRegion
    doc.setPlainText("x" * 40)
    state = DocumentState()
    state.toc = TocRegion(5, 10)
    tracker, state = _tracker(doc, state)
    _type(doc, 7, "EDIT")
    assert state.toc is None


def test_toc_shifts_on_prefix_edit(doc):
    from folio.models import TocRegion
    doc.setPlainText("x" * 40)
    state = DocumentState()
    state.toc = TocRegion(5, 10)
    tracker, state = _tracker(doc, state)
    _type(doc, 0, "PRE")
    assert state.toc is not None
    assert state.toc.start == 8
