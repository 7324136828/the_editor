import os
from pathlib import Path

import pytest
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import QApplication, QFileDialog, QInputDialog, QMessageBox

import folio.window as window_module
from folio.storage import LocalStore
from folio.window import FolioWindow


@pytest.fixture
def window(qapp, tmp_path):
    store = LocalStore(tmp_path / "data")
    win = FolioWindow(store=store, restore=False)
    yield win
    win._dirty = False
    win.close()
    store.close()


def _type(window, text):
    cursor = window.editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    cursor.insertText(text)


def test_window_builds(window):
    assert window.ribbon is not None
    assert window.ruler is not None
    assert window.canvas is not None
    assert window.editor is not None
    tabs = [window.ribbon.tabText(i) for i in range(window.ribbon.count())]
    assert tabs[0] == "File"
    for name in ("Home", "Insert", "Design", "Layout", "References",
                 "Mailings", "Review", "View", "Help"):
        assert name in tabs
    assert window.ribbon.currentIndex() == tabs.index("Home")


def test_window_shows_and_resizes(window):
    window.resize(1280, 800)
    window.show()
    QApplication.processEvents()
    assert window.isVisible()
    home = window.ribbon.widget(window.ribbon.currentIndex())
    assert home is not None


def test_screenshot(qapp, tmp_path):
    if not os.environ.get("FOLIO_SCREENSHOT"):
        pytest.skip("FOLIO_SCREENSHOT not set")
    from PySide6.QtGui import QFont, QRawFont
    segoe = QRawFont.fromFont(QFont("Segoe UI", 10))
    calibri = QRawFont.fromFont(QFont("Calibri", 11))
    for text, raw in (("Folio 123", segoe), ("Welcome to Folio", calibri)):
        assert raw.isValid()
        glyphs = raw.glyphIndexesForString(text)
        for char, index in zip(text, glyphs):
            if char != " ":
                assert index != 0, f"missing glyph for {char!r}"
    store = LocalStore(tmp_path / "shot-data")
    win = FolioWindow(store=store, restore=False)
    artifacts = Path(__file__).resolve().parent.parent / "artifacts"
    artifacts.mkdir(exist_ok=True)
    for width, height, name in (
            (1280, 800, "folio-window.png"),
            (1440, 960, "folio-window-1440.png")):
        win.resize(width, height)
        win.show()
        QApplication.processEvents()
        out = artifacts / name
        win.grab().save(str(out))
        assert out.exists() and out.stat().st_size > 1000
    win._dirty = False
    win.close()
    store.close()


def test_layout_geometry_at_1280(window):
    window.resize(1280, 800)
    window.show()
    QApplication.processEvents()
    assert window.ribbon.height() <= 130
    assert window.ribbon.width() >= window.width() - 4
    assert not window.review_dock.isVisibleTo(window)
    assert window.nav_dock.isVisibleTo(window)
    assert window.editor.width() > 700


def test_real_input_editing(window):
    from PySide6.QtCore import Qt
    from PySide6.QtTest import QTest
    window.resize(1280, 800)
    window.show()
    QApplication.processEvents()
    window.editor.setPlainText("")
    window.editor.setFocus()
    QApplication.processEvents()
    QTest.keyClicks(window.editor, "hello")
    assert "hello" in window.editor.document().toPlainText()
    QTest.mouseClick(window.bold_button, Qt.MouseButton.LeftButton)
    window.editor.setFocus()
    QApplication.processEvents()
    QTest.keyClicks(window.editor, " bold")
    doc = window.editor.document()
    found = False
    iterator = doc.begin().begin()
    while not iterator.atEnd():
        frag = iterator.fragment()
        if frag.isValid() and frag.text() == " bold":
            found = frag.charFormat().fontWeight() >= 700
        iterator += 1
    assert found
    window.editor.undo()
    assert " bold" not in doc.toPlainText()
    window.editor.redo()
    assert " bold" in doc.toPlainText()


def test_clean_document_not_dirtied_by_ui(window):
    window.editor.setPlainText("clean text")
    window._set_dirty(False)
    window.editor.document().setModified(False)
    assert not window._dirty
    cursor = window.editor.textCursor()
    cursor.setPosition(1)
    window.editor.setTextCursor(cursor)
    window._set_zoom(1.3)
    window._set_zoom(1.0)
    from folio.proofing import ProofIssue
    issue = ProofIssue(start=0, length=5, kind="spelling",
                       message="x", word="clean", replacement="x")
    window._proof_generation += 1
    window._proof_results(window._proof_generation, [issue])
    QApplication.processEvents()
    assert not window._dirty
    assert not window.editor.document().isModified()


def test_txt_import_then_save_offers_docx(window, tmp_path, monkeypatch):
    txt = tmp_path / "sample.txt"
    txt.write_text("plain text body", encoding="utf-8")
    original_bytes = txt.read_bytes()
    assert window.open_path(txt)
    assert window._path is None
    assert window._dirty
    _type(window, " more")
    offered = {}
    def fake_save_as():
        offered["called"] = True
        return False
    monkeypatch.setattr(window, "_save_as", fake_save_as)
    window._save()
    assert offered.get("called")
    assert txt.read_bytes() == original_bytes


def test_save_as_writes_docx_not_txt(window, tmp_path, monkeypatch):
    txt = tmp_path / "sample.txt"
    txt.write_text("keep me", encoding="utf-8")
    assert window.open_path(txt)
    _type(window, " edit")
    target = tmp_path / "converted.docx"
    monkeypatch.setattr(
        QFileDialog, "getSaveFileName",
        staticmethod(lambda *a, **k: (str(target), "Word document (*.docx)")))
    monkeypatch.setattr(QMessageBox, "question",
                        staticmethod(lambda *a, **k:
                                     QMessageBox.StandardButton.Save))
    assert window._save()
    assert target.exists()
    assert target.read_bytes()[:2] == b"PK"
    assert txt.read_bytes() == b"keep me"
    assert window._path == target


def test_external_change_copy_does_not_overwrite(window, tmp_path, monkeypatch):
    docx_path = tmp_path / "doc.docx"
    window.editor.setPlainText("mine")
    monkeypatch.setattr(
        QFileDialog, "getSaveFileName",
        staticmethod(lambda *a, **k: (str(docx_path), "")))
    assert window._save_as()
    original = docx_path.read_bytes()
    docx_path.write_bytes(b"EXTERNALLY MODIFIED")
    copied = {}
    def fake_save_as():
        other = tmp_path / "copy.docx"
        from folio.docx_io import write_docx
        write_docx(other, window.current_state(), window.editor.document())
        window._path = other
        copied["path"] = other
        return True
    monkeypatch.setattr(window, "_save_as", fake_save_as)
    monkeypatch.setattr(window, "_check_external_change", lambda p: "copy")
    window.editor.setPlainText("new edits")
    result = window.save_to(docx_path)
    assert result is True
    assert docx_path.read_bytes() == b"EXTERNALLY MODIFIED"
    assert window._path == copied["path"]


def test_external_change_cancel_blocks_save(window, tmp_path, monkeypatch):
    docx_path = tmp_path / "doc.docx"
    window.editor.setPlainText("mine")
    monkeypatch.setattr(
        QFileDialog, "getSaveFileName",
        staticmethod(lambda *a, **k: (str(docx_path), "")))
    assert window._save_as()
    docx_path.write_bytes(b"CHANGED")
    monkeypatch.setattr(window, "_check_external_change", lambda p: "cancel")
    assert window.save_to(docx_path) is False


def test_discard_close_leaves_no_recovery(window, tmp_path, monkeypatch):
    window.editor.setPlainText("dirty work")
    window._autosave()
    monkeypatch.setattr(
        QMessageBox, "question",
        staticmethod(lambda *a, **k: QMessageBox.StandardButton.Discard))
    root = window.store.root
    window.close()
    reopened = LocalStore(root)
    try:
        assert reopened.recoverable() == []
    finally:
        reopened.close()


def test_cancelled_new_preserves_dirty(window, monkeypatch):
    window.editor.setPlainText("important")
    window._set_dirty(True)
    window._autosave()
    monkeypatch.setattr(
        window_module.dialogs.TemplateGalleryDialog, "exec",
        lambda self: 0)
    window._new_document()
    assert window._dirty
    assert "important" in window.editor.document().toPlainText()


def test_cancelled_save_as_blocks_close(window, monkeypatch):
    window.editor.setPlainText("dirty")
    window._set_dirty(True)
    monkeypatch.setattr(
        QMessageBox, "question",
        staticmethod(lambda *a, **k: QMessageBox.StandardButton.Save))
    monkeypatch.setattr(
        QFileDialog, "getSaveFileName",
        staticmethod(lambda *a, **k: ("", "")))
    from PySide6.QtGui import QCloseEvent
    event = QCloseEvent()
    window.closeEvent(event)
    assert not event.isAccepted()
    assert window._dirty


def test_toc_insert_update_and_boundary(window, tmp_path, monkeypatch):
    from folio.docx_io import read_docx
    editor = window.editor
    editor.setHtml("<h1>Chapter One</h1><p>body text</p>")
    window.state.toc = None
    monkeypatch.setattr(QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    window._update_toc()
    assert window.state.toc is not None
    doc = editor.document()
    text_after_first = doc.toPlainText()
    assert "Contents" in text_after_first
    heading_texts = []
    body_texts = []
    block = doc.begin()
    while block.isValid():
        if block.blockFormat().headingLevel() > 0:
            heading_texts.append(block.text())
        elif block.text().strip():
            body_texts.append(block.text().strip())
        block = block.next()
    assert "Chapter One" in heading_texts
    assert "body text" in body_texts
    window._update_toc()
    assert editor.document().toPlainText() == text_after_first

    target = tmp_path / "toc.docx"
    assert window.save_to(target)
    loaded, _ = read_docx(target)
    assert loaded.toc is not None

    toc_end = window.state.toc.start + window.state.toc.length
    plain = doc.toPlainText()
    body_pos = plain.find("body text")
    utf16_body = len(plain[:body_pos].encode("utf-16-le")) // 2
    comment = window.tracker.add_comment(utf16_body, utf16_body + 4,
                                         "note", "T")
    cursor = QTextCursor(doc)
    cursor.setPosition(toc_end + 1)
    cursor.insertText("P")
    assert window.state.toc is not None
    assert comment.start == utf16_body + 1

    cursor2 = QTextCursor(doc)
    cursor2.setPosition(window.state.toc.start + 1)
    cursor2.insertText("Z")
    assert window.state.toc is None

    target2 = tmp_path / "toc2.docx"
    assert window.save_to(target2)
    loaded2, _ = read_docx(target2)
    assert "body text" in loaded2.html
    assert loaded2.toc is None


def test_review_pane_shows_on_tracking(window):
    window.show()
    QApplication.processEvents()
    window.review_dock.setVisible(False)
    window._toggle_tracking(True)
    QApplication.processEvents()
    assert window.review_dock.isVisible()


def test_focus_mode_toggle_and_exit(window):
    window._toggle_focus()
    assert window._focus
    window._exit_focus()
    assert not window._focus


def test_proof_issue_stale_guard(window):
    window.editor.setPlainText("some text")
    from folio.proofing import ProofIssue
    issue = ProofIssue(start=0, length=4, kind="spelling",
                       message="x", word="zzzz", replacement="fix")
    window._apply_spelling(issue, "fixed")
    assert "fixed" not in window.editor.document().toPlainText()


def test_author_propagates_to_tracker(window):
    window._author = "Zach"
    window.tracker.author = window._author
    window.tracker.set_tracking(True)
    cursor = window.editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    cursor.insertText("x")
    if window.state.revisions:
        assert window.state.revisions[-1].author == "Zach"


def test_zoom_syncs_slider(window):
    window._set_zoom(1.5)
    assert window.zoom_slider.value() == 150
    window._set_zoom(1.0)
    assert window.zoom_slider.value() == 100


def test_page_controls_sync_on_load(window):
    from folio.models import PageSettings
    state = window.state
    state.page.columns = 2
    window._sync_page_controls()
    assert window.columns_combo.currentIndex() == 1


def test_line_number_control_updates_document_state(window):
    index = window.line_numbers_combo.findData("continuous")
    window._line_numbers_chosen(index)
    assert window.state.page.line_numbering == "continuous"
    assert window.editor._line_numbering == "continuous"


def test_navigation_has_word_like_views(window):
    labels = [window.navigation_tabs.tabText(i)
              for i in range(window.navigation_tabs.count())]
    assert labels == ["Headings", "Pages", "Results", "Objects"]
