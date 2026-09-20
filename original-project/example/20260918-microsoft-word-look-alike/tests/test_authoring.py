import os
import zipfile
from pathlib import Path

import pytest
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import QDialog, QDialogButtonBox, QMenu

from folio import authoring
from folio.docx_io import read_docx, write_docx
from folio.models import Comment, DocumentState, PageSettings, Revision
from folio.window import FolioWindow


@pytest.fixture
def window(qapp, store):
    win = FolioWindow(store=store, restore=False)
    win.load_state(DocumentState(html="<p>Hello world.</p>"), clean=True)
    yield win
    win._dirty = False
    win.close()


@pytest.mark.parametrize("url", [
    "https://youtube.com.evil.example/watch?v=x", "http://vimeo.com/123",
    "https://user:password@ted.com/talks/test", "file:///C:/video.mp4",
])
def test_video_provider_rejects_unsupported_addresses(url):
    with pytest.raises(ValueError):
        authoring.video_provider(url)


def test_inserted_content_survives_docx(editor, tmp_path):
    url = "https://www.youtube.com/watch?v=example"
    assert authoring.video_provider(url) == "YouTube"
    editor.insert_link(url, "Example video")
    editor.textCursor().insertBlock()
    equation = authoring.TextInsertDialog("Equation")
    equation.insert(editor)
    editor.textCursor().insertBlock()
    art = authoring.TextInsertDialog("WordArt")
    art.text.setText("Headline")
    art.insert(editor)
    signature = authoring.SignatureDialog()
    signature.name.setText("A <Signer>")
    editor.textCursor().insertHtml(signature.field_html())
    editor.insert_image(authoring.stock_art("Cartoon People", 1))
    path = tmp_path / "insertions.docx"
    write_docx(path, DocumentState(), editor.document())
    loaded, warnings = read_docx(path)
    assert not warnings
    doc = authoring.SafeDocument()
    doc.setHtml(loaded.html)
    assert "A <Signer>" in doc.toPlainText()
    assert "a² + b² = c²" in doc.toPlainText()
    assert "Headline" in doc.toPlainText()
    with zipfile.ZipFile(path) as archive:
        assert any(name.startswith("word/media/") for name in archive.namelist())
        assert url in archive.read("word/_rels/document.xml.rels").decode()


def test_stock_picker_disables_insert_on_category_change(qapp):
    dialog = authoring.StockMediaDialog()
    items = dialog.tabs.widget(0).findChild(authoring.QListWidget)
    dialog._select(items.item(0))
    assert not dialog.selected_image.isNull()
    dialog.tabs.setCurrentIndex(1)
    assert dialog.selected_image is None
    assert not dialog.box.button(QDialogButtonBox.StandardButton.Ok).isEnabled()
    for category in ("Icons", "Stickers", "Illustrations", "Cartoon People"):
        for variant in range(4):
            assert not authoring.stock_art(category, variant).isNull()


def test_statistics_use_publication_and_do_not_modify(editor):
    editor.setPlainText("Hello 😀 world\nSecond paragraph")
    editor.document().setModified(False)
    before = editor.document().toHtml()
    stats = authoring.document_statistics(editor.document(), PageSettings(), "Stats")
    assert stats["Words"] == 4
    assert stats["Paragraphs"] == 2
    assert stats["Characters (with spaces)"] == 29
    assert stats["Pages"] == 1
    assert stats["Lines"] >= 2
    assert editor.document().toHtml() == before
    assert not editor.document().isModified()


def test_design_preserves_images_links_heading_structure_and_undo(editor):
    editor.setHtml('<h1>Title</h1><p>Body <a href="https://example.com">link</a></p>')
    editor.insert_image(authoring.stock_art("Icons", 0))
    before = editor.document().toHtml()
    authoring.apply_design(editor, "Forest", ("Cambria", "Calibri"), "Large Print")
    assert editor.document().begin().blockFormat().headingLevel() == 1
    assert "#166534" in editor.document().toHtml()
    assert "https://example.com" in editor.document().toHtml()
    assert "data:image" in editor.document().toHtml()
    editor.undo()
    assert editor.document().toHtml() == before


def test_reading_controls_never_mutate_document(editor, qapp):
    editor.setHtml("<h1>Reading</h1><p>Communication and understanding.</p>")
    before = editor.document().toHtml()
    reader = authoring.ReadingDialog(editor.document())
    reader.show()
    qapp.processEvents()
    reader.syllables.setChecked(True)
    assert "·" in reader.reader.toPlainText()
    reader.color.setCurrentText("Dark")
    reader.width.setCurrentIndex(0)
    reader.line_focus.setCurrentIndex(2)
    assert len(reader.reader.extraSelections()) == 1
    reader.syllables.setChecked(False)
    assert "Communication" in reader.reader.toPlainText()
    assert editor.document().toHtml() == before
    reader.close()


def test_comparison_escapes_content_and_shows_both_versions():
    result = authoring.comparison_html("keep\n<script>old</script>\n", "keep\nnew\n")
    assert "<script>" not in result
    assert "&lt;script&gt;old&lt;/script&gt;" in result
    assert "+ new" in result
    assert "No text differences" in authoring.comparison_html("same", "same")


def test_page_numbers_preserve_header_and_export_fields(window, tmp_path):
    window.state.page.header = "Report"
    window._page_numbers("top")
    window._page_numbers("top")
    assert window.state.page.header.count("{page}") == 1
    assert window.state.page.header.startswith("Report")
    path = tmp_path / "pages.docx"
    assert window.save_to(path)
    with zipfile.ZipFile(path) as archive:
        headers = [name for name in archive.namelist() if name.startswith("word/header")]
        assert any("PAGE" in archive.read(name).decode() for name in headers)
    window._page_numbers("remove")
    assert window.state.page.header == "Report"
    assert "{page}" not in window.state.page.header


def test_markup_filter_does_not_accept_or_delete_annotations(window):
    window.state.comments.append(Comment(0, 5, "Hello", "Note"))
    window.state.revisions.append(Revision(0, 5, "", "Hello", "", "Hello"))
    window._markup("comments", False)
    window._markup("changes", False)
    assert window.comments_list.count() == window.changes_list.count() == 0
    assert not window.editor.extraSelections()
    assert not window.state.comments[0].resolved
    assert window.state.revisions[0].status == "pending"
    window._markup("comments", True)
    window._markup("changes", True)
    assert window.comments_list.count() == window.changes_list.count() == 1


def test_thesaurus_replaces_utf16_selection_and_cannot_replace_new_document(window):
    window.editor.setPlainText("😀 good")
    cursor = window.editor.textCursor()
    cursor.setPosition(3)
    cursor.setPosition(7, QTextCursor.MoveMode.KeepAnchor)
    window.editor.setTextCursor(cursor)
    window._thesaurus()
    window.synonym_list.setCurrentRow(0)
    window._replace_synonym()
    assert window.editor.toPlainText() == "😀 excellent"
    window.load_state(DocumentState(html="<p>good</p>"), clean=True)
    window._replace_synonym()
    assert window.editor.toPlainText() == "good"


def test_zoom_facing_page_preview_opens_and_closes(window, monkeypatch, qapp):
    monkeypatch.setattr(authoring.ZoomDialog, "exec", lambda self: (
        self.mode.setCurrentText("Two Pages") or QDialog.DialogCode.Accepted))
    observed = []

    def close_preview():
        for widget in qapp.topLevelWidgets():
            if isinstance(widget, authoring.PagePreviewDialog):
                observed.append(widget.pages.columns)
                widget.accept()

    QTimer.singleShot(150, close_preview)
    window._zoom_dialog()
    assert observed == [2]


def test_paper_menu_actions_and_new_window_share_only_store_location(window, qapp):
    menus = window.findChildren(QMenu)
    action = next(action for menu in menus for action in menu.actions() if action.text() == "Letter")
    action.trigger()
    assert window.state.page.paper == "Letter"
    window._new_window()
    other = window._secondary_windows[0]
    assert other.editor.toPlainText() == ""
    assert other.state.id != window.state.id
    assert other.store.root == window.store.root
    other._dirty = False
    other.close()
    qapp.processEvents()
    window.store.set_setting("still_open", True)
    assert window.store.get_setting("still_open") is True


def test_read_aloud_uses_source_text_and_stops_on_close(editor, monkeypatch):
    from PySide6 import QtTextToSpeech

    class Speech:
        class State:
            Error = "error"

        @staticmethod
        def availableEngines():
            return ["sapi"]

        def __init__(self, engine, parent):
            self.engine = engine
            self.errorOccurred = type("Signal", (), {"connect": lambda self, fn: None})()
            self.spoken = ""
            self.stopped = False

        def state(self):
            return "ready"

        def say(self, text):
            self.spoken = text

        def stop(self):
            self.stopped = True

    monkeypatch.setattr(QtTextToSpeech, "QTextToSpeech", Speech)
    editor.setPlainText("Communication")
    reader = authoring.ReadingDialog(editor.document())
    reader.syllables.setChecked(True)
    reader.read_aloud()
    assert reader.speech.engine == "sapi"
    assert reader.speech.spoken == "Communication"
    reader.reject()
    assert reader.speech.stopped


def test_new_views_render(window, qapp):
    if not os.environ.get("FOLIO_SCREENSHOT"):
        pytest.skip("FOLIO_SCREENSHOT not set")
    artifacts = Path(__file__).resolve().parent.parent / "artifacts"
    artifacts.mkdir(exist_ok=True)
    window.resize(1440, 960)
    window.show()
    window.editor.setHtml("<h1>Reading with Folio</h1>" + "<p>"
                          + "Clear communication helps readers understand complex ideas. " * 16
                          + "</p><p>Use line focus and page color to make reading comfortable.</p>")
    cursor = window.editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    window.editor.setTextCursor(cursor)
    window.editor.insert_page_break()
    window.editor.textCursor().insertText("A second page for the multi-page preview.")
    window.editor.apply_page_settings(window.state.page)
    cursor.movePosition(QTextCursor.MoveOperation.Start)
    window.editor.setTextCursor(cursor)
    window._sync_all_panes()
    tabs = [window.ribbon.tabText(i) for i in range(window.ribbon.count())]
    window.ribbon.setCurrentIndex(tabs.index("Review"))
    qapp.processEvents()
    assert window.grab().save(str(artifacts / "folio-review-tools.png"))
    picker = authoring.StockMediaDialog(window)
    picker.show()
    qapp.processEvents()
    assert picker.grab().save(str(artifacts / "folio-stock-media.png"))
    picker.close()
    reader = authoring.ReadingDialog(window.editor.document(), window)
    reader.show()
    reader.color.setCurrentText("Sepia")
    reader.line_focus.setCurrentIndex(2)
    qapp.processEvents()
    assert reader.grab().save(str(artifacts / "folio-immersive-reader.png"))
    reader.close()
    publication = authoring.Publication(window.editor.document(), window.state.page, "Preview")
    preview = authoring.PagePreviewDialog(publication, 2, window)
    preview.show()
    qapp.processEvents()
    assert preview.grab().save(str(artifacts / "folio-multipage-preview.png"))
    preview.close()
