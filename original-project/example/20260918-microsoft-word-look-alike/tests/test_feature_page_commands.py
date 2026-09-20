from PySide6.QtGui import QTextCursor

from folio.document_features import insert_section
from folio.models import DocumentState, PageSettings
from folio.publishing import Publication
from folio.window import FolioWindow


def test_active_section_geometry_and_page_number_aliases(qapp, store):
    window = FolioWindow(store=store, restore=False)
    try:
        window.load_state(DocumentState(html="<p>Original</p>"), clean=True)
        cursor = window.editor.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        window.editor.setTextCursor(cursor)
        key = insert_section(window.editor, window.state, "next", PageSettings())
        window.editor.textCursor().insertText("Section body")
        window._quick_orientation(True)
        window._quick_columns(2)
        assert not window.state.page.landscape
        section = next(s for s in window.state.features["sections"] if s["id"] == key)
        assert section["page"]["landscape"]
        assert section["page"]["columns"] == 2
        page = window._active_page_settings()
        page.header = "Report Page {CurrentPage} of {TotalPages}"
        window._set_active_page(page)
        window._page_numbers("top")
        assert window._active_page_settings().header == "Report Page {CurrentPage} of {TotalPages}"
        window._page_numbers("remove")
        assert window._active_page_settings().header == "Report"
    finally:
        window._dirty = False
        window.close()


def test_blank_page_places_cursor_in_the_empty_page(qapp, store):
    window = FolioWindow(store=store, restore=False)
    try:
        window.load_state(DocumentState(html="<p>Before</p>"), clean=True)
        cursor = window.editor.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        window.editor.setTextCursor(cursor)
        window._blank_page()
        assert window.editor.textCursor().block().text() == ""
        assert window.editor.textCursor().block().next().isValid()
        publication = Publication(window.editor.document(), window.state.page, "Blank")
        assert publication.page_count() == 3
    finally:
        window._dirty = False
        window.close()
