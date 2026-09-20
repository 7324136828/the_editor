import pytest
from PySide6.QtCore import QMimeData, Qt
from PySide6.QtGui import QColor, QImage, QTextCharFormat, QTextCursor, QTextFormat


def _select(editor, start, end):
    cursor = editor.textCursor()
    cursor.setPosition(start)
    cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
    editor.setTextCursor(cursor)
    return cursor


def test_bold_italic_underline_toggle(editor):
    editor.setPlainText("hello world")
    _select(editor, 0, 5)
    editor.toggle_bold()
    fmt = _select(editor, 0, 5).charFormat()
    assert fmt.fontWeight() >= 700
    editor.toggle_italic()
    editor.toggle_underline()
    fmt = _select(editor, 0, 5).charFormat()
    assert fmt.fontItalic()
    assert fmt.fontUnderline()
    editor.toggle_strikethrough()
    assert _select(editor, 0, 5).charFormat().fontStrikeOut()


def test_apply_style_and_headings(editor):
    editor.setPlainText("Intro\nBody")
    _select(editor, 0, 5)
    editor.apply_style("Heading 1")
    headings = editor.headings()
    assert len(headings) == 1
    pos, level, text = headings[0]
    assert level == 1 and text == "Intro"
    assert editor.current_style_name() == "Heading 1"
    editor.apply_style("Normal")
    assert editor.headings() == []


def test_unknown_style_rejected(editor):
    with pytest.raises(ValueError):
        editor.apply_style("No Such Style")


def test_set_highlight_and_clear(editor):
    editor.setPlainText("highlight me")
    _select(editor, 0, 9)
    editor.set_highlight(QColor("#ffff00"))
    fmt = _select(editor, 0, 9).charFormat()
    assert fmt.background().color().name() == "#ffff00"
    editor.set_highlight(None)
    fmt = _select(editor, 0, 9).charFormat()
    assert fmt.background().style() == Qt.BrushStyle.NoBrush


def test_clear_formatting_resets_selection_only(editor):
    editor.setPlainText("bold text plain")
    _select(editor, 0, 9)
    editor.toggle_bold()
    editor.set_highlight(QColor("#ffff00"))
    _select(editor, 0, 9)
    editor.clear_formatting()
    fmt = _select(editor, 0, 9).charFormat()
    assert fmt.fontWeight() < 700
    assert fmt.background().style() == Qt.BrushStyle.NoBrush
    assert editor.document().toPlainText() == "bold text plain"


def test_lists_toggle(editor):
    editor.setPlainText("one\ntwo")
    _select(editor, 0, 7)
    editor.set_list(True)
    block = editor.document().begin()
    assert block.textList() is not None
    editor.set_list(True)
    assert editor.document().begin().textList() is None


def test_insert_merge_field_validation(editor):
    editor.insert_merge_field("FirstName")
    assert "{{FirstName}}" in editor.document().toPlainText()
    for bad in ("", "  ", "a b", "x{y}", "line\nbreak", "9start"):
        with pytest.raises(ValueError):
            editor.insert_merge_field(bad)


def test_insert_link_scheme_validation(editor):
    editor.insert_link("https://example.com", "site")
    assert "site" in editor.document().toPlainText()
    for bad in ("javascript:alert(1)", "file:///etc/passwd", "ftp://x"):
        with pytest.raises(ValueError):
            editor.insert_link(bad)


def test_insert_table_and_ops(editor):
    editor.setPlainText("cell host")
    table = editor.insert_table(2, 3)
    assert table.rows() == 2 and table.columns() == 3
    cursor = table.cellAt(0, 0).firstCursorPosition()
    editor.setTextCursor(cursor)
    assert editor.current_table() is not None
    assert editor.table_insert_row()
    assert editor.current_table().rows() == 3
    assert editor.table_insert_column()
    assert editor.current_table().columns() == 4
    assert editor.table_delete_row()
    editor.setTextCursor(table.cellAt(0, 0).firstCursorPosition())
    assert editor.table_delete_column()


def test_image_insert_and_resize_ratio(editor):
    image = QImage(200, 100, QImage.Format.Format_RGB32)
    image.fill(QColor("#336699"))
    editor.insert_image(image, width=100.0)
    fmt = editor.selected_image_format()
    if fmt is None:
        cursor = editor.textCursor()
        cursor.setPosition(0)
        editor.setTextCursor(cursor)
        fmt = editor.selected_image_format()
    assert fmt is not None
    assert fmt.width() == pytest.approx(100.0)
    assert fmt.height() == pytest.approx(50.0)
    assert editor.resize_selected_image(50.0)
    fmt = editor.selected_image_format()
    assert fmt.width() == pytest.approx(50.0)
    assert fmt.height() == pytest.approx(25.0)


def test_page_break_sets_policy(editor):
    editor.setPlainText("before")
    cursor = editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    editor.setTextCursor(cursor)
    editor.insert_page_break()
    found = False
    b = editor.document().begin()
    while b.isValid():
        if b.blockFormat().pageBreakPolicy() == QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore:
            found = True
        b = b.next()
    assert found


def test_column_break_is_distinct_and_survives_html(editor):
    from folio.editor import block_break_kind

    editor.setPlainText("before")
    cursor = editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    editor.setTextCursor(cursor)
    editor.insert_column_break()
    cursor = editor.textCursor()
    cursor.insertText("after")
    block = editor.document().findBlock(cursor.position())
    assert block_break_kind(block) == "column"
    html = editor.document().toHtml()
    editor.document().setHtml(html)
    block = editor.document().findBlock(editor.document().toPlainText().find("after"))
    assert block_break_kind(block) == "column"


def test_manual_hyphenation_inserts_discretionary_hyphens(editor):
    editor.setPlainText("internationalization SHORTWORD UPPERCASELONGWORD")
    count = editor.apply_manual_hyphenation(False)
    assert count == 1
    assert "\u00ad" in editor.toPlainText()
    assert "UPPERCASELONGWORD" in editor.toPlainText()


def test_paste_special_plain_text_discards_html(editor):
    mime = QMimeData()
    mime.setHtml("<b>formatted</b>")
    mime.setText("formatted")
    editor.insert_mime_data(mime, "text")
    cursor = editor.textCursor()
    cursor.select(QTextCursor.SelectionType.Document)
    assert cursor.selectedText() == "formatted"
    assert cursor.charFormat().fontWeight() < 700


def test_page_settings_applied(editor):
    from folio.models import PageSettings
    settings = PageSettings(paper="Letter", landscape=True)
    editor.apply_page_settings(settings)
    size = editor.document().pageSize()
    assert size.width() > size.height()


def test_object_marker_is_fffc():
    from folio.editor import OBJECT_REPLACEMENT
    assert OBJECT_REPLACEMENT == "￼"
    assert ord(OBJECT_REPLACEMENT) == 0xFFFC


def test_clear_formatting_preserves_images(editor):
    image = QImage(40, 30, QImage.Format.Format_RGB32)
    image.fill(QColor("#224466"))
    editor.setPlainText("before")
    editor.insert_image(image, width=40.0)
    cursor = editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    cursor.insertText("after")
    doc = editor.document()
    assert "before" in doc.toPlainText() and "after" in doc.toPlainText()
    editor.selectAll()
    editor.clear_formatting()
    plain = doc.toPlainText()
    assert "before" in plain and "after" in plain
    block = doc.begin()
    found_image = False
    while block.isValid():
        it = block.begin()
        while not it.atEnd():
            frag = it.fragment()
            if frag.isValid() and frag.charFormat().isImageFormat():
                found_image = True
            it += 1
        block = block.next()
    assert found_image


def test_block_style_excludes_next_block_at_selection_edge(editor):
    editor.setPlainText("first\nsecond")
    _select(editor, 0, 6)
    editor.apply_style("Heading 1")
    doc = editor.document()
    first = doc.begin()
    second = first.next()
    assert first.blockFormat().headingLevel() == 1
    assert second.blockFormat().headingLevel() == 0


def test_page_break_not_propagated_by_enter(editor):
    from PySide6.QtTest import QTest
    editor.setPlainText("AB")
    cursor = editor.textCursor()
    cursor.setPosition(1)
    editor.setTextCursor(cursor)
    editor.insert_page_break()
    QTest.keyClick(editor, Qt.Key.Key_Return)
    doc = editor.document()
    breaks = 0
    block = doc.begin()
    while block.isValid():
        if block.blockFormat().pageBreakPolicy() == (
                QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore):
            breaks += 1
        block = block.next()
    assert breaks == 1


def test_safe_document_rejects_external_resources(qapp):
    from folio.editor import SafeDocument

    class SpyDocument(SafeDocument):
        def __init__(self):
            super().__init__()
            self.results = []

        def loadResource(self, resource_type, url):
            result = super().loadResource(resource_type, url)
            self.results.append((url.toString(), result))
            return result

    doc = SpyDocument()
    doc.setHtml('<p><img src="file:///C:/Windows/notepad.exe"/></p>'
                '<p><img src="https://example.invalid/x.png"/></p>')
    doc.pageCount()
    doc.adjustSize()
    assert doc.results
    assert all(result is None for _url, result in doc.results)


def test_safe_document_embedded_image_roundtrip(qapp):
    from folio.editor import SafeDocument, _image_to_data_uri
    image = QImage(20, 20, QImage.Format.Format_RGB32)
    image.fill(QColor("#112233"))
    uri = _image_to_data_uri(image)
    doc = SafeDocument()
    doc.setHtml(f'<p><img src="{uri}" width="20" height="20"/></p>')
    out = doc.toHtml()
    assert "data:image" in out or "￼" in doc.toPlainText()
