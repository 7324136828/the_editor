import base64
import json

import pytest
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QImage, QPainter, QTextCharFormat, QTextCursor

from folio.editor import SafeDocument
from folio.formatting import (
    FONT_SIZE_STEPS,
    MULTILEVEL_LIST_PROPERTY,
    paragraph_border,
    stepped_font_size,
)
from folio.models import DocumentState
from folio.review import ReviewTracker


def select(editor, start, end):
    cursor = editor.textCursor()
    cursor.setPosition(start)
    cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
    editor.setTextCursor(cursor)
    return cursor


def test_format_painter_is_idempotent_and_preserves_object_and_link(editor):
    editor.setHtml('<p><span style="font-size:18pt;font-weight:700;">source</span> '
                   '<a href="https://example.test">target</a></p>')
    select(editor, 0, 6)
    sample = editor.sample_inline_format()
    cursor = editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    editor.setTextCursor(cursor)
    image = QImage(8, 8, QImage.Format.Format_RGB32)
    image.fill(QColor("#123456"))
    editor.insert_image(image)
    select(editor, 7, editor.document().characterCount() - 1)
    text = editor.toPlainText()
    editor.apply_inline_format(sample)
    first = editor.document().toHtml()
    editor.apply_inline_format(sample)
    assert editor.document().toHtml() == first
    assert editor.toPlainText() == text
    target = select(editor, 7, 13).charFormat()
    assert target.anchorHref() == "https://example.test"
    assert target.fontPointSize() == 18
    assert target.fontWeight() == 700
    assert select(editor, 13, 14).charFormat().isImageFormat()


def test_painter_can_remove_prior_visual_attributes(editor):
    editor.setHtml('<p>plain <b><i><u>styled</u></i></b></p>')
    select(editor, 0, 5)
    sample = editor.sample_inline_format()
    select(editor, 6, 12)
    editor.apply_inline_format(sample)
    fmt = editor.textCursor().charFormat()
    assert fmt.fontWeight() < 700
    assert not fmt.fontItalic()
    assert not fmt.fontUnderline()


def test_script_toggles_are_exclusive_and_clearable(editor):
    editor.setPlainText("H2O")
    select(editor, 1, 2)
    editor.toggle_subscript()
    assert editor.textCursor().charFormat().verticalAlignment() == (
        QTextCharFormat.VerticalAlignment.AlignSubScript)
    editor.toggle_superscript()
    assert editor.textCursor().charFormat().verticalAlignment() == (
        QTextCharFormat.VerticalAlignment.AlignSuperScript)
    editor.toggle_superscript()
    assert editor.textCursor().charFormat().verticalAlignment() == (
        QTextCharFormat.VerticalAlignment.AlignNormal)


def test_scalar_font_steps_and_mixed_selection(editor):
    assert stepped_font_size(13, 1) == 14
    assert stepped_font_size(13, -1) == 12
    assert stepped_font_size(FONT_SIZE_STEPS[0], -1) == FONT_SIZE_STEPS[0]
    assert stepped_font_size(FONT_SIZE_STEPS[-1], 1) == FONT_SIZE_STEPS[-1]
    editor.setHtml('<p><span style="font-size:11pt">a</span>'
                   '<span style="font-size:20pt">b</span></p>')
    editor.selectAll()
    editor.step_font_size(1)
    assert select(editor, 0, 1).charFormat().fontPointSize() == 12
    assert select(editor, 1, 2).charFormat().fontPointSize() == 22


@pytest.mark.parametrize("mode, expected", [
    ("sentence", "Hello world. Next!"),
    ("lower", "hello world. next!"),
    ("upper", "HELLO WORLD. NEXT!"),
    ("title", "Hello World. Next!"),
])
def test_case_modes_preserve_runs_and_undo(editor, mode, expected):
    editor.setHtml('<p><b>hELLO</b> wORLD. nEXT!</p>')
    editor.selectAll()
    editor.change_case(mode)
    assert editor.toPlainText() == expected
    assert select(editor, 0, 5).charFormat().fontWeight() == 700
    assert select(editor, 6, 11).charFormat().fontWeight() < 700
    editor.undo()
    assert editor.toPlainText() == "hELLO wORLD. nEXT!"


def test_unicode_case_expansion_rebases_unchanged_comment_and_preserves_link(editor):
    editor.setHtml('<p>\U0001f600<a href="https://example.test">stra\u00dfe</a> 123</p>')
    state = DocumentState()
    tracker = ReviewTracker(editor.document(), state)
    comment = tracker.add_comment(9, 12, "number")
    editor.selectAll()
    editor.change_case("upper")
    assert editor.toPlainText() == "\U0001f600STRASSE 123"
    assert (comment.start, comment.end, comment.orphaned) == (10, 13, False)
    assert select(editor, 2, 9).charFormat().anchorHref() == "https://example.test"


def test_lowercase_preserves_contextual_greek_final_sigma(editor):
    editor.setPlainText("\u039f\u03a3")
    editor.selectAll()
    editor.change_case("lower")
    assert editor.toPlainText() == "\u03bf\u03c2"


def test_multilevel_lists_reindex_after_indent_delete_and_outdent(editor):
    editor.setPlainText("A\nB\nC\nD")
    editor.selectAll()
    editor.set_multilevel_list()
    select(editor, 2, 5)
    editor.indent_paragraph(1)
    a = editor.document().findBlockByNumber(0)
    b = a.next()
    c = b.next()
    d = c.next()
    assert b.textList() == c.textList()
    assert b.textList().format().indent() == 2
    assert b.textList().itemText(b) == "a."
    assert b.textList().itemText(c) == "b."
    assert a.textList() == d.textList()
    assert a.textList().itemText(d) == "2."
    select(editor, 2, 4).removeSelectedText()
    c = editor.document().findBlockByNumber(1)
    assert c.text() == "C"
    assert c.textList().itemText(c) == "a."
    select(editor, c.position(), c.position())
    editor.indent_paragraph(-1)
    assert c.textList() == a.textList()
    assert c.textList().itemText(c) == "2."
    assert d.textList().itemText(d) == "3."


def test_numbering_continues_when_adjacent_paragraph_added(editor):
    editor.setPlainText("A\nB\nC")
    select(editor, 0, 3)
    editor.set_list(True)
    select(editor, 4, 5)
    editor.set_list(True)
    first = editor.document().begin()
    last = editor.document().lastBlock()
    assert first.textList() == last.textList()
    assert last.textList().itemText(last) == "3."


def test_borders_shading_and_lists_survive_html_without_extra_text(editor):
    editor.setPlainText("first\nsecond")
    select(editor, 0, 6)
    editor.set_paragraph_border("box", "#c02030", 2)
    editor.set_paragraph_shading(QColor("#f5e8aa"))
    editor.set_multilevel_list()
    source = editor.document().toHtml()
    restored = SafeDocument()
    restored.setHtml(source)
    first = restored.begin()
    second = first.next()
    assert restored.toPlainText() == "first\nsecond"
    assert paragraph_border(first) == {
        "sides": ["top", "bottom", "left", "right"],
        "color": "#c02030", "width": 2.0,
    }
    assert paragraph_border(second) is None
    assert first.blockFormat().background().color().name() == "#f5e8aa"
    assert second.blockFormat().background().style() == Qt.BrushStyle.NoBrush
    assert first.textList().format().property(MULTILEVEL_LIST_PROPERTY)


def test_paragraph_border_render_and_undo(editor):
    editor.setPlainText("Border")
    editor.set_paragraph_border("box", "#ff0000", 3)
    doc = editor.document()
    doc.setTextWidth(240)
    image = QImage(300, 300, QImage.Format.Format_RGB32)
    image.fill(QColor("white"))
    painter = QPainter(image)
    doc.drawContents(painter)
    painter.end()
    assert any(image.pixelColor(x, y) == QColor("#ff0000")
               for y in range(image.height()) for x in range(image.width()))
    editor.undo()
    assert paragraph_border(doc.begin()) is None


def test_invalid_border_metadata_ignored(qapp):
    payload = {"blocks": [{"block": 0, "border": {
        "sides": ["top"], "color": "#000000", "width": -100}}]}
    encoded = base64.b64encode(json.dumps(payload).encode()).decode()
    doc = SafeDocument()
    doc.setHtml("<p>content</p><!--folio-block-formats:" + encoded + "-->")
    assert doc.toPlainText() == "content"
    assert paragraph_border(doc.begin()) is None
