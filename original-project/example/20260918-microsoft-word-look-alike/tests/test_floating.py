import pytest
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QImage, QPainter, QTextCharFormat, QTextCursor

from folio.editor import SafeDocument, _image_to_data_uri
from folio.floating import FloatingLayout
from folio.models import PageSettings


def body(qapp, text=None):
    document = SafeDocument()
    document.setDefaultFont(QFont("Calibri", 11))
    document.setPlainText(text or "The quick brown fox jumps over the lazy dog. " * 28)
    return document


def square(**values):
    return {"id": "shape", "kind": "text", "page": 0, "x": 270, "y": 96,
            "width": 190, "height": 120, "wrap": "square", **values}


def test_square_wrap_splits_actual_lines_and_keeps_all_utf16_text(qapp):
    document = body(qapp, "Hello \U0001f600. " * 100)
    original = document.toHtml()
    layout = FloatingLayout(document, PageSettings(), [square()])
    obstruction = QRectF(264, 90, 202, 132)
    assert all(not line.rect.intersects(obstruction) for line in layout.lines)
    beside = [line for line in layout.lines if line.rect.top() < 216]
    assert any(line.rect.right() < 270 for line in beside)
    assert any(line.rect.left() > 460 for line in beside)
    assert layout.lines[0].start == 0
    assert layout.lines[-1].end == document.characterCount() - 1
    assert all(previous.end == current.start for previous, current in zip(layout.lines, layout.lines[1:]))
    assert document.toHtml() == original


def test_tight_wrap_uses_alpha_outline_instead_of_image_rectangle(qapp):
    image = QImage(100, 100, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    painter.fillRect(QRectF(40, 0, 20, 100), QColor("#123456"))
    painter.end()
    obj = square(kind="image", wrap="tight", image=_image_to_data_uri(image))
    layout = FloatingLayout(body(qapp), PageSettings(), [obj])
    interval = layout.excluded_intervals(0, 100, 15)[0]
    assert interval[0] == pytest.approx(340)
    assert interval[1] == pytest.approx(390)
    assert any(270 < line.rect.right() < 346 for line in layout.lines[:4])
    square_layout = FloatingLayout(body(qapp), PageSettings(), [{**obj, "wrap": "square"}])
    assert layout.lines[0].end > square_layout.lines[0].end


def test_hidden_and_overlay_objects_do_not_displace_text(qapp):
    document = body(qapp)
    baseline = FloatingLayout(document, PageSettings(), [])
    for obj in (square(visible=False), square(wrap="front"), square(wrap="behind")):
        layout = FloatingLayout(document, PageSettings(), [obj])
        assert [(line.start, line.end, line.rect) for line in layout.lines] == [
            (line.start, line.end, line.rect) for line in baseline.lines]


def test_full_width_obstacle_moves_text_below_object(qapp):
    layout = FloatingLayout(body(qapp), PageSettings(),
                            [square(x=0, width=900, height=100)])
    assert layout.lines[0].rect.top() >= 202


def test_wrapped_lines_keep_inline_formatting_and_images(qapp):
    document = body(qapp, "Bold image ")
    cursor = QTextCursor(document)
    cursor.setPosition(0)
    cursor.setPosition(4, QTextCursor.MoveMode.KeepAnchor)
    fmt = QTextCharFormat()
    fmt.setFontWeight(QFont.Weight.Bold)
    fmt.setForeground(QColor("#ff0000"))
    cursor.mergeCharFormat(fmt)
    cursor.movePosition(QTextCursor.MoveOperation.End)
    image = QImage(35, 30, QImage.Format.Format_ARGB32)
    image.fill(QColor("#00ff00"))
    cursor.insertHtml('<img width="35" height="30" src="' + _image_to_data_uri(image) + '">')
    cursor.insertText(" after image " * 20)
    layout = FloatingLayout(document, PageSettings(), [square()])
    first = layout.lines[0].document.firstBlock().begin().fragment().charFormat()
    assert first.fontWeight() == QFont.Weight.Bold
    assert first.foreground().color().name() == "#ff0000"
    assert any("<img" in line.document.toHtml() for line in layout.lines)
    target = QImage(794, 1123, QImage.Format.Format_ARGB32)
    target.fill(QColor("#ffffff"))
    painter = QPainter(target)
    layout.draw_page(painter, 0)
    painter.end()
    assert any(target.pixelColor(x, y).name() == "#00ff00"
               for y in range(90, 200) for x in range(90, 260))


def test_tables_remain_structured_and_are_painted_without_mutating_source(editor):
    editor.setPlainText("Before")
    cursor = editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    cursor.insertBlock()
    table = cursor.insertTable(2, 2)
    table.cellAt(0, 0).firstCursorPosition().insertText("Cell one")
    table.cellAt(0, 1).firstCursorPosition().insertText("Cell two")
    table.cellAt(1, 0).firstCursorPosition().insertText("Cell three")
    table.cellAt(1, 1).firstCursorPosition().insertText("Cell four")
    before = editor.document().toHtml()
    layout = FloatingLayout(editor.document(), PageSettings(), [square()])
    tables = [line for line in layout.lines if "<table" in line.document.toHtml()]
    assert tables
    assert all("Cell one" in line.document.toHtml() and "Cell four" in line.document.toHtml()
               for line in tables)
    assert editor.document().toHtml() == before


def test_multicolumn_pages_and_position_mapping(qapp):
    document = body(qapp, "Content with several words on each line. " * 500)
    layout = FloatingLayout(document, PageSettings(columns=2), [square(width=90)])
    assert layout.page_count() > 1
    assert {line.column for line in layout.lines} == {0, 1}
    for line in layout.lines:
        assert layout.page_for_position(line.start) == line.page
        assert line.rect.bottom() <= layout.bottom + 0.01
    assert layout.lines[-1].end == document.characterCount() - 1


def test_painting_rejects_invalid_page(qapp):
    layout = FloatingLayout(body(qapp), PageSettings(), [])
    image = QImage(10, 10, QImage.Format.Format_ARGB32)
    painter = QPainter(image)
    with pytest.raises(IndexError):
        layout.draw_page(painter, -1)
    painter.end()
