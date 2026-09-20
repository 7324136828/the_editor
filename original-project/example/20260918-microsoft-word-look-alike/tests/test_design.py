import pytest
from PySide6.QtCore import QRectF
from PySide6.QtGui import QColor, QFont, QImage, QPainter, QTextCharFormat, QTextCursor

from folio.design import (
    DesignSettings, STYLE_SETS, apply_design, bind_style, capture_style_bindings,
    paint_page_decoration, resolved_style, restore_style_bindings, style_name,
    theme_shape_colors, theme_tokens,
)
from folio.editor import SafeDocument, _image_to_data_uri


def select(document, start, end):
    cursor = QTextCursor(document)
    cursor.setPosition(start)
    cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
    return cursor


def char_at(document, position):
    return select(document, position, position + 1).charFormat()


def test_semantic_style_set_switch_preserves_title_and_outline(qapp):
    document = SafeDocument()
    document.setPlainText("A title\nAn outline\nA body")
    bind_style(QTextCursor(document.firstBlock()), "Title")
    bind_style(QTextCursor(document.findBlockByNumber(1)), "Heading 1")
    bind_style(QTextCursor(document.lastBlock()), "Normal")
    apply_design(document, DesignSettings(theme="Woodland", style_set="Compact"))
    assert style_name(document.firstBlock()) == "Title"
    assert char_at(document, 0).fontPointSize() == 26
    heading = document.findBlockByNumber(1)
    assert heading.blockFormat().headingLevel() == 1
    assert char_at(document, heading.position()).fontPointSize() == 18
    assert char_at(document, heading.position()).foreground().color().name() == "#166534"
    assert char_at(document, document.lastBlock().position()).fontPointSize() == 10


def test_theme_change_preserves_local_overrides_and_link_targets(qapp):
    document = SafeDocument()
    document.setPlainText("plain local linked")
    bind_style(QTextCursor(document), "Normal")
    custom = QTextCharFormat()
    custom.setFontWeight(QFont.Weight.Bold)
    custom.setForeground(QColor("#ff0000"))
    custom.setFontPointSize(19)
    select(document, 6, 11).mergeCharFormat(custom)
    link = QTextCharFormat()
    link.setAnchor(True)
    link.setAnchorHref("https://example.com")
    select(document, 12, 18).mergeCharFormat(link)
    apply_design(document, DesignSettings(theme="Editorial", style_set="Large Print"))
    assert char_at(document, 0).fontPointSize() == 16
    assert char_at(document, 0).fontFamilies() == ["Georgia"]
    assert char_at(document, 6).fontPointSize() == 19
    assert char_at(document, 6).fontWeight() == QFont.Weight.Bold
    assert char_at(document, 6).foreground().color().name() == "#ff0000"
    assert char_at(document, 12).anchorHref() == "https://example.com"
    assert char_at(document, 12).fontPointSize() == 16
    apply_design(document, DesignSettings())
    assert char_at(document, 0).fontPointSize() == 11
    assert char_at(document, 6).fontPointSize() == 19


def test_spacing_policy_is_reversible_and_retains_paragraph_override(qapp):
    document = SafeDocument()
    document.setPlainText("Heading\nLocal")
    bind_style(QTextCursor(document.firstBlock()), "Heading 1")
    bind_style(QTextCursor(document.lastBlock()), "Normal")
    cursor = QTextCursor(document.lastBlock())
    fmt = cursor.blockFormat()
    fmt.setBottomMargin(73)
    cursor.setBlockFormat(fmt)
    original = STYLE_SETS["Modern"]["Heading 1"]
    apply_design(document, DesignSettings(paragraph_spacing="None"))
    assert document.firstBlock().blockFormat().topMargin() == 0
    assert document.firstBlock().blockFormat().bottomMargin() == 0
    assert document.lastBlock().blockFormat().bottomMargin() == 73
    assert STYLE_SETS["Modern"]["Heading 1"] == original
    apply_design(document, DesignSettings())
    assert document.firstBlock().blockFormat().topMargin() == pytest.approx(20 * 96 / 72)
    assert document.firstBlock().blockFormat().bottomMargin() == pytest.approx(10 * 96 / 72)
    assert document.lastBlock().blockFormat().bottomMargin() == 73


def test_bindings_round_trip_across_html_utf16(qapp):
    document = SafeDocument()
    document.setPlainText("Emoji \U0001f600\nSubtitle")
    bind_style(QTextCursor(document.firstBlock()), "Title")
    bind_style(QTextCursor(document.lastBlock()), "Subtitle")
    records = capture_style_bindings(document)
    assert records[1]["position"] == 9
    restored = SafeDocument()
    restored.setHtml(document.toHtml())
    restore_style_bindings(restored, records)
    assert capture_style_bindings(restored) == records
    apply_design(restored, DesignSettings(style_set="Large Print"))
    assert style_name(restored.lastBlock()) == "Subtitle"
    assert char_at(restored, restored.lastBlock().position()).fontPointSize() == 22


def test_binding_restore_validates_all_records_before_mutating(qapp):
    document = SafeDocument()
    document.setPlainText("first\nsecond")
    bind_style(QTextCursor(document.firstBlock()), "Title")
    records = capture_style_bindings(document)
    restored = SafeDocument()
    restored.setPlainText(document.toPlainText())
    with pytest.raises(ValueError):
        restore_style_bindings(restored, [*records, {**records[0], "position": 1}])
    assert capture_style_bindings(restored) == []


def test_style_selection_excludes_next_block_when_ending_at_start(qapp):
    document = SafeDocument()
    document.setPlainText("first\nsecond")
    bind_style(select(document, 0, 6), "Heading 2")
    assert style_name(document.firstBlock()) == "Heading 2"
    assert style_name(document.lastBlock()) == "Normal"


def test_design_undo_restores_semantic_binding_and_typography(qapp):
    document = SafeDocument()
    document.setPlainText("Title")
    bind_style(QTextCursor(document), "Title")
    apply_design(document, DesignSettings(style_set="Compact"))
    assert char_at(document, 0).fontPointSize() == 26
    document.undo()
    assert char_at(document, 0).fontPointSize() == 32
    assert capture_style_bindings(document)[0]["baseline"]["size"] == 32


def test_theme_tokens_and_shape_shading_follow_theme(qapp):
    office = theme_shape_colors(DesignSettings())
    forest = theme_shape_colors(DesignSettings(theme="Woodland"))
    assert office != forest
    assert forest["stroke"] == "#166534"
    assert QColor(forest["fill"]).lightness() > QColor(forest["stroke"]).lightness()
    custom = DesignSettings(theme="Woodland", color_palette="Plum",
                            font_pairing="Georgia / Georgia")
    assert theme_tokens(custom)["heading_font"] == "Georgia"
    assert resolved_style("Title", custom)["color"] == "#6b21a8"


@pytest.mark.parametrize("values", [
    {"theme": "missing"}, {"style_set": "missing"}, {"page_color": "invalid"},
    {"watermark_opacity": float("nan")}, {"page_border_width_pt": True},
    {"watermark_image": "https://example.com/image.png"},
    {"watermark_image": "file:///C:/secret.png"}, {"watermark_kind": "image"},
    {"watermark_angle": 181}, {"font_pairing": 4},
])
def test_design_settings_reject_invalid_input(values):
    with pytest.raises(ValueError):
        DesignSettings.from_dict(values)


def test_page_decorations_render_background_border_and_separate_watermark_layers(qapp):
    settings = DesignSettings(page_color="#fef3c7", page_border="box",
                              page_border_color="#ff0000", page_border_inset_mm=0,
                              watermark_kind="text", watermark_text="DRAFT",
                              watermark_opacity=0.5, watermark_angle=0,
                              watermark_layer="front")
    image = QImage(400, 500, QImage.Format.Format_ARGB32)
    image.fill(QColor("#ffffff"))
    painter = QPainter(image)
    rect = QRectF(0, 0, 400, 500)
    paint_page_decoration(painter, rect, settings, "behind")
    assert image.pixelColor(200, 250).name() == "#fef3c7"
    background = image.copy()
    paint_page_decoration(painter, rect, settings, "front")
    painter.end()
    assert image.pixelColor(0, 0).red() > image.pixelColor(0, 0).green()
    assert image.copy(100, 200, 200, 100) != background.copy(100, 200, 200, 100)
    assert DesignSettings.from_dict(settings.to_dict()) == settings


def test_embedded_image_watermark_can_be_painted_without_file_access(qapp):
    stamp = QImage(8, 8, QImage.Format.Format_ARGB32)
    stamp.fill(QColor("#0000ff"))
    settings = DesignSettings(watermark_kind="image", watermark_opacity=0.5,
                              watermark_angle=0, watermark_image=_image_to_data_uri(stamp))
    image = QImage(100, 100, QImage.Format.Format_ARGB32)
    painter = QPainter(image)
    paint_page_decoration(painter, QRectF(0, 0, 100, 100), settings)
    painter.end()
    assert image.pixelColor(50, 50).blue() == 255
    assert 120 < image.pixelColor(50, 50).red() < 135
    assert image.pixelColor(0, 0).name() == "#ffffff"
