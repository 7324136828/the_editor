import copy

import pytest
from PySide6.QtCore import QRectF
from PySide6.QtGui import QColor, QFont, QImage, QPainter, QTextCursor
from PySide6.QtPdf import QPdfDocument

from folio.authoring import PagePreview
from folio.design import DesignSettings
from folio.document_features import insert_section
from folio.editor import SafeDocument
from folio.models import DocumentState, PageSettings
from folio.publishing import Publication, write_pdf


def content(qapp, text):
    document = SafeDocument()
    document.setDefaultFont(QFont("Calibri", 11))
    document.setPlainText(text)
    return document


def section_document(editor, kind, second_settings, before="First section", after="Second section"):
    state = DocumentState()
    editor.setPlainText(before)
    cursor = editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    editor.setTextCursor(cursor)
    insert_section(editor, state, kind, second_settings)
    editor.textCursor().insertText(after)
    editor.document().folio_state = state
    return state


def test_page_lookup_tracks_lines_inside_a_single_long_paragraph(qapp):
    document = content(qapp, "Words inside a long paragraph. " * 800)
    publication = Publication(document, PageSettings(), "Long paragraph")
    assert publication.page_count() > 2
    assert publication.page_for_position(document.characterCount() - 2) == publication.page_count()
    for page in range(1, publication.page_count()):
        position = publication.first_position_on_page(page)
        assert publication.page_for_position(position) == page + 1


def test_long_document_balances_only_last_page_without_changing_text(qapp, tmp_path):
    document = content(qapp, "Balanced content with enough text for several pages. " * 450)
    before = document.toHtml()
    publication = Publication(document, PageSettings(columns=2), "Balanced")
    assert publication.page_count() > 1
    assert publication._tail is not None
    tail = publication._tail[2]
    locations = tail._line_locations()
    heights = [max(y for page, slot, y in locations if slot == column) for column in (0, 1)]
    assert abs(heights[0] - heights[1]) <= 40
    assert tail.body_h < publication.body_h
    assert publication.page_for_position(document.characterCount() - 2) == publication.page_count()
    path = tmp_path / "balanced.pdf"
    count = write_pdf(path, document, PageSettings(columns=2), "Balanced")
    reader = QPdfDocument()
    reader.load(str(path))
    text = "".join(reader.getAllText(page).text() for page in range(count))
    assert text.count("Balanced") == 450
    assert document.toHtml() == before


@pytest.mark.parametrize("kind, expected", [("next", 2), ("even", 2), ("odd", 3)])
def test_section_break_parity_and_mixed_geometry(editor, kind, expected):
    state = section_document(editor, kind, PageSettings(landscape=True, paper="Letter"))
    publication = Publication(editor.document(), state.page, "Sections", state)
    assert publication.page_count() == expected
    assert publication.page_size(0)[0] < publication.page_size(0)[1]
    assert publication.page_size(expected - 1)[0] > publication.page_size(expected - 1)[1]
    assert publication.page_for_position(editor.document().characterCount() - 2) == expected


def test_continuous_section_shares_page_and_moves_next_columns_below_body(editor):
    state = section_document(editor, "continuous", PageSettings(columns=2),
                             after="Second section body " * 20)
    publication = Publication(editor.document(), state.page, "Continuous", state)
    assert publication.page_count() == 1
    assert publication._sections[1][3] > 0
    assert publication.page_for_position(editor.document().characterCount() - 2) == 1


def test_long_continuous_section_uses_remaining_space_then_full_pages(editor, tmp_path):
    state = section_document(editor, "continuous", PageSettings(),
                             before="First section paragraph.\n" * 20,
                             after="Second section paragraph. " * 1500)
    publication = Publication(editor.document(), state.page, "Continuous", state)
    assert publication.page_count() > 2
    assert publication._sections[1][3] > 0
    assert publication._sections[-1][3] == 0
    assert publication._sections[-1][4].body_h == pytest.approx(publication.body_h)
    path = tmp_path / "continuous.pdf"
    pages = write_pdf(path, editor.document(), state.page, "Continuous")
    reader = QPdfDocument()
    reader.load(str(path))
    text = "".join(reader.getAllText(page).text() for page in range(pages))
    assert text.count("Second") == 1500


def test_mixed_paper_size_pdf_and_preview_use_each_page_geometry(editor, tmp_path):
    state = section_document(editor, "next", PageSettings(landscape=True, paper="Letter"))
    publication = Publication(editor.document(), state.page, "Mixed", state)
    preview = PagePreview(publication, 2)
    assert preview.page_rects[0].width() < preview.page_rects[0].height()
    assert preview.page_rects[1].width() > preview.page_rects[1].height()
    assert preview.page_rects[0].right() < preview.page_rects[1].left()
    path = tmp_path / "mixed.pdf"
    assert write_pdf(path, editor.document(), state.page, "Mixed") == 2
    reader = QPdfDocument()
    reader.load(str(path))
    assert reader.pagePointSize(0).width() == pytest.approx(595, abs=2)
    assert reader.pagePointSize(1).width() == pytest.approx(792, abs=2)
    assert reader.pagePointSize(1).height() == pytest.approx(612, abs=2)
    assert "Second" in reader.getAllText(1).text()
    preview.deleteLater()


def test_later_floating_object_extends_output_without_extending_section_bodies(editor):
    state = section_document(editor, "next", PageSettings())
    state.features["objects"] = [{"id": "later", "kind": "text", "page": 4,
                                   "x": 96, "y": 96, "width": 200, "height": 60,
                                   "wrap": "front", "text": "Last page object"}]
    publication = Publication(editor.document(), state.page, "Later", state)
    assert publication.page_count() == 5
    assert len(publication._physical_settings) == 2
    assert publication.page_for_position(editor.document().characterCount() - 2) == 2
    image = QImage(794, 1123, QImage.Format.Format_ARGB32)
    painter = QPainter(image)
    publication.paint_page(painter, 4)
    painter.end()
    assert image.pixelColor(100, 100).name() == "#e0ecff"


def test_wrapping_survives_section_composition_and_maps_one_based_pages(editor):
    state = section_document(editor, "next", PageSettings(), after="Wrapped section text. " * 100)
    state.features["objects"] = [{"id": "shape", "kind": "text", "page": 1,
                                   "x": 270, "y": 96, "width": 190, "height": 120,
                                   "wrap": "square", "text": "Object"}]
    publication = Publication(editor.document(), state.page, "Wrapped", state)
    child = publication._sections[1][4]
    assert child._flow is not None
    obstacle = QRectF(264, 90, 202, 132)
    assert all(not line.rect.intersects(obstacle) for line in child._flow.lines if line.page == 0)
    assert publication.page_for_position(publication._sections[1][0]) == 2
    assert child.page_for_position(0) == 1


def test_decorations_and_wrapping_render_to_pdf_without_changing_source(qapp, tmp_path):
    document = content(qapp, "Original body text. " * 20)
    state = DocumentState()
    state.design = DesignSettings(page_color="#ffffee", page_border="box",
                                   watermark_kind="text", watermark_text="DRAFT").to_dict()
    state.features["objects"] = [{"id": "shape", "kind": "text", "page": 0,
                                   "x": 270, "y": 96, "width": 190, "height": 120,
                                   "wrap": "square", "text": "Floating object"}]
    document.folio_state = state
    before = document.toHtml()
    state_before = copy.deepcopy(state.to_dict())
    path = tmp_path / "decorated.pdf"
    count = write_pdf(path, document, state.page, "Decorated")
    reader = QPdfDocument()
    reader.load(str(path))
    text = "".join(reader.getAllText(page).text() for page in range(count))
    assert "DRAFT" in text
    assert "object" in text
    assert text.count("Original") == 20
    assert document.toHtml() == before
    assert state.to_dict() == state_before


def test_hyphenation_is_dictionary_driven_and_preserves_source_positions(qapp):
    document = content(qapp, "\U0001f600 implementation documentation configuration. " * 300)
    before = document.toHtml()
    settings = PageSettings(columns=3, hyphenation="automatic", consecutive_hyphen_limit=1,
                            balance_columns=False)
    publication = Publication(document, settings, "Hyphenation")
    assert publication._inserted_hyphens
    assert "\u00ad" in publication.document.toPlainText()
    assert document.toHtml() == before
    assert publication.page_for_position(document.characterCount() - 2) == publication.page_count()
    run = 0
    block = publication.document.begin()
    while block.isValid():
        layout = block.layout()
        for index in range(layout.lineCount()):
            line = layout.lineAt(index)
            last = block.position() + line.textStart() + line.textLength() - 1
            run = run + 1 if publication.document.characterAt(last) == "\u00ad" else 0
            assert run <= 1
        block = block.next()


def test_page_paint_failure_restores_painter_transform(qapp, monkeypatch):
    document = content(qapp, "Test")
    publication = Publication(document, PageSettings(), "Test")
    image = QImage(100, 100, QImage.Format.Format_ARGB32)
    painter = QPainter(image)
    painter.translate(10, 20)
    original = painter.transform()
    monkeypatch.setattr(publication, "_paint_body", lambda *args: (_ for _ in ()).throw(RuntimeError("fail")))
    with pytest.raises(RuntimeError):
        publication.paint_page(painter, 0, 2)
    assert painter.transform() == original
    painter.end()
