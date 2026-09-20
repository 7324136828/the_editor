import pytest
from PySide6.QtCore import QSizeF
from PySide6.QtGui import QColor, QImage, QTextCursor, QTextFormat, QTextImageFormat

from folio.editor import SafeDocument, _image_to_data_uri, mm_to_px
from folio.models import PageSettings
from folio.publishing import Publication, write_pdf


@pytest.fixture
def document(qapp):
    return SafeDocument()


def _long_doc(document, paragraphs=60):
    html = "".join(f"<p>Paragraph {i} with some body text to fill the page "
                   "and make it longer so pagination has real content.</p>"
                   for i in range(paragraphs))
    document.setHtml(html)


def test_pdf_written(document, tmp_path):
    document.setPlainText("hello pdf")
    out = tmp_path / "out.pdf"
    pages = write_pdf(out, document, PageSettings(), "T")
    assert pages >= 1
    data = out.read_bytes()
    assert data.startswith(b"%PDF-")
    assert len(data) > 500


def test_pdf_multiple_pages(document, tmp_path):
    _long_doc(document, 80)
    out = tmp_path / "long.pdf"
    pages = write_pdf(out, document, PageSettings(), "T")
    assert pages >= 2
    assert out.read_bytes().startswith(b"%PDF-")


def test_portrait_and_landscape_dimensions(document, tmp_path):
    from PySide6.QtPdf import QPdfDocument
    document.setPlainText("dims")
    p = tmp_path / "p.pdf"
    write_pdf(p, document, PageSettings(landscape=False), "T")
    land = tmp_path / "l.pdf"
    write_pdf(land, document, PageSettings(landscape=True), "T")
    reader = QPdfDocument()
    reader.load(str(p))
    assert reader.pageCount() >= 1
    size = reader.pagePointSize(0)
    assert size.width() == pytest.approx(595, abs=2)
    assert size.height() == pytest.approx(842, abs=2)
    reader2 = QPdfDocument()
    reader2.load(str(land))
    size2 = reader2.pagePointSize(0)
    assert size2.width() == pytest.approx(842, abs=2)
    assert size2.height() == pytest.approx(595, abs=2)


def test_two_column_narrower_body(document):
    _long_doc(document, 40)
    one = Publication(document, PageSettings(columns=1), "t")
    two = Publication(document, PageSettings(columns=2), "t")
    assert two.column_w < one.column_w


def test_page_break_advances_physical_page(document):
    document.setHtml("<p>first</p><p>second</p>")
    block = document.begin().next()
    fmt = block.blockFormat()
    fmt.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
    cursor = QTextCursor(block)
    cursor.setBlockFormat(fmt)
    publication = Publication(document, PageSettings(), "t")
    assert publication.page_count() >= 2
    two_col = Publication(document, PageSettings(columns=2), "t")
    assert two_col.page_count() >= 2


def test_pdf_text_and_page_count(document, tmp_path):
    from PySide6.QtPdf import QPdfDocument
    _long_doc(document, 80)
    settings = PageSettings(header="Hdr {title}",
                            footer="Page {page} of {pages}")
    out = tmp_path / "text.pdf"
    pages = write_pdf(out, document, settings, "My Title")
    reader = QPdfDocument()
    reader.load(str(out))
    assert reader.pageCount() == pages
    text = reader.getAllText(0).text()
    assert "Paragraph" in text
    assert "Hdr My Title" in text
    assert "Page 1 of" in text


def test_headers_footers_rendered(document):
    document.setPlainText("body")
    settings = PageSettings(header="Hdr {title}",
                            footer="Page {page} of {pages}")
    publication = Publication(document, settings, "My Title")
    assert publication.running_text("header", 0) == "Hdr My Title"
    assert publication.running_text("footer", 0) == "Page 1 of 1"


def test_export_does_not_mutate_document(document):
    html = "<p>keep me <b>bold</b></p>"
    document.setHtml(html)
    document.setModified(False)
    before = document.toHtml()
    Publication(document, PageSettings(), "t")
    assert document.toHtml() == before
    assert not document.isModified()


def test_unicode_text(document, tmp_path):
    document.setPlainText("Héllo wörld — 你好")
    out = tmp_path / "u.pdf"
    write_pdf(out, document, PageSettings(), "T")
    assert out.read_bytes().startswith(b"%PDF-")


def test_oversized_image_bounded(document):
    image = QImage(2000, 1000, QImage.Format.Format_RGB32)
    image.fill(QColor("#123456"))
    fmt = QTextImageFormat()
    fmt.setName(_image_to_data_uri(image))
    fmt.setWidth(2000)
    fmt.setHeight(1000)
    cursor = QTextCursor(document)
    document.addResource(2, "", image)
    cursor.insertImage(fmt)
    publication = Publication(document, PageSettings(), "t")
    block = publication.document.begin()
    iterator = block.begin()
    found = None
    while not iterator.atEnd():
        frag = iterator.fragment()
        if frag.isValid() and frag.charFormat().isImageFormat():
            found = QTextImageFormat(frag.charFormat().toImageFormat())
        iterator += 1
    assert found is not None
    assert found.width() <= publication.column_w + 1
    assert found.height() <= publication.body_h + 1
    assert found.width() / found.height() == pytest.approx(2.0, abs=0.05)


def test_page_for_position(document):
    _long_doc(document, 80)
    publication = Publication(document, PageSettings(), "t")
    first = publication.page_for_position(0)
    assert first == 1
    last_pos = document.characterCount() - 2
    last = publication.page_for_position(last_pos)
    assert last >= first


def test_page_range_paint(document, tmp_path, monkeypatch):
    _long_doc(document, 80)
    publication = Publication(document, PageSettings(), "t")
    if publication.page_count() < 3:
        pytest.skip("document did not paginate far enough")
    from PySide6.QtGui import QPageRanges, QPdfWriter
    painted = []
    original = publication.paint_page
    def recorder(painter, index, scale=1.0):
        painted.append(index)
        return original(painter, index, scale)
    publication.paint_page = recorder
    writer = QPdfWriter(str(tmp_path / "range.pdf"))
    ranges = QPageRanges()
    ranges.addRange(2, 3)
    writer.setPageRanges(ranges)
    count = publication.paint(writer)
    assert count == 2
    assert painted == [1, 2]
    del writer


def test_disjoint_page_range_paint(document, tmp_path):
    _long_doc(document, 80)
    publication = Publication(document, PageSettings(), "t")
    if publication.page_count() < 3:
        pytest.skip("document did not paginate far enough")
    from PySide6.QtGui import QPageRanges, QPdfWriter
    painted = []
    original = publication.paint_page
    def recorder(painter, index, scale=1.0):
        painted.append(index)
        return original(painter, index, scale)
    publication.paint_page = recorder
    writer = QPdfWriter(str(tmp_path / "disjoint.pdf"))
    ranges = QPageRanges()
    ranges.addPage(1)
    ranges.addPage(3)
    writer.setPageRanges(ranges)
    count = publication.paint(writer)
    assert count == 2
    assert painted == [0, 2]
    del writer


def test_empty_page_range_errors(document, tmp_path):
    _long_doc(document, 20)
    publication = Publication(document, PageSettings(), "t")
    from PySide6.QtGui import QPageRanges, QPdfWriter
    writer = QPdfWriter(str(tmp_path / "none.pdf"))
    ranges = QPageRanges()
    ranges.addPage(publication.page_count() + 5)
    writer.setPageRanges(ranges)
    with pytest.raises(RuntimeError):
        publication.paint(writer)
    del writer


def test_inactive_painter_errors(document, tmp_path):
    document.setPlainText("x")
    publication = Publication(document, PageSettings(), "t")
    from PySide6.QtGui import QImage
    import folio.publishing as pub

    class DeadDevice(QImage):
        def __init__(self):
            super().__init__(16, 16, QImage.Format.Format_ARGB32)

    real_painter = pub.QPainter

    class FailingPainter(real_painter):
        def __init__(self, device):
            super().__init__()

        def isActive(self):
            return False

    pub.QPainter = FailingPainter
    try:
        with pytest.raises(RuntimeError):
            publication.paint(DeadDevice())
    finally:
        pub.QPainter = real_painter


def test_new_page_failure_preserves_existing_pdf(document, tmp_path):
    _long_doc(document, 80)
    publication = Publication(document, PageSettings(), "t")
    if publication.page_count() < 2:
        pytest.skip("document did not paginate far enough")
    from PySide6.QtGui import QPdfWriter
    existing = tmp_path / "existing.pdf"
    existing.write_bytes(b"ORIGINAL PDF")
    writer = QPdfWriter(str(tmp_path / "out.pdf"))

    def failing_new_page():
        return False

    writer.newPage = failing_new_page
    with pytest.raises(RuntimeError):
        publication.paint(writer)
    del writer
    assert existing.read_bytes() == b"ORIGINAL PDF"


def test_publication_proof_render(document, tmp_path, qapp):
    import os
    if not os.environ.get("FOLIO_SCREENSHOT"):
        pytest.skip("FOLIO_SCREENSHOT not set")
    from PySide6.QtGui import QImage, QPainter
    _long_doc(document, 30)
    settings = PageSettings(header="Hdr {title}",
                            footer="Page {page} of {pages}")
    publication = Publication(document, settings, "My Title")
    image = QImage(int(publication.page_w), int(publication.page_h),
                   QImage.Format.Format_ARGB32)
    image.fill(QColor("white"))
    painter = QPainter(image)
    publication.paint_page(painter, 0)
    painter.end()
    from pathlib import Path
    artifacts = Path(__file__).resolve().parent.parent / "artifacts"
    artifacts.mkdir(exist_ok=True)
    out = artifacts / "folio-publication.png"
    image.save(str(out))
    assert out.exists() and out.stat().st_size > 1000


def test_painter_ends_when_page_fails(document, tmp_path, monkeypatch):
    document.setPlainText("x")
    publication = Publication(document, PageSettings(), "t")
    import folio.publishing as pub
    from PySide6.QtGui import QPainter as RealPainter, QPdfWriter
    ended = []
    class RecordingPainter(RealPainter):
        def end(self):
            ended.append(True)
            super().end()
    monkeypatch.setattr(pub, "QPainter", RecordingPainter)
    def boom(*args, **kwargs):
        raise RuntimeError("simulated paint failure")
    publication.paint_page = boom
    writer = QPdfWriter(str(tmp_path / "fail.pdf"))
    with pytest.raises(RuntimeError):
        publication.paint(writer)
    assert ended
    del writer
