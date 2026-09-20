from folio.editor import SafeDocument
from folio.hyphenation import apply_dictionary, layout_position
from folio.models import PageSettings
from folio.publishing import Publication


def test_dictionary_uses_known_breaks_not_vowel_guessing(qapp):
    document = SafeDocument()
    document.setPlainText("document unknownwordxyz AUTOMATIC")
    inserted = apply_dictionary(document)
    assert document.toPlainText() == "doc\u00adu\u00adment unknownwordxyz AUTOMATIC"
    assert inserted == [3, 4]


def test_dictionary_caps_and_utf16_positions(qapp):
    document = SafeDocument()
    document.setPlainText("🙂 DOCUMENT")
    inserted = apply_dictionary(document, True)
    assert inserted == [6, 7]
    assert document.toPlainText() == "🙂 DOC\u00adU\u00adMENT"
    assert layout_position(11, inserted) == 13


def test_custom_dictionary_boundaries(qapp):
    document = SafeDocument()
    document.setPlainText("customword example")
    apply_dictionary(document, entries=["cus-tom-word"])
    assert document.toPlainText() == "cus\u00adtom\u00adword example"


def test_publishing_hyphenates_without_mutating_editor(qapp):
    document = SafeDocument()
    document.setPlainText("document application information")
    before = document.toHtml()
    publication = Publication(document, PageSettings(hyphenation="automatic"), "Dictionary")
    assert "\u00ad" in publication.document.toPlainText()
    assert document.toHtml() == before
