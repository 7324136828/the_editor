from email import policy
from email.parser import BytesParser

import pytest
from PySide6.QtCore import Qt

from folio.docx_io import read_docx
from folio.editor import SafeDocument
from folio.mailmerge import (Condition, ENVELOPE_TEMPLATES, LABEL_TEMPLATES,
    MergeError, MergeSkipped, RecipientSet, compound_document, envelope_settings,
    label_sheets, merge_batch, merge_document, rule_token)
from folio.models import DocumentState
from folio.publishing import Publication


def document(text):
    doc = SafeDocument()
    doc.setPlainText(text)
    return doc


def test_recipient_sort_filter_and_exclusion_are_stable():
    rows = [{"Name": "Zoe", "Age": "10"}, {"Name": "Ada", "Age": "2"},
            {"Name": "Bea", "Age": "30"}, {"Name": "Kai", "Age": "20"}]
    recipients = RecipientSet(rows, excluded={3}, filters=[Condition("Age", "greater than", "5")], sort_field="Age")
    assert [index for index, row in recipients.active()] == [0, 2]
    recipients.descending = True
    assert [index for index, row in recipients.active()] == [2, 0]
    assert rows[0]["Name"] == "Zoe"
    with pytest.raises(MergeError, match="finite numbers"):
        Condition("Age", "greater than", "NaN").evaluate(rows[0])


def test_rule_mapping_address_greeting_and_utf16(qapp):
    rule = rule_token("IF", Condition("Score", "greater than", "10"), "Premium", "Standard")
    doc = document("😀 {{GreetingLine}}\n{{AddressBlock}}\n" + rule)
    result = merge_document(doc, {"Person": "Ada", "Road": "1 Main St", "Score": "12", "City": "Paris"},
                            {"FullName": "Person", "Address1": "Road"})
    assert result.toPlainText() == "😀 Dear Ada,\nAda\n1 Main St\nParis\nPremium"
    assert "{{GreetingLine}}" in doc.toPlainText()


def test_rule_never_evaluates_python_or_recurses(qapp):
    doc = document(rule_token("IF", Condition("Name", "equals", "Ada"), "__import__('os')", "Nothing"))
    assert merge_document(doc, {"Name": "Ada"}).toPlainText() == "__import__('os')"
    with pytest.raises(MergeError):
        rule_token("IF", Condition("Name"), "{{Secret}}")
    with pytest.raises(MergeError):
        merge_document(document("{{IF|Name|__import__|x|a|b}}"), {"Name": "x"})


def test_skip_if_and_preflight_leave_directory_untouched(qapp, tmp_path):
    doc = document("{{Name}}" + rule_token("SKIPIF", Condition("Enabled", "equals", "no")))
    with pytest.raises(MergeSkipped):
        merge_document(doc, {"Name": "Ada", "Enabled": "no"})
    destination = tmp_path / "batch"
    outputs = merge_batch(doc, DocumentState(title="Test"),
        [{"Name": "Ada", "Enabled": "no"}, {"Name": "Bea", "Enabled": "yes"}], destination)
    assert len(outputs) == 1
    loaded, _ = read_docx(outputs[0])
    assert "Bea" in loaded.html and "Ada" not in loaded.html
    invalid = tmp_path / "invalid"
    with pytest.raises(MergeError, match="Record 2"):
        merge_batch(document("{{Name}}"), DocumentState(), [{"Name": "Ada"}, {}], invalid)
    assert not invalid.exists()


def test_compound_document_page_boundaries(qapp):
    combined = compound_document([document("Ada"), document("Bea"), document("Zoe")])
    assert combined.toPlainText() == "Ada\nBea\nZoe"
    assert Publication(combined, DocumentState().page, "Batch").page_count() == 3


def test_compound_table_documents_start_new_pages(qapp):
    first = SafeDocument()
    first.setHtml("<table><tr><td>Ada</td></tr></table>")
    second = SafeDocument()
    second.setHtml("<table><tr><td>Bea</td></tr></table>")
    combined = compound_document([first, second])
    assert Publication(combined, DocumentState().page, "Batch").page_count() == 2


def test_batch_compound_roundtrip_pdf_and_unique_files(qapp, tmp_path):
    doc = document("Dear {{Name}},")
    state = DocumentState(title="Letters", track_changes=True)
    records = [{"Name": "Ada"}, {"Name": "Bea"}]
    paths = merge_batch(doc, state, records, tmp_path, "compound_docx")
    loaded, _ = read_docx(paths[0])
    restored = SafeDocument()
    restored.setHtml(loaded.html)
    assert "Ada" in restored.toPlainText() and "Bea" in restored.toPlainText()
    assert Publication(restored, loaded.page, loaded.title).page_count() == 2
    assert not loaded.track_changes
    assert state.track_changes
    second = merge_batch(doc, state, records, tmp_path, "compound_docx")
    assert second[0] != paths[0]
    pdfs = merge_batch(doc, state, records, tmp_path, "compound_pdf")
    assert pdfs[0].read_bytes().startswith(b"%PDF")


def test_eml_is_unsent_multipart_with_validated_headers(qapp, tmp_path):
    doc = document("Hello {{Name}}")
    records = [{"Name": "Ada", "Email": "ada@example.test"}]
    output = merge_batch(doc, DocumentState(title="Greetings"), records, tmp_path, "eml")[0]
    message = BytesParser(policy=policy.default).parsebytes(output.read_bytes())
    assert message["To"] == "ada@example.test"
    assert message["X-Unsent"] == "1"
    assert "Ada" in message.get_body(preferencelist=("plain",)).get_content()
    assert message.get_body(preferencelist=("html",)) is not None
    with pytest.raises(MergeError, match="invalid email"):
        merge_batch(doc, DocumentState(), [{"Name": "Ada", "Email": "a@b.test\nBcc: secret@x.test"}], tmp_path, "eml")
    with pytest.raises(MergeError, match="subject"):
        merge_batch(doc, DocumentState(), records, tmp_path, "eml", subject="Hello\nBcc: secret@x.test")


def test_standard_envelope_dimensions():
    expected = [(241.3, 104.775), (220.0, 110.0), (229.0, 162.0)]
    for name, size in zip(ENVELOPE_TEMPLATES, expected):
        settings = envelope_settings(name)
        assert settings.size_mm() == size
        assert settings.footer == ""
        assert settings.content_width_mm() > 40


@pytest.mark.parametrize("name", list(LABEL_TEMPLATES))
def test_label_rectangles_and_sheet_page_count(qapp, name):
    template = LABEL_TEMPLATES[name]
    rectangles = template.rectangles()
    assert len(rectangles) == template.capacity
    for index, rect in enumerate(rectangles):
        for other in rectangles[index + 1:]:
            overlap = rect.intersected(other)
            assert overlap.width() < 1e-6 or overlap.height() < 1e-6
    docs = [document("Ada\n1 Main St\nParis")] * (template.capacity + 1)
    sheet = label_sheets(docs, template)
    assert Publication(sheet, template.settings(), "Labels").page_count() == 2


def test_label_overflow_fails_before_writing(qapp, tmp_path):
    name = next(iter(LABEL_TEMPLATES))
    with pytest.raises(MergeError, match="does not fit"):
        merge_batch(document("{{Address}}"), DocumentState(), [{"Address": "Line\n" * 60}], tmp_path,
                    labels=name)
    assert not list(tmp_path.glob("*.docx"))


def test_batch_cancel_retains_completed_files(qapp, tmp_path):
    progress = []
    files = merge_batch(document("{{Name}}"), DocumentState(), [{"Name": str(i)} for i in range(5)],
        tmp_path, progress=lambda done, total: progress.append(done), cancelled=lambda: len(progress) == 2)
    assert len(files) == 2
    assert all(path.exists() for path in files)


@pytest.mark.parametrize("feature", ["nodes", "objects", "sections"])
@pytest.mark.parametrize("output", ["compound_docx", "compound_pdf", "pdf", "eml"])
def test_complex_templates_cannot_lose_features(qapp, tmp_path, feature, output):
    state = DocumentState()
    state.features[feature] = {"note": {"kind": "note"}} if feature == "nodes" else [{"id": "object"}]
    destination = tmp_path / "output"
    with pytest.raises(MergeError, match="individual DOCX"):
        merge_batch(document("Hi {{Name}}"), state, [{"Name": "Ada"}], destination, output)
    assert not destination.exists()


def test_workbench_exclusions_survive_sort_and_preview(qapp):
    from folio.mailing_dialogs import MailingWorkbenchDialog

    recipients = RecipientSet([{"Name": "Zoe"}, {"Name": "Ada"}])
    dialog = MailingWorkbenchDialog(document("Hi {{Name}}"), DocumentState(), recipients)
    assert dialog.preview.toPlainText() == "Hi Zoe"
    dialog.records_table.item(0, 0).setCheckState(Qt.CheckState.Unchecked)
    dialog.sort_column.setCurrentIndex(dialog.sort_column.findData("Name"))
    assert dialog.records_table.item(0, 1).text() == "Ada"
    assert dialog.records_table.item(1, 0).checkState() == Qt.CheckState.Unchecked
    assert dialog.preview.toPlainText() == "Hi Ada"
    assert dialog.document.toPlainText() == "Hi {{Name}}"
    dialog.close()
