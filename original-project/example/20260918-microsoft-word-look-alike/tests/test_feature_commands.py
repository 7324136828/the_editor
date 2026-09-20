import copy
import os
from pathlib import Path

import pytest
from PySide6.QtCore import Qt
from PySide6.QtGui import QTextCursor
from PySide6.QtTest import QTest
from PySide6.QtWidgets import QDialog, QInputDialog, QToolButton

from folio import dialogs, document_features, feature_dialogs, structured_media
from folio.design import DesignSettings, resolved_style
from folio.docx_io import read_docx
from folio.editor import pt_to_px
from folio.models import DocumentState
from folio.window import FolioWindow


@pytest.fixture
def window(qapp, store):
    win = FolioWindow(store=store, restore=False)
    win.load_state(DocumentState(html="<p>Draft document.</p>"), clean=True)
    yield win
    win._dirty = False
    win.close()
    qapp.processEvents()


def _button(window, label):
    return next(button for button in window.findChildren(QToolButton) if button.text() == label)


def _cursor_at_end(window):
    cursor = window.editor.textCursor()
    cursor.movePosition(QTextCursor.MoveOperation.End)
    window.editor.setTextCursor(cursor)


def test_equation_ribbon_insert_edit_and_docx_roundtrip(window, monkeypatch, tmp_path):
    sources = [r"\frac{x^2}{\sqrt{y}}", r"\sum_{n=1}^{5}n^2"]
    seen = []

    def accept(dialog):
        seen.append(dialog.source())
        dialog.source_edit.setPlainText(sources.pop(0))
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(structured_media.EquationDialog, "exec", accept)
    _cursor_at_end(window)
    _button(window, "Equation").click()
    key, node = next(iter(window.state.features["nodes"].items()))
    original_key = key
    assert node["source"] == r"\frac{x^2}{\sqrt{y}}"
    assert key in document_features.node_ranges(window.editor.document())
    window._edit_content_node()
    nodes = window.state.features["nodes"]
    active = document_features.node_ranges(window.editor.document())
    assert len(active) == 1
    key = next(iter(active))
    node = nodes[key]
    assert seen[1] == r"\frac{x^2}{\sqrt{y}}"
    assert node["source"] == r"\sum_{n=1}^{5}n^2"
    window.editor.undo()
    assert set(document_features.node_ranges(window.editor.document())) == {original_key}
    assert nodes[original_key]["source"] == r"\frac{x^2}{\sqrt{y}}"
    assert window._selected_node("equation")[1]["source"] == r"\frac{x^2}{\sqrt{y}}"
    window.editor.redo()
    assert set(document_features.node_ranges(window.editor.document())) == {key}
    assert window._selected_node("equation")[1]["source"] == r"\sum_{n=1}^{5}n^2"
    path = tmp_path / "equation.docx"
    assert window.save_to(path)
    loaded, warnings = read_docx(path)
    assert not warnings
    window.load_state(loaded, path)
    assert window.state.features["nodes"][key]["source"] == node["source"]
    assert key in document_features.node_ranges(window.editor.document())
    assert "data:image/png;base64," in window.editor.document().toHtml()


def test_structured_diagram_edit_preserves_only_active_source(window, monkeypatch, tmp_path):
    steps = ["Plan\nBuild\nShip", "Plan\nTest\nShip"]

    def accept(dialog):
        dialog.steps.setPlainText(steps.pop(0))
        dialog._accept()
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(dialogs.DiagramDialog, "exec", accept)
    _cursor_at_end(window)
    _button(window, "SmartArt").click()
    original_key = next(iter(document_features.node_ranges(window.editor.document())))
    assert window.state.features["nodes"][original_key]["steps"] == ["Plan", "Build", "Ship"]
    window._edit_content_node()
    active = document_features.node_ranges(window.editor.document())
    assert len(active) == 1
    key = next(iter(active))
    assert window.state.features["nodes"][key]["steps"] == ["Plan", "Test", "Ship"]
    window.editor.undo()
    assert set(document_features.node_ranges(window.editor.document())) == {original_key}
    assert window._selected_node("diagram")[1]["steps"] == ["Plan", "Build", "Ship"]
    window.editor.redo()
    assert set(document_features.node_ranges(window.editor.document())) == {key}
    assert window._selected_node("diagram")[1]["steps"] == ["Plan", "Test", "Ship"]
    path = tmp_path / "process.docx"
    assert window.save_to(path)
    loaded, warnings = read_docx(path)
    assert not warnings
    assert loaded.features["nodes"][key]["steps"] == ["Plan", "Test", "Ship"]
    window.load_state(loaded, path)
    assert set(document_features.node_ranges(window.editor.document())) == {key}


def test_signature_ui_records_consent_and_roundtrips(window, monkeypatch, tmp_path):
    monkeypatch.setattr(QInputDialog, "getText", lambda *args, **kwargs: ("Ada Lovelace", True))

    def sign(dialog):
        dialog.name_edit.setText("Ada Lovelace")
        dialog.consent_check.setChecked(True)
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(feature_dialogs.SignatureInputDialog, "exec", sign)
    _cursor_at_end(window)
    window._insert_signature_fields()
    window._edit_content_node()
    key, node = next(iter(window.state.features["nodes"].items()))
    assert node["signed_by"] == "Ada Lovelace" and node["consent"] is True
    assert node["method"] == "typed" and len(node["body_sha256"]) == 64
    assert "Signed by Ada Lovelace" in window.editor.toPlainText()
    path = tmp_path / "signed.docx"
    assert window.save_to(path)
    loaded, warnings = read_docx(path)
    assert not warnings and loaded.features["nodes"][key] == node


def test_theme_dialog_drives_future_semantic_styles(window, monkeypatch, tmp_path):
    settings = DesignSettings(theme="Woodland", style_set="Large Print", paragraph_spacing="Compact",
                              page_color="#fffde0", page_border="box", watermark_kind="text", watermark_text="DRAFT")
    monkeypatch.setattr(feature_dialogs.DesignDialog, "exec", lambda self: QDialog.DialogCode.Accepted)
    monkeypatch.setattr(feature_dialogs.DesignDialog, "settings", lambda self: settings)
    _button(window, "Document Theme").click()
    window.editor.apply_style("Heading 1")
    target = resolved_style("Heading 1", settings)
    cursor = QTextCursor(window.editor.document().begin())
    cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
    assert cursor.charFormat().fontPointSize() == target["size"]
    assert cursor.charFormat().foreground().color().name() == target["color"]
    assert window.editor.design_settings == settings
    path = tmp_path / "design.docx"
    assert window.save_to(path)
    loaded, warnings = read_docx(path)
    assert not warnings and loaded.design == settings.to_dict()
    window.load_state(loaded, path)
    assert window.editor.design_settings == settings
    window.editor.apply_style("Normal")
    cursor = QTextCursor(window.editor.document().begin())
    cursor.select(QTextCursor.SelectionType.BlockUnderCursor)
    assert cursor.charFormat().fontPointSize() == resolved_style("Normal", settings)["size"]


def test_floating_inspector_tracks_document_replacement(window, monkeypatch):
    monkeypatch.setattr(QInputDialog, "getMultiLineText", lambda *args, **kwargs: ("Caption", True))
    _button(window, "Text Box").click()
    assert len(window.state.features["objects"]) == 1
    inspector = window.floating_dock.widget()
    key = window.state.features["objects"][0]["id"]
    inspector.select_ids({key})
    inspector.x_spin.setValue(210)
    inspector.text_edit.setPlainText("Edited caption")
    inspector.apply_properties()
    assert window.state.features["objects"][0]["x"] == 210
    assert window.state.features["objects"][0]["text"] == "Edited caption"
    old = copy.deepcopy(window.state.features)
    window.load_state(DocumentState(html="<p>New document</p>"), clean=True)
    assert inspector.objects is window.state.features["objects"]
    assert inspector.tree.topLevelItemCount() == 0
    assert old["objects"][0]["text"] == "Edited caption"


def test_table_dimensions_dialog_updates_columns_and_minimum_height(window, monkeypatch):
    table = window.editor.insert_table(2, 3)
    window.editor.setTextCursor(table.cellAt(0, 0).firstCursorPosition())

    def dimensions(dialog):
        dialog.widths_edit.setText("80, 120, 100")
        dialog.heights_edit.setText("28, 36")
        return QDialog.DialogCode.Accepted

    monkeypatch.setattr(feature_dialogs.TableDimensionsDialog, "exec", dimensions)
    window._table_dimensions()
    table = window.editor.textCursor().currentTable()
    assert [length.rawValue() for length in table.format().columnWidthConstraints()] == [pt_to_px(80), pt_to_px(120), pt_to_px(100)]
    assert table.cellAt(0, 0).firstCursorPosition().blockFormat().lineHeight() == pt_to_px(28)
    assert table.cellAt(1, 0).firstCursorPosition().blockFormat().lineHeight() == pt_to_px(36)
    assert table.cellAt(0, 0).firstCursorPosition().blockFormat().lineHeightType() == 3


def test_viewing_prevents_feature_mutation_and_reviewing_tracks(window, monkeypatch):
    def unexpected(*args, **kwargs):
        pytest.fail("A write command opened a dialog in Viewing mode")

    monkeypatch.setattr(structured_media.EquationDialog, "exec", unexpected)
    monkeypatch.setattr(feature_dialogs.DesignDialog, "exec", unexpected)
    monkeypatch.setattr(QInputDialog, "getText", unexpected)
    monkeypatch.setattr(QInputDialog, "getMultiLineText", unexpected)
    window.permission_combo.setCurrentText("Viewing")
    before = window.editor.document().toHtml()
    before_features = copy.deepcopy(window.state.features)
    assert window.editor.isReadOnly()
    window._equation_editor()
    window._semantic_design()
    window._float_insert()
    window._signature_node()
    window._new_comment()
    window._blank_page()
    assert window.editor.document().toHtml() == before
    assert window.state.features == before_features
    window.permission_combo.setCurrentText("Reviewing")
    assert not window.editor.isReadOnly() and window.state.track_changes
    window._toggle_tracking(False)
    assert window.state.track_changes
    window.editor.textCursor().insertText("Tracked ")
    assert window.state.revisions
    window.permission_combo.setCurrentText("Editing")
    assert not window.state.track_changes


def test_comment_and_transcript_hooks_dispatch_locally(window, monkeypatch):
    comments = []
    window.hooks.commentDispatched.connect(comments.append)
    monkeypatch.setattr(QInputDialog, "getMultiLineText", lambda *args, **kwargs: ("Check this", True))
    window._new_comment()
    assert len(comments) == 1 and comments[0]["text"] == "Check this"
    comments[0]["text"] = "Changed outside"
    assert window.state.comments[0].text == "Check this"
    _cursor_at_end(window)
    window.hooks.set_dictation_state("listening")
    window.hooks.submit_transcript(" Local transcript")
    assert window.editor.toPlainText().endswith(" Local transcript")
    window.permission_combo.setCurrentText("Viewing")
    assert window.hooks.dictation_state == "idle"
    with pytest.raises(ValueError):
        window.hooks.submit_transcript("Should not insert")


def test_viewing_keyboard_shortcuts_and_find_cannot_mutate(window, qapp):
    window.editor.set_list(False)
    window.permission_combo.setCurrentText("Viewing")
    window.show()
    window.editor.setFocus()
    qapp.processEvents()
    before = window.editor.document().toHtml()
    for key, modifiers in ((Qt.Key.Key_Tab, Qt.KeyboardModifier.NoModifier),
                           (Qt.Key.Key_Backtab, Qt.KeyboardModifier.ShiftModifier),
                           (Qt.Key.Key_Return, Qt.KeyboardModifier.ControlModifier),
                           (Qt.Key.Key_1, Qt.KeyboardModifier.ControlModifier)):
        QTest.keyClick(window.editor, key, modifiers)
        assert window.editor.document().toHtml() == before
    window._find()
    assert not window._find_dialog.replace_button.isEnabled()
    assert not window._find_dialog.replace_all_button.isEnabled()
    window._find_dialog.find_edit.setText("Draft")
    window._find_dialog.replace_edit.setText("Changed")
    window._find_dialog.replace_all_button.click()
    assert window.editor.document().toHtml() == before


def test_integrated_ribbon_and_object_pane_capture(window, monkeypatch, qapp):
    if os.environ.get("FOLIO_SCREENSHOT") != "1":
        pytest.skip("Screenshots require FOLIO_SCREENSHOT=1")
    monkeypatch.setattr(QInputDialog, "getMultiLineText", lambda *args, **kwargs: ("A floating caption", True))
    window.resize(1440, 960)
    window._float_insert()
    inspector = window.floating_dock.widget()
    inspector.select_ids({window.state.features["objects"][0]["id"]})
    window.ribbon.setCurrentIndex(2)
    window.show()
    qapp.processEvents()
    assert inspector.tree.topLevelItemCount() == 1
    target = Path(__file__).resolve().parents[1] / "artifacts" / "folio-engineering-insert.png"
    assert window.grab().save(str(target))
