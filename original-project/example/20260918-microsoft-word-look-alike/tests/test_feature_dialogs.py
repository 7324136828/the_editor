import copy

import pytest
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QImage
from PySide6.QtWidgets import QDialog, QDialogButtonBox

from folio.design import DesignSettings
from folio.feature_dialogs import (
    DesignDialog, FloatingInspector, SignatureInputDialog, TableDimensionsDialog,
)


def _objects():
    return [
        {"id": "a", "kind": "text", "text": "Alpha", "name": "Alpha", "x": 10, "y": 20,
         "width": 100, "height": 50, "page": 0, "z": 0, "wrap": "front", "visible": True},
        {"id": "b", "kind": "text", "text": "Beta", "name": "Beta", "x": 160, "y": 70,
         "width": 100, "height": 50, "page": 0, "z": 1, "wrap": "square", "visible": True},
    ]


def test_design_dialog_edits_copy_and_validates_missing_watermark(qapp):
    original = DesignSettings()
    dialog = DesignDialog(original)
    dialog.theme_combo.setCurrentIndex(dialog.theme_combo.findData("Woodland"))
    dialog.style_set_combo.setCurrentIndex(dialog.style_set_combo.findData("Compact"))
    dialog.spacing_combo.setCurrentIndex(dialog.spacing_combo.findData("Relaxed"))
    dialog.page_color_button.set_color("#fef5e0")
    dialog.border_combo.setCurrentIndex(dialog.border_combo.findData("double"))
    settings = dialog.settings()
    assert settings.theme == "Woodland" and settings.style_set == "Compact"
    assert settings.paragraph_spacing == "Relaxed" and settings.page_color == "#fef5e0"
    assert original == DesignSettings()
    assert not dialog.preview.pixmap().isNull()
    dialog.watermark_kind_combo.setCurrentIndex(dialog.watermark_kind_combo.findData("image"))
    assert not dialog.buttons.button(QDialogButtonBox.StandardButton.Ok).isEnabled()
    dialog.accept()
    assert dialog.result() != QDialog.DialogCode.Accepted
    assert "image" in dialog.error_label.text().lower()


def test_design_watermark_embeds_image_and_retains_source_settings(tmp_path, qapp):
    image = QImage(10, 10, QImage.Format.Format_ARGB32)
    image.fill(QColor("red"))
    path = tmp_path / "mark.png"
    assert image.save(str(path))
    dialog = DesignDialog(DesignSettings())
    dialog.set_watermark_image(path)
    settings = dialog.settings()
    assert settings.watermark_kind == "image"
    assert settings.watermark_image.startswith("data:image/png;base64,")
    assert settings == DesignDialog(settings).settings()
    path.unlink()
    assert not DesignDialog(settings).preview.pixmap().isNull()


def test_floating_inspector_group_geometry_visibility_and_ungroup(qapp):
    objects = _objects()
    calls = []
    inspector = FloatingInspector(objects, lambda: calls.append(True))
    inspector.select_ids({"a", "b"})
    inspector.perform("group")
    assert len(objects) == 3
    group = objects[-1]
    assert inspector.selected_ids() == {group["id"]}
    assert inspector.tree.topLevelItemCount() == 1
    assert inspector.tree.topLevelItem(0).childCount() == 2
    inspector.x_spin.setValue(30)
    inspector.y_spin.setValue(40)
    inspector.width_spin.setValue(500)
    inspector.height_spin.setValue(200)
    inspector.page_spin.setValue(2)
    inspector.apply_properties()
    assert objects[0]["x"] == 30 and objects[0]["y"] == 40
    assert objects[1]["x"] == 330 and objects[1]["y"] == 140
    assert objects[0]["width"] == 200 and objects[0]["height"] == 100
    assert all(obj["page"] == 1 for obj in objects)
    inspector.tree.topLevelItem(0).setCheckState(0, Qt.CheckState.Unchecked)
    qapp.processEvents()
    assert not group["visible"]
    inspector.perform("ungroup")
    assert len(objects) == 2
    assert all("group" not in obj for obj in objects)
    assert len(calls) == 4


def test_floating_inspector_edits_text_orders_layers_and_deletes_groups(qapp):
    objects = _objects()
    inspector = FloatingInspector(objects, lambda: None)
    inspector.select_ids({"a"})
    inspector.text_edit.setPlainText("Changed text")
    inspector.wrap_combo.setCurrentIndex(inspector.wrap_combo.findData("tight"))
    inspector.apply_properties()
    assert objects[0]["text"] == "Changed text" and objects[0]["wrap"] == "tight"
    inspector.perform("forward")
    assert objects[0]["z"] > objects[1]["z"]
    inspector.perform("backward")
    assert objects[0]["z"] < objects[1]["z"]
    inspector.select_ids({"a", "b"})
    inspector.perform("group")
    inspector.perform("delete")
    assert objects == []
    assert inspector.tree.topLevelItemCount() == 0


def test_floating_inspector_invalid_alignment_is_atomic(qapp):
    objects = _objects()
    objects[1]["page"] = 1
    before = copy.deepcopy(objects)
    calls = []
    inspector = FloatingInspector(objects, lambda: calls.append(True))
    inspector.select_ids({"a", "b"})
    inspector.perform("left")
    assert objects == before and not calls
    assert "same page" in inspector.error_label.text()


def test_signature_requires_name_and_deliberate_consent(qapp):
    dialog = SignatureInputDialog(name="Ada Lovelace")
    assert not dialog.buttons.button(QDialogButtonBox.StandardButton.Ok).isEnabled()
    dialog.accept()
    assert dialog.result() != QDialog.DialogCode.Accepted
    dialog.consent_check.setChecked(True)
    assert dialog.buttons.button(QDialogButtonBox.StandardButton.Ok).isEnabled()
    dialog.name_edit.setText("   ")
    dialog.accept()
    assert dialog.result() != QDialog.DialogCode.Accepted
    dialog.name_edit.setText("  Ada Lovelace  ")
    dialog.accept()
    assert dialog.result() == QDialog.DialogCode.Accepted
    assert dialog.signature_name() == "Ada Lovelace" and dialog.consent()


def test_table_dimensions_support_uniform_and_per_column_values(qapp):
    dialog = TableDimensionsDialog(3, 2)
    assert dialog.dimensions() == (None, None)
    dialog.widths_edit.setText("120, 180, 120")
    dialog.heights_edit.setText("24")
    assert dialog.dimensions() == ([120, 180, 120], [24, 24])
    for invalid in ("20, 30", "nan", "inf", "0", "abc", "12,"):
        dialog.widths_edit.setText(invalid)
        with pytest.raises(ValueError):
            dialog.dimensions()
        dialog.accept()
        assert dialog.result() != QDialog.DialogCode.Accepted
        assert dialog.error_label.text()
