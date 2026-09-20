from __future__ import annotations

import base64
import copy
import inspect
from functools import wraps
from pathlib import Path

from PySide6.QtCore import QSignalBlocker, Qt, QTimer
from PySide6.QtGui import QColor, QImage, QTextCursor, QTextLength
from PySide6.QtWidgets import (QComboBox, QDialog, QDockWidget, QFileDialog,
                              QInputDialog, QLabel, QMessageBox, QPushButton,
                              QTextEdit, QVBoxLayout, QWidget)

from . import document_features as features
from .design import DesignSettings, apply_design, theme_shape_colors
from .editor import _image_to_data_uri, pt_to_px
from .models import PageSettings
from .publishing import Publication


def editable(method):
    arity = len(inspect.signature(method).parameters)
    @wraps(method)
    def invoke(self, *args, **kwargs):
        if self._can_edit():
            if arity == 1 and len(args) == 1 and isinstance(args[0], bool):
                args = ()
            return method(self, *args, **kwargs)
        return None
    return invoke


class FeatureCommands:
    def _initialize_features(self):
        self.hooks = features.CollaborationHooks(self)
        self.hooks.transcriptReady.connect(self._insert_transcript)
        self._format_sample = None
        self._fields_updating = False
        self._fields_timer = QTimer(self)
        self._fields_timer.setInterval(60_000)
        self._fields_timer.timeout.connect(self._refresh_fields)
        self._fields_timer.start()
        self.editor.document().folio_state = self.state
        self.editor.folio_state = self.state

    def _feature_changed(self):
        self._set_dirty(True)
        self._autosave_timer.start()
        self.canvas.viewport().update()
        self._outline_timer.start()

    def _can_edit(self):
        return self.hooks.mode != "Viewing"

    def _permission_mode(self, mode):
        self.hooks.set_mode(mode)
        self.editor.setReadOnly(mode == "Viewing")
        if mode == "Reviewing":
            self._toggle_tracking(True)
            with QSignalBlocker(self.track_button):
                self.track_button.setChecked(True)
        elif mode == "Editing":
            self._toggle_tracking(False)
            with QSignalBlocker(self.track_button):
                self.track_button.setChecked(False)
        for index in range(1, 7):
            self.ribbon.widget(index).setEnabled(mode != "Viewing")
        if hasattr(self, "floating_dock"):
            self.floating_dock.widget().setEnabled(mode != "Viewing")
        if hasattr(self, "_find_dialog") and self._find_dialog:
            self._find_dialog.close()
        self.statusBar().showMessage(f"{mode} mode", 3000)

    def _sample_format(self):
        self._format_sample = self.editor.sample_inline_format()
        self.statusBar().showMessage("Format sampled. Select a target and choose Apply Format.", 5000)

    def _apply_format(self):
        if self._can_edit() and self._format_sample is not None:
            self.editor.apply_inline_format(self._format_sample)

    @editable
    def _paragraph_shading(self):
        from PySide6.QtWidgets import QColorDialog
        color = QColorDialog.getColor(QColor("#dbeafe"), self, "Paragraph shading")
        if color.isValid():
            self.editor.set_paragraph_shading(color)

    @editable
    def _semantic_design(self):
        from .feature_dialogs import DesignDialog
        previous = DesignSettings.from_dict(self.state.design)
        dialog = DesignDialog(previous, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        settings = dialog.settings()
        apply_design(self.editor.document(), settings, previous)
        self.state.design = settings.to_dict()
        self.editor.document().design_settings = settings
        self.editor.design_settings = settings
        tokens = theme_shape_colors(settings)
        for obj in self.state.features["objects"]:
            if obj.get("theme_bound"):
                obj.update(fill=tokens["fill"], stroke=tokens["stroke"], color=tokens["text"])
        self._feature_changed()

    @editable
    def _blank_page(self):
        self.editor.insert_page_break()
        cursor = self.editor.textCursor()
        cursor.setKeepPositionOnInsert(True)
        self.editor.insert_page_break()
        self.editor.setTextCursor(cursor)

    @editable
    def _equation_editor(self):
        from .structured_media import EquationDialog
        selected = self._selected_node("equation")
        node = selected[1] if selected else {}
        dialog = EquationDialog(self, node.get("source", ""), node.get("syntax", "tex"))
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if selected:
            self._select_node_range(selected[0])
        features.insert_node(self.editor, self.state,
                             {"kind": "equation", "source": dialog.source(), "syntax": dialog.syntax()},
                             image=dialog.rendered_image())

    @editable
    def _model_insert(self):
        from .structured_media import load_model, render_model
        path, _ = QFileDialog.getOpenFileName(self, "Insert 3D model", "", "Models (*.gltf *.glb)")
        if not path:
            return
        try:
            scene = load_model(path)
            features.insert_node(self.editor, self.state,
                                 {"kind": "model", "name": Path(path).name, "scene": scene.to_dict()},
                                 image=render_model(scene))
            self.statusBar().showMessage("Embedded a static 3D mesh view; the mesh is retained in Folio.", 7000)
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "3D model", str(exc))

    @editable
    def _bookmark(self):
        name, ok = QInputDialog.getText(self, "Bookmark", "Name for the selection or current paragraph:")
        if ok:
            try:
                features.bookmark(self.editor, self.state, name)
                self._feature_changed()
            except ValueError as exc:
                QMessageBox.warning(self, "Bookmark", str(exc))

    @editable
    def _cross_reference(self):
        nodes = self.state.features["nodes"]
        ranges = features.node_ranges(self.editor.document())
        choices = [(n.get("name", key), key) for key, n in nodes.items()
                   if n["kind"] == "bookmark" and key in ranges]
        block = self.editor.document().begin()
        headings = []
        while block.isValid():
            if block.blockFormat().headingLevel():
                headings.append((f"Heading: {block.text()}", block.position(), block.length() - 1))
            block = block.next()
        for name, position, length in headings:
            choices.append((name, (position, length)))
        if not choices:
            QMessageBox.information(self, "Cross-reference", "Add a heading or bookmark a figure or text range first.")
            return
        labels = [f"{index + 1}. {label}" for index, (label, _) in enumerate(choices)]
        label, ok = QInputDialog.getItem(self, "Cross-reference", "Target:", labels, editable=False)
        if not ok:
            return
        display, ok = QInputDialog.getItem(self, "Cross-reference", "Display:", ["Text", "Page number"], editable=False)
        if not ok:
            return
        target = choices[labels.index(label)][1]
        if isinstance(target, tuple):
            position, length = target
            cursor = QTextCursor(self.editor.document())
            cursor.setPosition(position)
            cursor.setPosition(position + length, QTextCursor.MoveMode.KeepAnchor)
            key = features.uid()
            nodes[key] = {"kind": "bookmark", "name": "Heading_" + key}
            features.mark_cursor(cursor, key)
            target = key
        features.insert_node(self.editor, self.state,
                             {"kind": "reference", "target": target,
                              "display": "page" if display == "Page number" else "text"}, "Reference")
        self._refresh_fields()

    @editable
    def _date_field(self):
        label, ok = QInputDialog.getItem(self, "Date and time", "Auto-updating field:",
                                        ["Date", "Time", "Date and time"], editable=False)
        if ok:
            style = {"Date": "date", "Time": "time", "Date and time": "datetime"}[label]
            features.insert_node(self.editor, self.state, {"kind": "datetime", "format": style}, "Date")
            self._refresh_fields()

    def _refresh_fields(self):
        if self._fields_updating or not self._can_edit():
            return
        self._fields_updating = True
        try:
            if not any(n["kind"] in ("datetime", "reference") for n in self.state.features["nodes"].values()):
                return
            publication = Publication(self.editor.document(), self.state.page, self._title)
            with self.tracker.applying():
                features.refresh_fields(self.editor.document(), self.state.features,
                                        page_for_position=publication.page_for_position,
                                        can_edit=lambda start, end: self.tracker.allows_edit(start, end, formatting=True))
        finally:
            self._fields_updating = False

    @editable
    def _attachment_insert(self):
        path, _ = QFileDialog.getOpenFileName(self, "Embed document or file (stored; never executed)")
        if not path:
            return
        try:
            node = features.attachment_node(Path(path))
            features.insert_node(self.editor, self.state, node, f"[Attached file: {node['name']}]")
        except (ValueError, OSError) as exc:
            QMessageBox.warning(self, "Embed file", str(exc))

    def _selected_node(self, kind=None):
        cursor = self.editor.textCursor()
        candidates = []
        for key, (start, end) in features.node_ranges(self.editor.document()).items():
            node = self.state.features["nodes"].get(key)
            if node and (kind is None or node["kind"] == kind):
                if start <= cursor.position() <= end:
                    candidates.append((node["kind"] == "bookmark", end - start, key, node))
        if candidates:
            _, _, key, node = min(candidates, key=lambda item: item[:2])
            return key, node
        return None

    def _select_node_range(self, key):
        start, end = features.node_ranges(self.editor.document())[key]
        cursor = self.editor.textCursor()
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        self.editor.setTextCursor(cursor)
        return cursor

    @editable
    def _edit_content_node(self):
        selected = self._selected_node()
        if selected is None:
            self.statusBar().showMessage("Place the cursor in an equation, diagram, signature, or attachment.", 5000)
            return
        key, node = selected
        if node["kind"] == "equation":
            self._equation_editor()
        elif node["kind"] == "signature":
            self._sign_field(key, node)
        elif node["kind"] == "diagram":
            from .dialogs import DiagramDialog
            dialog = DiagramDialog(self)
            dialog.steps.setPlainText("\n".join(node["steps"]))
            if dialog.exec() == QDialog.DialogCode.Accepted:
                self._select_node_range(key)
                steps = dialog.values()
                features.insert_node(self.editor, self.state, {"kind": "diagram", "steps": steps},
                                     image=DiagramDialog.render(steps))
        elif node["kind"] == "attachment":
            path, _ = QFileDialog.getSaveFileName(self, "Extract embedded file", node["name"])
            if path:
                from .docx_io import atomic_write
                try:
                    atomic_write(Path(path), base64.b64decode(node["data"]))
                except OSError as exc:
                    QMessageBox.warning(self, "Extract file", str(exc))

    @editable
    def _signature_node(self):
        name, ok = QInputDialog.getText(self, "Signature field", "Expected signer (optional):")
        if ok:
            features.insert_node(self.editor, self.state, {"kind": "signature", "signer": name},
                                 f"[Signature: {name or 'Signer'} — unsigned]")

    @editable
    def _sign_field(self, key, node):
        from .feature_dialogs import SignatureInputDialog
        dialog = SignatureInputDialog(self, name=node.get("signer", ""))
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        features.sign_node(node, dialog.signature_name(), dialog.consent(), self.editor.document())
        cursor = self._select_node_range(key)
        fmt = cursor.charFormat()
        cursor.insertText(f"Signed by {node['signed_by']} · {node['signed_at']}", fmt)
        self._feature_changed()

    @editable
    def _float_insert(self, image=False):
        tokens = theme_shape_colors(DesignSettings.from_dict(self.state.design))
        obj = {"id": features.uid(), "kind": "image" if image else "text",
               "name": "Picture" if image else "Text box", "page": 0,
               "x": 130.0, "y": 160.0, "width": 220.0, "height": 110.0,
               "wrap": "square", "visible": True, "z": len(self.state.features["objects"]),
               "fill": tokens["fill"], "stroke": tokens["stroke"], "color": tokens["text"],
               "theme_bound": True}
        if image:
            path, _ = QFileDialog.getOpenFileName(self, "Floating image", "", "Images (*.png *.jpg *.jpeg *.svg)")
            if not path:
                return
            try:
                if Path(path).suffix.lower() == ".svg":
                    from .structured_media import render_svg
                    raster = render_svg(Path(path).read_bytes())
                else:
                    raster = QImage(path)
                if raster.isNull():
                    raise ValueError("This image could not be read.")
                obj["image"] = _image_to_data_uri(raster)
                obj["height"] = obj["width"] * raster.height() / raster.width()
            except (ValueError, OSError) as exc:
                QMessageBox.warning(self, "Floating image", str(exc))
                return
        else:
            text, ok = QInputDialog.getMultiLineText(self, "Text box", "Text:")
            if not ok:
                return
            obj["text"] = text
        self.state.features["objects"].append(obj)
        self._feature_changed()
        self._floating_inspector()

    def _floating_inspector(self):
        from .feature_dialogs import FloatingInspector
        if not hasattr(self, "floating_dock"):
            self.floating_dock = QDockWidget("Selection Pane — Floating Objects", self)
            self.floating_dock.setObjectName("floating_objects")
            self.floating_dock.setWidget(FloatingInspector(self.state.features["objects"], self._feature_changed, self))
            self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.floating_dock)
        self.floating_dock.widget().set_objects(self.state.features["objects"])
        self.floating_dock.widget().setEnabled(self._can_edit())
        self.floating_dock.show()

    @editable
    def _drop_cap(self):
        cursor = self.editor.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.StartOfBlock)
        cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
        text = cursor.selectedText()
        if not text.strip():
            return
        publication = Publication(self.editor.document(), self.state.page, self._title)
        page = publication.page_for_position(cursor.selectionStart()) - 1
        rect = self.editor.document().documentLayout().blockBoundingRect(cursor.block())
        cursor.removeSelectedText()
        self.state.features["objects"].append({"id": features.uid(), "kind": "text", "name": "Drop cap",
            "text": text, "page": page, "x": publication.margin_left,
            "y": publication.margin_top + rect.top() % publication.body_h,
            "width": 55, "height": 64, "font_size": 44, "fill": "#ffffff", "stroke": "#ffffff",
            "wrap": "square", "z": 0, "visible": True})
        self._feature_changed()

    @editable
    def _section_break(self, kind):
        from .dialogs import PageSetupDialog
        dialog = PageSetupDialog(self._active_page_settings(), self)
        dialog.setWindowTitle(f"{kind.title()} section — page setup")
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.settings():
            try:
                features.insert_section(self.editor, self.state, kind, dialog.settings())
                self._feature_changed()
            except ValueError as exc:
                QMessageBox.warning(self, "Section break", str(exc))

    def _active_section(self):
        position = self.editor.textCursor().position()
        ranges = features.node_ranges(self.editor.document())
        preceding = [(ranges[s["id"]][0], s) for s in self.state.features["sections"]
                     if s["id"] in ranges and ranges[s["id"]][0] <= position]
        return max(preceding, key=lambda entry: entry[0])[1] if preceding else None

    def _active_page_settings(self):
        section = self._active_section()
        return PageSettings.from_dict(section["page"]) if section else copy.copy(self.state.page)

    @editable
    def _set_active_page(self, settings):
        settings.validate()
        section = self._active_section()
        if section:
            section["page"] = settings.to_dict()
        else:
            self.state.page = settings
        self.editor.apply_page_settings(settings)
        from .editor import mm_to_px
        self.canvas.set_page_metrics(*(mm_to_px(value) for value in settings.size_mm()))
        self._sync_page_controls()
        self._feature_changed()

    @editable
    def _table_dimensions(self):
        from .feature_dialogs import TableDimensionsDialog
        table = self.editor.textCursor().currentTable()
        if table is None:
            return
        dialog = TableDimensionsDialog(table.columns(), table.rows(), parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        widths, heights = dialog.dimensions()
        cursor = self.editor.textCursor()
        cursor.beginEditBlock()
        if widths:
            fmt = table.format()
            fmt.setColumnWidthConstraints([QTextLength(QTextLength.Type.FixedLength, pt_to_px(value)) for value in widths])
            table.setFormat(fmt)
        if heights:
            for row, height in enumerate(heights):
                for column in range(table.columns()):
                    cell = table.cellAt(row, column)
                    block = cell.firstCursorPosition().block()
                    fmt = block.blockFormat()
                    fmt.setLineHeight(pt_to_px(height), 3)
                    QTextCursor(block).setBlockFormat(fmt)
        cursor.endEditBlock()

    def _dictation(self):
        if not self._can_edit():
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Transcription input")
        layout = QVBoxLayout(dialog)
        note = QLabel("Local transcription integration: a speech provider can submit text through "
                      "window.hooks.submit_transcript(). No audio is recorded or uploaded by Folio. "
                      "You can also paste a transcript below.")
        note.setWordWrap(True)
        layout.addWidget(note)
        text = QTextEdit()
        layout.addWidget(text)
        insert = QPushButton("Insert transcript")
        insert.clicked.connect(lambda: (self.hooks.submit_transcript(text.toPlainText()), dialog.accept()))
        layout.addWidget(insert)
        self.hooks.set_dictation_state("listening")
        try:
            dialog.exec()
        finally:
            self.hooks.set_dictation_state("idle")

    def _insert_transcript(self, text):
        if self._can_edit():
            self.editor.textCursor().insertText(text)
