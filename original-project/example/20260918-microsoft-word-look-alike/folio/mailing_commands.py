from __future__ import annotations

from PySide6.QtWidgets import QDialog, QDialogButtonBox, QInputDialog, QLabel, QMenu, QMessageBox, QTextBrowser, QVBoxLayout

from . import mailmerge
from .mailing_dialogs import MailingWorkbenchDialog, MappingDialog, RuleDialog


class MailingCommands:
    def _build_mailings_tab(self):
        groups = []
        group, row = self._group("Create")
        row.addWidget(self._tool("page", "Envelopes", self._mailing_envelope, "Choose a standard envelope format", text_mode=True))
        row.addWidget(self._tool("table", "Labels", self._mailing_labels, "Choose an adhesive label sheet", text_mode=True))
        groups.append(group)
        group, row = self._group("Recipients")
        row.addWidget(self._tool("open", "Select Recipients", self._mailing_recipients, "Import and manage CSV or XLSX recipients", text_mode=True))
        row.addWidget(self._tool("field", "Match Fields", self._mailing_mapping, "Map source columns to merge and address fields", text_mode=True))
        groups.append(group)
        group, row = self._group("Write & Insert")
        row.addWidget(self._tool("field", "Insert Field", self._mailing_insert_field, "Insert a recipient column", text_mode=True))
        menu = QMenu(self)
        menu.addAction("Address Block", lambda: self._mailing_insert_token("{{AddressBlock}}"))
        menu.addAction("Greeting Line", lambda: self._mailing_insert_token("{{GreetingLine}}"))
        menu.addAction("If / Skip If Rule", self._mailing_rule)
        menu.addAction("Update Labels", self._mailing_update_labels)
        row.addWidget(self._menu_tool("Blocks & Rules", menu))
        groups.append(group)
        group, row = self._group("Preview & Finish")
        row.addWidget(self._tool("preview", "Preview Results", lambda: self._mailing_workbench(1), "Preview each active recipient", text_mode=True))
        row.addWidget(self._tool("merge", "Finish & Merge", lambda: self._mailing_workbench(2), "DOCX, PDF, print batches, or email draft files", text_mode=True))
        groups.append(group)
        self.ribbon.addTab(self._wrap(groups), "Mailings")

    def _mailing_dataset(self):
        if getattr(self, "_mailing_document_id", None) != self.state.id:
            self._mailing_document_id = self.state.id
            self._mailing_records = mailmerge.RecipientSet([])
        return self._mailing_records

    def _mailing_config(self):
        return mailmerge.checked_mailing(self.state.features.get("mailing"))

    def _mailing_save_config(self, config):
        current = self._mailing_config()
        checked = mailmerge.checked_mailing(config)
        if current != checked:
            self.state.features["mailing"] = checked
            self._set_dirty(True)

    def _mailing_columns(self):
        records = self._mailing_dataset().records
        return list(records[0]) if records else []

    def _mailing_workbench(self, tab=0):
        dialog = MailingWorkbenchDialog(self.editor.document(), self.state, self._mailing_dataset(), self)
        dialog.tabs.setCurrentIndex(tab)
        dialog.exec()
        self._mailing_records = dialog.recipients
        config = self._mailing_config()
        changed = config.get("mapping", {}) != dialog.mapping
        config["mapping"] = dialog.mapping
        if changed and self._can_edit():
            self._mailing_save_config(config)

    def _mailing_recipients(self):
        self._mailing_workbench(0)

    def _mail_merge(self):
        self._mailing_workbench(2)

    def _mailing_mapping(self):
        if not self._can_edit():
            return
        config = self._mailing_config()
        dialog = MappingDialog(self._mailing_columns(), config.get("mapping", {}),
                               mailmerge.discover_fields(self.editor.document()), self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            config["mapping"] = dialog.mapping()
            self._mailing_save_config(config)

    def _mailing_insert_token(self, token):
        if not self._can_edit():
            return
        self.editor.textCursor().insertText(token)

    def _mailing_insert_field(self):
        if not self._can_edit():
            return
        columns = self._mailing_columns()
        name, ok = QInputDialog.getItem(self, "Insert Merge Field", "Recipient column:", columns, 0, True)
        if ok and name.strip():
            name = name.strip()
            if any(char in name for char in "{}\r\n") or name.startswith(("IF|", "SKIPIF|")):
                QMessageBox.warning(self, "Merge Field", "Map this column to a name without braces, newlines, or rule prefixes.")
                return
            self._mailing_insert_token("{{" + name + "}}")

    def _mailing_rule(self):
        if not self._can_edit():
            return
        fields = list(dict.fromkeys([*self._mailing_columns(), *self._mailing_config().get("mapping", {})]))
        dialog = RuleDialog(fields, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._mailing_insert_token(dialog.token)

    def _mailing_envelope(self):
        if not self._can_edit():
            return
        name, ok = QInputDialog.getItem(self, "Envelope Format",
            "Apply envelope dimensions and delivery-address margins:", mailmerge.ENVELOPE_TEMPLATES, 0, False)
        if ok:
            try:
                self._set_active_page(mailmerge.envelope_settings(name))
                config = self._mailing_config()
                config["label_template"] = ""
                self._mailing_save_config(config)
                if not self.editor.toPlainText().strip():
                    self._mailing_insert_token("{{AddressBlock}}")
                self.statusBar().showMessage("Envelope dimensions applied. Insert an Address Block and match recipient fields.", 6000)
            except ValueError as exc:
                QMessageBox.warning(self, "Envelope", str(exc))

    def _mailing_labels(self):
        if not self._can_edit():
            return
        names = ["Letter / ordinary document", *mailmerge.LABEL_TEMPLATES]
        name, ok = QInputDialog.getItem(self, "Label Sheet",
            "Edit one label in the document; Finish & Merge advances recipients across each sheet:", names, 0, False)
        if not ok:
            return
        config = self._mailing_config()
        config["label_template"] = "" if name == names[0] else name
        self._mailing_save_config(config)
        if config["label_template"] and not self.editor.toPlainText().strip():
            self._mailing_insert_token("{{AddressBlock}}")
        self.statusBar().showMessage("Label format saved. Update Labels previews the sheet; final sheets preserve dimensions as page images.", 7000)

    def _mailing_update_labels(self):
        config = self._mailing_config()
        name = config.get("label_template")
        if not name:
            self._mailing_labels()
            config = self._mailing_config()
            name = config.get("label_template")
        if not name:
            return
        try:
            template = mailmerge.LABEL_TEMPLATES[name]
            records = [record for _, record in self._mailing_dataset().active()]
            if records:
                documents = [row[2] for row in mailmerge.prepared_records(self.editor.document(), records, config.get("mapping"))]
            else:
                documents = [self.editor.document()] * template.capacity
            sheets = mailmerge.label_sheets(documents[:template.capacity], template)
        except mailmerge.MergeError as exc:
            QMessageBox.warning(self, "Update Labels", str(exc))
            return
        dialog = QDialog(self)
        dialog.setWindowTitle("Updated Label Sheet Preview")
        layout = QVBoxLayout(dialog)
        note = QLabel("First sheet: " + name + ". Edit the label in the document and run Update Labels again to refresh.")
        note.setWordWrap(True)
        layout.addWidget(note)
        preview = QTextBrowser()
        preview.setHtml(sheets.toHtml())
        preview.setOpenExternalLinks(False)
        layout.addWidget(preview)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        box.rejected.connect(dialog.reject)
        layout.addWidget(box)
        dialog.resize(940, 750)
        dialog.exec()
