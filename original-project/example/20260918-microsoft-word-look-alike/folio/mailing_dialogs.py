from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QDialog,
    QDialogButtonBox, QFileDialog, QFormLayout, QHBoxLayout, QLabel, QLineEdit,
    QMessageBox, QProgressBar, QPushButton, QSpinBox, QTabWidget, QTableWidget,
    QTableWidgetItem, QTextBrowser, QVBoxLayout, QWidget)

from . import mailmerge


class MappingDialog(QDialog):
    def __init__(self, columns, mapping, fields=(), parent=None):
        super().__init__(parent)
        self.setWindowTitle("Match Merge Fields")
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Match document and address fields to recipient columns."))
        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Document field", "Recipient column"])
        self.table.horizontalHeader().setStretchLastSection(True)
        names = list(dict.fromkeys([*fields, *mailmerge.ADDRESS_FIELDS, "Email"]))
        self.selectors = {}
        for name in names:
            row = self.table.rowCount()
            self.table.insertRow(row)
            item = QTableWidgetItem(name)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, item)
            selector = QComboBox()
            selector.addItem("(Unmapped)", "")
            for column in columns:
                selector.addItem(column, column)
            source = mapping.get(name, name if name in columns else "")
            selector.setCurrentIndex(max(0, selector.findData(source)))
            self.table.setCellWidget(row, 1, selector)
            self.selectors[name] = selector
        layout.addWidget(self.table)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)
        self.resize(480, 500)

    def mapping(self):
        return {name: selector.currentData() for name, selector in self.selectors.items()
                if selector.currentData()}


class RuleDialog(QDialog):
    def __init__(self, fields, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Merge Rule")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.kind = QComboBox()
        self.kind.addItems(["If Then Else", "Skip If"])
        self.field = QComboBox()
        self.field.setEditable(True)
        self.field.addItems(list(fields))
        self.operator = QComboBox()
        self.operator.addItems(mailmerge.OPERATORS)
        self.value = QLineEdit()
        self.then_text = QLineEdit()
        self.else_text = QLineEdit()
        for name, widget in (("Rule", self.kind), ("Field", self.field),
                ("Comparison", self.operator), ("Compare to", self.value),
                ("Then insert", self.then_text), ("Otherwise insert", self.else_text)):
            form.addRow(name, widget)
        layout.addLayout(form)
        note = QLabel("Text comparisons ignore case. Greater/less comparisons use numbers. Rule output is literal text.")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.kind.currentIndexChanged.connect(lambda index: (
            self.then_text.setEnabled(index == 0), self.else_text.setEnabled(index == 0)))
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)
        self.token = ""

    def _accept(self):
        try:
            self.token = mailmerge.rule_token("IF" if self.kind.currentIndex() == 0 else "SKIPIF",
                mailmerge.Condition(self.field.currentText().strip(), self.operator.currentText(), self.value.text()),
                self.then_text.text(), self.else_text.text())
        except mailmerge.MergeError as exc:
            QMessageBox.warning(self, "Merge Rule", str(exc))
            return
        self.accept()


class MailingWorkbenchDialog(QDialog):
    def __init__(self, document, state, recipients=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Mail Merge Recipients, Preview and Output")
        self.document = document
        self.state = state
        self.recipients = recipients or mailmerge.RecipientSet([])
        self.config = mailmerge.checked_mailing(state.features.get("mailing"))
        self.mapping = self.config.get("mapping", {})
        self.outputs = []
        self._refreshing = False
        self._cancelled = False
        self._running = False
        layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)
        page = QWidget()
        content = QVBoxLayout(page)
        source_row = QHBoxLayout()
        self.load_button = QPushButton("Select CSV / XLSX")
        self.load_button.clicked.connect(self.choose_source)
        self.map_button = QPushButton("Match Fields")
        self.map_button.clicked.connect(self.match_fields)
        source_row.addWidget(self.load_button)
        source_row.addWidget(self.map_button)
        source_row.addStretch()
        content.addLayout(source_row)
        sort_row = QHBoxLayout()
        sort_row.addWidget(QLabel("Sort by"))
        self.sort_column = QComboBox()
        self.sort_column.addItem("Source order", "")
        self.reverse = QCheckBox("Descending")
        sort_row.addWidget(self.sort_column)
        sort_row.addWidget(self.reverse)
        sort_row.addStretch()
        content.addLayout(sort_row)
        self.sort_column.currentIndexChanged.connect(self.refresh_records)
        self.reverse.toggled.connect(self.refresh_records)
        filter_row = QHBoxLayout()
        self.filter_column = QComboBox()
        self.filter_column.addItem("All records", "")
        self.filter_operator = QComboBox()
        self.filter_operator.addItems(mailmerge.OPERATORS)
        self.filter_value = QLineEdit()
        apply_filter = QPushButton("Apply Filter")
        apply_filter.clicked.connect(self.refresh_records)
        for widget in (self.filter_column, self.filter_operator, self.filter_value, apply_filter):
            filter_row.addWidget(widget)
        content.addLayout(filter_row)
        self.records_table = QTableWidget()
        self.records_table.itemChanged.connect(self._record_changed)
        content.addWidget(self.records_table)
        self.record_count = QLabel()
        content.addWidget(self.record_count)
        self.tabs.addTab(page, "Recipients")
        page = QWidget()
        preview_layout = QVBoxLayout(page)
        navigation = QHBoxLayout()
        previous = QPushButton("Previous")
        following = QPushButton("Next")
        self.record_number = QSpinBox()
        self.record_number.setMinimum(1)
        self.preview_status = QLabel()
        for widget in (previous, self.record_number, following, self.preview_status):
            navigation.addWidget(widget)
        navigation.addStretch()
        previous.clicked.connect(lambda: self.record_number.setValue(self.record_number.value() - 1))
        following.clicked.connect(lambda: self.record_number.setValue(self.record_number.value() + 1))
        self.record_number.valueChanged.connect(self.refresh_preview)
        preview_layout.addLayout(navigation)
        self.preview = QTextBrowser()
        self.preview.setOpenExternalLinks(False)
        preview_layout.addWidget(self.preview)
        self.tabs.addTab(page, "Preview Results")
        page = QWidget()
        export_layout = QVBoxLayout(page)
        form = QFormLayout()
        self.output_format = QComboBox()
        for title, kind in (("Individual DOCX files", "docx"), ("Individual PDF files", "pdf"),
                ("One compound DOCX file", "compound_docx"), ("One compound PDF file", "compound_pdf"),
                ("Email drafts (.eml files, no sending)", "eml"), ("Print batch", "print")):
            self.output_format.addItem(title, kind)
        self.email_column = QComboBox()
        self.subject = QLineEdit(state.title)
        form.addRow("Output", self.output_format)
        form.addRow("Email column", self.email_column)
        form.addRow("Email subject", self.subject)
        export_layout.addLayout(form)
        self.output_format.currentIndexChanged.connect(self._output_changed)
        self._output_changed()
        self.export_note = QLabel("Excluded recipients and matching Skip If rules are omitted. Existing output files are preserved.")
        self.export_note.setWordWrap(True)
        export_layout.addWidget(self.export_note)
        if self.config.get("label_template"):
            note = QLabel("Labels: " + self.config["label_template"] +
                ". Recipients advance across each sheet. Sheet output uses embedded page images to preserve dimensions.")
            note.setWordWrap(True)
            export_layout.addWidget(note)
        self.run_button = QPushButton("Finish & Merge")
        self.run_button.clicked.connect(self.run_merge)
        export_layout.addWidget(self.run_button)
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        export_layout.addWidget(self.progress)
        export_layout.addStretch()
        self.tabs.addTab(page, "Finish & Merge")
        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.reject)
        layout.addWidget(self.close_button)
        self.resize(850, 600)
        self._populate_columns()

    def _output_changed(self):
        enabled = self.output_format.currentData() == "eml"
        self.email_column.setEnabled(enabled)
        self.subject.setEnabled(enabled)

    def columns(self):
        return list(self.recipients.records[0]) if self.recipients.records else []

    def choose_source(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select recipients", "", "Recipient data (*.csv *.xlsx)")
        if not path:
            return
        try:
            records = mailmerge.read_records(Path(path))
            if not records:
                raise mailmerge.MergeError("The source contains no recipient rows.")
        except (OSError, ValueError, mailmerge.MergeError) as exc:
            QMessageBox.warning(self, "Recipients", str(exc))
            return
        self.recipients = mailmerge.RecipientSet(records)
        self._populate_columns()

    def _populate_columns(self):
        self._refreshing = True
        for widget, initial in ((self.sort_column, "Source order"), (self.filter_column, "All records"),
                                (self.email_column, "Select email column")):
            widget.clear()
            widget.addItem(initial, "")
            for name in self.columns():
                widget.addItem(name, name)
        self.sort_column.setCurrentIndex(max(0, self.sort_column.findData(self.recipients.sort_field)))
        self.reverse.setChecked(self.recipients.descending)
        self.email_column.setCurrentIndex(max(0, self.email_column.findData("Email")))
        if self.recipients.filters:
            rule = self.recipients.filters[0]
            self.filter_column.setCurrentIndex(max(0, self.filter_column.findData(rule.field)))
            self.filter_operator.setCurrentText(rule.operator)
            self.filter_value.setText(rule.value)
        self._refreshing = False
        self.refresh_records()

    def refresh_records(self):
        if self._refreshing:
            return
        self.recipients.sort_field = self.sort_column.currentData() or ""
        self.recipients.descending = self.reverse.isChecked()
        field = self.filter_column.currentData()
        self.recipients.filters = ([mailmerge.Condition(field, self.filter_operator.currentText(), self.filter_value.text())]
                                   if field else [])
        visible = mailmerge.RecipientSet(self.recipients.records, filters=self.recipients.filters,
            sort_field=self.recipients.sort_field, descending=self.recipients.descending)
        try:
            rows = visible.active()
        except mailmerge.MergeError as exc:
            self.record_count.setText(str(exc))
            self.run_button.setEnabled(False)
            return
        self._refreshing = True
        columns = self.columns()
        self.records_table.clear()
        self.records_table.setColumnCount(len(columns) + 1)
        self.records_table.setHorizontalHeaderLabels(["Include", *columns])
        self.records_table.setRowCount(len(rows))
        for row, (index, record) in enumerate(rows):
            checkbox = QTableWidgetItem()
            checkbox.setData(Qt.ItemDataRole.UserRole, index)
            checkbox.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsUserCheckable)
            checkbox.setCheckState(Qt.CheckState.Unchecked if index in self.recipients.excluded else Qt.CheckState.Checked)
            self.records_table.setItem(row, 0, checkbox)
            for column, key in enumerate(columns, 1):
                item = QTableWidgetItem(record[key])
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.records_table.setItem(row, column, item)
        self._refreshing = False
        self._update_counts()

    def _record_changed(self, item):
        if self._refreshing or item.column() != 0:
            return
        index = item.data(Qt.ItemDataRole.UserRole)
        if item.checkState() == Qt.CheckState.Checked:
            self.recipients.excluded.discard(index)
        else:
            self.recipients.excluded.add(index)
        self._update_counts()

    def _update_counts(self):
        active = self.recipients.active()
        self.record_count.setText(f"{len(active)} active of {len(self.recipients.records)} recipients")
        self.run_button.setEnabled(bool(active))
        self.record_number.setMaximum(max(1, len(active)))
        self.refresh_preview()

    def match_fields(self):
        dialog = MappingDialog(self.columns(), self.mapping, mailmerge.discover_fields(self.document), self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.mapping = dialog.mapping()
            self.refresh_preview()

    def refresh_preview(self):
        if self._refreshing:
            return
        try:
            active = self.recipients.active()
            if not active:
                self.preview.clear()
                self.preview_status.setText("No active recipients")
                return
            ordinal = min(self.record_number.value() - 1, len(active) - 1)
            source_index, record = active[ordinal]
            doc = mailmerge.merge_document(self.document, record, self.mapping)
            self.preview.setHtml(doc.toHtml())
            self.preview_status.setText(f"{ordinal + 1} of {len(active)}; source row {source_index + 2}")
        except mailmerge.MergeError as exc:
            self.preview.clear()
            self.preview_status.setText(str(exc))

    def run_merge(self):
        try:
            records = [record for _, record in self.recipients.active()]
            mailmerge.prepared_records(self.document, records, self.mapping)
            kind = self.output_format.currentData()
            if kind == "print":
                from PySide6.QtPrintSupport import QPrintDialog, QPrinter

                printer = QPrinter(QPrinter.PrinterMode.HighResolution)
                dialog = QPrintDialog(printer, self)
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    return
                count = mailmerge.print_batch(self.document, self.state, records, printer,
                    mapping=self.mapping, labels=self.config.get("label_template", ""))
                QMessageBox.information(self, "Mail Merge", f"Sent {count} pages to the selected printer.")
                return
            directory = QFileDialog.getExistingDirectory(self, "Choose output folder")
            if not directory:
                return
            self._running = True
            self._cancelled = False
            self.tabs.setEnabled(False)
            self.close_button.setText("Cancel batch")
            self.progress.setVisible(True)
            def progress(done, total):
                self.progress.setRange(0, total)
                self.progress.setValue(done)
                QApplication.processEvents()
            self.outputs = mailmerge.merge_batch(self.document, self.state, records, Path(directory),
                output=kind, mapping=self.mapping, labels=self.config.get("label_template", ""),
                email_field=self.email_column.currentData() or "Email", subject=self.subject.text(),
                progress=progress, cancelled=lambda: self._cancelled)
            QMessageBox.information(self, "Mail Merge", f"Created {len(self.outputs)} files in {directory}." +
                (" Batch cancelled." if self._cancelled else ""))
        except Exception as exc:
            QMessageBox.warning(self, "Mail Merge", str(exc))
        finally:
            self._running = False
            self.tabs.setEnabled(True)
            self.close_button.setText("Close")

    def reject(self):
        if self._running:
            self._cancelled = True
            return
        super().reject()

    def closeEvent(self, event):
        if self._running:
            self._cancelled = True
            event.ignore()
            return
        super().closeEvent(event)
