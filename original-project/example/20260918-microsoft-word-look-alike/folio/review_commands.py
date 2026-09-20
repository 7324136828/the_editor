from __future__ import annotations

import uuid
from pathlib import Path

from PySide6.QtCore import QSignalBlocker, Qt
from PySide6.QtGui import QColor, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox,
    QFileDialog, QHBoxLayout, QInputDialog, QLabel, QListWidget,
    QListWidgetItem, QMenu, QMessageBox, QPushButton, QTextEdit,
    QVBoxLayout,
)

from .editor import SafeDocument
from .review import ReviewError
from .review_dialogs import CommentThreadsDialog, InkDialog
from .review_services import (
    SpeechReader, accessibility_issues, compare_three_way, load_glossary,
    review_document, translate_glossary,
)


class ReviewCommands:
    def _review_extra_groups(self):
        groups = []
        group, row = self._group("Accessibility & Language")
        row.addWidget(self._tool("spell", "Accessibility", self._accessibility_check, text_mode=True))
        row.addWidget(self._tool("find", "Glossary translation", self._translate_local, text_mode=True))
        row = self._extra_row(group)
        speech = QMenu(self)
        for label, slot in (("Read selection or document", self._read_aloud),
                            ("Pause", lambda: self.speech_reader.pause()),
                            ("Resume", lambda: self.speech_reader.resume()),
                            ("Stop", lambda: self.speech_reader.stop())):
            speech.addAction(label, slot)
        row.addWidget(self._menu_tool("Read Aloud", speech))
        groups.append(group)
        group, row = self._group("Threads & Display")
        row.addWidget(self._tool("comment", "Threads", self._comment_threads, text_mode=True))
        self.review_display = QComboBox()
        self.review_display.addItem("Review display…")
        self.review_display.addItems(["All Markup", "Simple Markup", "No Markup", "Original"])
        self.review_display.activated.connect(self._review_display_selected)
        row.addWidget(self.review_display)
        row = self._extra_row(group)
        self.format_tracking = QCheckBox("Track formatting")
        self.format_tracking.setToolTip("Record formatting changes while Track Changes is enabled")
        self.format_tracking.toggled.connect(self._track_formatting)
        row.addWidget(self.format_tracking)
        groups.append(group)
        group, row = self._group("Compare & Protect")
        row.addWidget(self._tool("review", "Three-way compare", self._three_way_compare, text_mode=True))
        row.addWidget(self._tool("author", "Restrictions", self._editing_restrictions, text_mode=True))
        row = self._extra_row(group)
        row.addWidget(self._tool("image", "Draw ink", self._draw_ink, text_mode=True))
        groups.append(group)
        return groups

    def _initialize_review_services(self):
        self.speech_reader = SpeechReader(self)
        self._speech_active = False
        self.speech_reader.word_range.connect(self._speech_word)
        self.speech_reader.finished.connect(self._speech_finished)
        self.speech_reader.message.connect(lambda message: self.statusBar().showMessage(message, 8000))
        self.tracker.edit_blocked.connect(lambda message: self.statusBar().showMessage(message, 5000))
        self.editor.document().contentsChanged.connect(self._stop_speech_on_edit)
        self._review_state_loaded()

    def _review_state_loaded(self):
        if hasattr(self, "speech_reader"):
            self.speech_reader.stop()
        if hasattr(self, "format_tracking"):
            with QSignalBlocker(self.format_tracking):
                self.format_tracking.setChecked(self.state.review.get("track_formatting", False))
        self.tracker.refresh_protection()

    def _track_formatting(self, checked):
        if not self._can_edit():
            with QSignalBlocker(self.format_tracking):
                self.format_tracking.setChecked(self.state.review.get("track_formatting", False))
            return
        self.state.review["track_formatting"] = bool(checked)
        self.tracker.changed.emit()

    def _comment_threads(self):
        CommentThreadsDialog(self.tracker, self.editor, self, self._can_edit()).exec()

    def _review_display_selected(self, index):
        if index <= 0:
            return
        mode = self.review_display.itemText(index)
        self.review_display.setCurrentIndex(0)
        try:
            document = review_document(self.editor.document(), self.state, mode)
        except ReviewError as error:
            QMessageBox.warning(self, mode, f"This view cannot safely reconstruct overlapping edits.\n{error}")
            return
        dialog = QDialog(self)
        dialog.setWindowTitle(mode)
        dialog.resize(900, 700)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(f"{mode} · read-only snapshot. The source document is unchanged."))
        view = QTextEdit()
        view.setReadOnly(True)
        document.setParent(view)
        view.setDocument(document)
        layout.addWidget(view)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec()

    def _accessibility_check(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Accessibility checks")
        dialog.resize(760, 460)
        layout = QVBoxLayout(dialog)
        note = QLabel("Local checks cover heading order, image alternative text, table headers, link labels, "
                      "and text contrast. They use a white background when no color is set and do not "
                      "certify WCAG compliance or accessible PDF output.")
        note.setWordWrap(True)
        layout.addWidget(note)
        items = QListWidget()
        layout.addWidget(items)
        def refresh():
            items.clear()
            findings = accessibility_issues(self.editor.document())
            for issue in findings:
                item = QListWidgetItem(issue.message)
                item.setData(Qt.ItemDataRole.UserRole, issue)
                items.addItem(item)
            if not findings:
                items.addItem("No issues found by these checks; manual accessibility review is still needed.")
        def jump():
            item = items.currentItem()
            issue = item.data(Qt.ItemDataRole.UserRole) if item else None
            if issue is not None:
                cursor = self.editor.textCursor()
                cursor.setPosition(issue.start)
                cursor.setPosition(issue.start + issue.length, QTextCursor.MoveMode.KeepAnchor)
                self.editor.setTextCursor(cursor)
                self.canvas.scroll_to_cursor()
        def alternate_text():
            item = items.currentItem()
            issue = item.data(Qt.ItemDataRole.UserRole) if item else None
            if issue is None or issue.rule != "image_alt" or not self._can_edit():
                return
            value, accepted = QInputDialog.getText(dialog, "Image description", "Alternative text:")
            if accepted and value.strip():
                cursor = QTextCursor(self.editor.document())
                cursor.setPosition(issue.start)
                cursor.setPosition(issue.start + issue.length, QTextCursor.MoveMode.KeepAnchor)
                fmt = QTextCharFormat()
                fmt.setProperty(fmt.Property.ImageAltText, value.strip())
                cursor.mergeCharFormat(fmt)
                refresh()
        buttons = QHBoxLayout()
        for label, callback in (("Go to issue", jump), ("Add image description", alternate_text), ("Recheck", refresh)):
            button = QPushButton(label)
            button.clicked.connect(callback)
            buttons.addWidget(button)
        layout.addLayout(buttons)
        items.itemDoubleClicked.connect(lambda *_: jump())
        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.rejected.connect(dialog.reject)
        layout.addWidget(close)
        refresh()
        dialog.exec()

    def _read_aloud(self):
        cursor = self.editor.textCursor()
        text = cursor.selectedText().replace("\u2029", "\n") if cursor.hasSelection() else self.editor.toPlainText()
        offset = cursor.selectionStart() if cursor.hasSelection() else 0
        self._speech_active = self.speech_reader.start(text, offset)

    def _speech_word(self, start, length):
        if not self._speech_active:
            return
        self._apply_review_highlights()
        extra = QTextEdit.ExtraSelection()
        extra.cursor = QTextCursor(self.editor.document())
        bound = self.editor.document().characterCount() - 1
        extra.cursor.setPosition(min(start, bound))
        extra.cursor.setPosition(min(start + length, bound), QTextCursor.MoveMode.KeepAnchor)
        extra.format.setBackground(QColor("#ffe08a"))
        self.editor.setExtraSelections(self.editor.extraSelections() + [extra])

    def _speech_finished(self):
        self._speech_active = False
        if hasattr(self, "comments_list"):
            self._apply_review_highlights()

    def _stop_speech_on_edit(self):
        if self._speech_active:
            self.speech_reader.stop()

    def _text_output(self, title, message, text, editable=False):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(900, 660)
        layout = QVBoxLayout(dialog)
        note = QLabel(message)
        note.setWordWrap(True)
        layout.addWidget(note)
        view = QTextEdit()
        view.setAcceptRichText(False)
        view.setReadOnly(not editable)
        view.setPlainText(text)
        layout.addWidget(view)
        actions = QHBoxLayout()
        copy = QPushButton("Copy text")
        copy.clicked.connect(lambda: QApplication.clipboard().setText(view.toPlainText()))
        actions.addWidget(copy)
        save = QPushButton("Save text…")
        def save_text():
            path, _ = QFileDialog.getSaveFileName(dialog, "Save result", "", "Text files (*.txt)")
            if path:
                try:
                    Path(path).write_text(view.toPlainText(), encoding="utf-8")
                except OSError as error:
                    QMessageBox.warning(dialog, "Save result", str(error))
        save.clicked.connect(save_text)
        actions.addWidget(save)
        close = QPushButton("Close")
        close.clicked.connect(dialog.accept)
        actions.addWidget(close)
        layout.addLayout(actions)
        dialog.exec()

    def _translate_local(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import local translation glossary", "", "JSON glossaries (*.json)")
        if not path:
            return
        try:
            source, target, entries = load_glossary(path)
        except (OSError, ValueError) as error:
            QMessageBox.warning(self, "Glossary translation", str(error))
            return
        cursor = self.editor.textCursor()
        text = cursor.selectedText().replace("\u2029", "\n") if cursor.hasSelection() else self.editor.toPlainText()
        result, count = translate_glossary(text, entries)
        self._text_output("Glossary translation", f"{source} → {target} · {count} glossary matches. "
                          "This replaces supplied terms and leaves unmatched text unchanged. "
                          "It is not machine translation; no translation service is contacted.", result)

    def _three_way_compare(self):
        paths = []
        for label in ("Select common base document", "Select incoming version"):
            path, _ = QFileDialog.getOpenFileName(self, label, "", "Documents (*.docx *.txt)")
            if not path:
                return
            paths.append(Path(path))
        def read_text(path):
            if path.suffix.lower() == ".docx":
                from .docx_io import read_docx
                state, warnings = read_docx(path)
                document = SafeDocument()
                document.setHtml(state.html)
                return document.toPlainText(), warnings
            if path.stat().st_size > 16 * 1024 * 1024:
                raise ValueError("Comparison text files must be smaller than 16 MiB.")
            return path.read_text(encoding="utf-8-sig"), []
        try:
            base, first_warnings = read_text(paths[0])
            incoming, second_warnings = read_text(paths[1])
            result = compare_three_way(base, self.editor.toPlainText(), incoming)
        except (OSError, ValueError) as error:
            QMessageBox.warning(self, "Three-way compare", str(error))
            return
        message = (f"Line-based text merge: {len(result.conflicts)} conflict(s). "
                   "Independent changes are combined. Resolve conflict markers in the result before saving. "
                   "Formatting and objects are not compared; the current document is unchanged.")
        warnings = first_warnings + second_warnings
        if warnings:
            message += "\n" + "\n".join(warnings)
        self._text_output("Three-way comparison", message, result.text, editable=True)

    def _editing_restrictions(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Editing restrictions")
        dialog.resize(640, 410)
        layout = QVBoxLayout(dialog)
        note = QLabel("Local editing restrictions prevent changes inside locked ranges. They are not encryption "
                      "or access control; anyone editing this file can remove them.")
        note.setWordWrap(True)
        layout.addWidget(note)
        formatting = QCheckBox("Restrict formatting changes throughout the document")
        formatting.setChecked(self.state.review.get("formatting_locked", False))
        formatting.setEnabled(self._can_edit())
        layout.addWidget(formatting)
        items = QListWidget()
        layout.addWidget(items)
        def refresh():
            items.clear()
            for lock in self.state.review.get("locks", []):
                item = QListWidgetItem(f'{lock["name"]} · {lock["start"]}–{lock["end"]}')
                item.setData(Qt.ItemDataRole.UserRole, lock["id"])
                items.addItem(item)
        def changed():
            self.tracker.refresh_protection()
            self.tracker.changed.emit()
            refresh()
        def restrict_formatting(value):
            self.state.review["formatting_locked"] = value
            changed()
        formatting.toggled.connect(restrict_formatting)
        def add_lock(section=False):
            cursor = self.editor.textCursor()
            start, end = cursor.selectionStart(), cursor.selectionEnd()
            if section:
                positions = [0] + [position for position, _, _ in self.editor.headings()]
                start = max(position for position in positions if position <= cursor.position())
                end = next((position for position in positions if position > cursor.position()),
                           self.editor.document().characterCount() - 1)
            if start == end:
                QMessageBox.information(dialog, "Lock range", "Select text or choose a nonempty heading section.")
                return
            name, accepted = QInputDialog.getText(dialog, "Lock range", "Range name:")
            if accepted:
                self.state.review.setdefault("locks", []).append({"id": str(uuid.uuid4()),
                    "name": name.strip() or "Protected section", "start": start, "end": end})
                changed()
        def remove():
            item = items.currentItem()
            if item:
                identifier = item.data(Qt.ItemDataRole.UserRole)
                self.state.review["locks"] = [lock for lock in self.state.review.get("locks", [])
                                             if lock["id"] != identifier]
                changed()
        actions = QHBoxLayout()
        for label, slot in (("Lock selection", lambda: add_lock()),
                            ("Lock heading section", lambda: add_lock(True)), ("Remove lock", remove)):
            button = QPushButton(label)
            button.setEnabled(self._can_edit())
            button.clicked.connect(slot)
            actions.addWidget(button)
        layout.addLayout(actions)
        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.rejected.connect(dialog.reject)
        layout.addWidget(close)
        refresh()
        dialog.exec()

    def _draw_ink(self):
        if not self._can_edit():
            return
        from .document_features import insert_node

        selected = self._selected_node("ink")
        dialog = InkDialog(self, selected[1]["ink"] if selected else None)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.canvas.strokes:
            if selected:
                self._select_node_range(selected[0])
            cursor = self.editor.textCursor()
            if not self.tracker.allows_edit(cursor.selectionStart(), cursor.selectionEnd()):
                self.statusBar().showMessage("This drawing is inside a protected range.", 5000)
                return
            try:
                insert_node(self.editor, self.state, {"kind": "ink", "ink": dialog.canvas.ink()},
                            image=dialog.canvas.image())
            except ValueError as error:
                QMessageBox.warning(self, "Digital ink", str(error))
