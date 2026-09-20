from __future__ import annotations

import copy

from PySide6.QtCore import QEvent, QObject, QSignalBlocker, Qt, QTimer
from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QDockWidget,
                              QFormLayout, QHBoxLayout, QInputDialog, QLabel,
                              QLineEdit, QListWidget, QListWidgetItem, QMenu,
                              QMessageBox, QPushButton, QTextEdit, QVBoxLayout, QWidget)

from . import document_features as features
from . import references
from .feature_commands import editable
from .publishing import Publication


class SourceDialog(QDialog):
    def __init__(self, parent=None, source=None):
        super().__init__(parent)
        self.setWindowTitle("Citation source")
        self.setMinimumWidth(480)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.source_type = QComboBox()
        self.source_type.addItems(["book", "article", "website"])
        self.source_type.setCurrentText((source or {}).get("type", "book"))
        form.addRow("Type", self.source_type)
        self.fields = {}
        for key, title in (("author", "Authors (Last, First; Last, First)"), ("title", "Title"),
                           ("year", "Year"), ("publisher", "Publisher"), ("journal", "Journal"),
                           ("volume", "Volume"), ("issue", "Issue"), ("pages", "Pages"),
                           ("doi", "DOI"), ("url", "URL")):
            field = QLineEdit((source or {}).get(key, ""))
            field.setMaxLength(2000)
            self.fields[key] = field
            form.addRow(title, field)
        layout.addLayout(form)
        note = QLabel("Common book, article, and website formats. Check publisher-specific style requirements.")
        note.setWordWrap(True)
        layout.addWidget(note)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self._accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def source(self):
        return references.checked_source({"type": self.source_type.currentText(),
                                          **{key: value.text() for key, value in self.fields.items()}})

    def _accept(self):
        try:
            self.source()
        except ValueError as exc:
            QMessageBox.warning(self, "Citation source", str(exc))
            return
        self.accept()


class ReferenceLinkFilter(QObject):
    def __init__(self, window):
        super().__init__(window)
        self.window = window

    def eventFilter(self, watched, event):
        if event.type() == QEvent.Type.MouseButtonRelease and event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            link = self.window.editor.anchorAt(event.position().toPoint())
            if link.startswith("folio-note:"):
                self.window._show_note(link[len("folio-note:"):])
                return True
            if link.startswith("#" + features.ANCHOR):
                self.window._jump_reference(link[len(features.ANCHOR) + 1:])
                return True
        return False


class ReferenceCommands:
    def _initialize_references(self):
        self._references_updating = False
        self._references_timer = QTimer(self)
        self._references_timer.setSingleShot(True)
        self._references_timer.setInterval(550)
        self._references_timer.timeout.connect(self._refresh_references)
        self.editor.document().contentsChanged.connect(self._schedule_references)
        self._reference_links = ReferenceLinkFilter(self)
        self.editor.viewport().installEventFilter(self._reference_links)
        self.notes_dock = QDockWidget("Footnotes & Endnotes", self)
        container = QWidget()
        layout = QVBoxLayout(container)
        self.notes_list = QListWidget()
        self.notes_list.currentItemChanged.connect(self._note_selected)
        layout.addWidget(self.notes_list)
        self.note_text = QTextEdit()
        self.note_text.setPlaceholderText("Select a note to read or edit it.")
        layout.addWidget(self.note_text)
        row = QHBoxLayout()
        for label, slot in (("Save Note", self._save_note), ("Go to Anchor", self._goto_note_anchor),
                            ("Next Note", self._next_note)):
            button = QPushButton(label)
            button.clicked.connect(slot)
            row.addWidget(button)
        layout.addLayout(row)
        self.notes_dock.setWidget(container)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.notes_dock)
        self.notes_dock.hide()

    def _build_references_tab(self):
        groups = []
        g, layout = self._group("Contents")
        for level in (1, 2, 3):
            layout.addWidget(self._tool("heading", f"H{level}",
                                       lambda _c=False, n=level: self.editor.apply_style(f"Heading {n}"), text_mode=True))
        layout.addWidget(self._tool("toc", "TOC", self._update_toc, text_mode=True))
        layout.addWidget(self._tool("toc-refresh", "Update All", self._refresh_all_references, text_mode=True))
        groups.append(g)
        g, layout = self._group("Notes")
        menu = QMenu(self)
        menu.addAction("Insert Footnote", lambda: self._insert_reference_note("footnote"))
        menu.addAction("Insert Endnote", lambda: self._insert_reference_note("endnote"))
        menu.addAction("Show Notes", self._show_notes)
        menu.addAction("Next Note", self._next_note)
        layout.addWidget(self._menu_tool("Notes", menu))
        groups.append(g)
        g, layout = self._group("Citations & Bibliography")
        self.citation_style_combo = QComboBox()
        self.citation_style_combo.addItems(references.STYLES)
        self.citation_style_combo.setCurrentText(references.registry(self.state).get("style", "APA"))
        self.citation_style_combo.currentTextChanged.connect(self._citation_style_changed)
        layout.addWidget(self.citation_style_combo)
        menu = QMenu(self)
        menu.addAction("Insert Citation", self._insert_citation)
        menu.addAction("Manage Sources", self._manage_sources)
        menu.addAction("Insert Bibliography", lambda: self._insert_reference_collection("bibliography"))
        layout.addWidget(self._menu_tool("Citations", menu))
        groups.append(g)
        g, layout = self._group("Captions & Cross-references")
        menu = QMenu(self)
        menu.addAction("Insert / Edit Caption", self._insert_caption)
        menu.addAction("Table of Figures", lambda: self._insert_reference_collection("figures"))
        menu.addAction("Insert Cross-reference", self._cross_reference)
        menu.addAction("Go to Reference", self._follow_reference)
        layout.addWidget(self._menu_tool("Captions", menu))
        groups.append(g)
        g, layout = self._group("Index & Authorities")
        menu = QMenu(self)
        menu.addAction("Mark Index Entry", self._mark_index)
        menu.addAction("Insert Index", lambda: self._insert_reference_collection("index"))
        menu.addAction("Mark Legal Authority", self._mark_authority)
        menu.addAction("Table of Authorities", lambda: self._insert_reference_collection("authorities"))
        layout.addWidget(self._menu_tool("Mark & Compile", menu))
        groups.append(g)
        self.ribbon.addTab(self._wrap(groups), "References")

    def _references_loaded(self):
        if hasattr(self, "_references_timer"):
            self._references_timer.stop()
            self._refresh_notes_list()
        if hasattr(self, "citation_style_combo"):
            with QSignalBlocker(self.citation_style_combo):
                self.citation_style_combo.setCurrentText(references.registry(self.state).get("style", "APA"))

    def _schedule_references(self):
        if not self._references_updating and self._can_edit():
            self._references_timer.start()

    def _refresh_references(self):
        if self._references_updating or not self._can_edit():
            return
        self._references_timer.stop()
        self._references_updating = True
        try:
            doc = self.editor.document()
            dynamic = any(node["kind"] in references.REFERENCE_KINDS or node["kind"] == "reference"
                          for node in self.state.features["nodes"].values())
            if dynamic:
                with self.tracker.applying():
                    can_edit = lambda start, end: self.tracker.allows_edit(start, end, formatting=True)
                    for _ in range(5):
                        publication = Publication(doc, self.state.page, self._title)
                        count = references.refresh_references(doc, self.state, publication.page_for_position, can_edit)
                        count += features.refresh_fields(doc, self.state.features,
                                                         page_for_position=publication.page_for_position, can_edit=can_edit)
                        if not count:
                            break
            if self.state.toc:
                self._refresh_toc_quietly()
            self._refresh_notes_list()
        finally:
            self._references_updating = False

    @editable
    def _refresh_all_references(self):
        self._refresh_references()
        self.statusBar().showMessage("References updated", 3000)

    @editable
    def _citation_style_changed(self, style):
        references.registry(self.state)["style"] = style
        self._refresh_references()
        self._feature_changed()

    @editable
    def _insert_reference_note(self, note_type):
        if not self._reference_selection_editable():
            return
        text, accepted = QInputDialog.getMultiLineText(self, "Insert " + note_type, "Note text:")
        if not accepted or not text.strip():
            return
        try:
            key = references.insert_note(self.editor, self.state, text, note_type)
        except ValueError as exc:
            QMessageBox.warning(self, "Note", str(exc))
            return
        self._refresh_references()
        self._feature_changed()
        self._show_note(key)

    def _refresh_notes_list(self):
        item = self.notes_list.currentItem()
        selected = item.data(Qt.ItemDataRole.UserRole) if item else None
        with QSignalBlocker(self.notes_list):
            self.notes_list.clear()
            for _, key, number, node in references.note_entries(self.editor.document(), self.state):
                item = QListWidgetItem(f"{node['type'].title()} {number}: {node['text'][:70]}")
                item.setData(Qt.ItemDataRole.UserRole, key)
                self.notes_list.addItem(item)
                if key == selected:
                    self.notes_list.setCurrentItem(item)
        if selected and selected not in features.node_ranges(self.editor.document()):
            self.note_text.clear()

    def _note_selected(self, item, previous=None):
        key = item.data(Qt.ItemDataRole.UserRole) if item else None
        node = self.state.features["nodes"].get(key, {})
        self.note_text.setPlainText(node.get("text", ""))
        self.note_text.setReadOnly(not self._can_edit())

    def _show_notes(self):
        self._refresh_notes_list()
        self.notes_dock.show()
        if self.notes_list.currentRow() < 0 and self.notes_list.count():
            self.notes_list.setCurrentRow(0)

    def _show_note(self, key):
        self._show_notes()
        for row in range(self.notes_list.count()):
            if self.notes_list.item(row).data(Qt.ItemDataRole.UserRole) == key:
                self.notes_list.setCurrentRow(row)
                self.note_text.setFocus()
                break

    def _next_note(self):
        self._show_notes()
        if self.notes_list.count():
            self.notes_list.setCurrentRow((self.notes_list.currentRow() + 1) % self.notes_list.count())
            self._goto_note_anchor()

    @editable
    def _save_note(self):
        item = self.notes_list.currentItem()
        if item:
            key = item.data(Qt.ItemDataRole.UserRole)
            start, end = features.node_ranges(self.editor.document()).get(key, (0, 0))
            if not self.tracker.allows_edit(start, end):
                self.statusBar().showMessage("This note anchor is in a protected range.", 5000)
                return
            text = self.note_text.toPlainText()
            if len(text) > 10000:
                QMessageBox.warning(self, "Note", "Notes are limited to 10,000 characters.")
                return
            self.state.features["nodes"][key]["text"] = text
            self._refresh_references()
            self._feature_changed()

    def _goto_note_anchor(self):
        item = self.notes_list.currentItem()
        if item:
            self._jump_reference(item.data(Qt.ItemDataRole.UserRole))

    def _jump_reference(self, key):
        if key in features.node_ranges(self.editor.document()):
            self._select_node_range(key)
            self.editor.ensureCursorVisible()
            self.editor.setFocus()

    def _follow_reference(self):
        selected = self._selected_node("reference")
        if selected:
            self._jump_reference(selected[1]["target"])
        else:
            selected = self._selected_node("note")
            if selected:
                self._show_note(selected[0])

    @editable
    def _manage_sources(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Manage Citation Sources")
        dialog.resize(610, 390)
        layout = QVBoxLayout(dialog)
        scope = QComboBox()
        scope.addItems(["This document", "Global library (this computer)"])
        layout.addWidget(scope)
        listing = QListWidget()
        layout.addWidget(listing)
        global_sources = self.store.get_setting("reference_sources", {})
        if not isinstance(global_sources, dict):
            global_sources = {}
        global_sources = {key: value for key, value in global_sources.items()
                          if isinstance(key, str) and isinstance(value, dict) and value.get("title")}

        def sources():
            return references.registry(self.state).setdefault("sources", {}) if scope.currentIndex() == 0 else global_sources

        def refresh():
            listing.clear()
            for key, source in sorted(sources().items(), key=lambda item: item[1].get("title", "").casefold()):
                item = QListWidgetItem(f"{source.get('author', '')} — {source['title']} ({source.get('year', '')})")
                item.setData(Qt.ItemDataRole.UserRole, key)
                listing.addItem(item)

        def persist():
            self.store.set_setting("reference_sources", global_sources)
            self._feature_changed()
            self._refresh_references()
            refresh()

        def edit(new=False):
            item = listing.currentItem()
            if not new and not item:
                return
            key = features.uid() if new else item.data(Qt.ItemDataRole.UserRole)
            entry = SourceDialog(dialog, None if new else sources()[key])
            if entry.exec() == QDialog.DialogCode.Accepted:
                sources()[key] = entry.source()
                persist()

        def remove():
            item = listing.currentItem()
            if not item:
                return
            key = item.data(Qt.ItemDataRole.UserRole)
            if scope.currentIndex() == 0 and any(node["kind"] == "citation" and node["source"] == key
                                                  for _, _, node in references.ordered_nodes(self.editor.document(), self.state)):
                QMessageBox.information(dialog, "Source in use", "Remove this source's citations before removing it from the document.")
                return
            sources().pop(key)
            persist()

        def transfer():
            item = listing.currentItem()
            if item:
                key = item.data(Qt.ItemDataRole.UserRole)
                target = global_sources if scope.currentIndex() == 0 else references.registry(self.state).setdefault("sources", {})
                target[key] = copy.deepcopy(sources()[key])
                persist()

        row = QHBoxLayout()
        for label, slot in (("Add", lambda: edit(True)), ("Edit", edit), ("Remove", remove),
                            ("Copy to Other Library", transfer)):
            button = QPushButton(label)
            button.clicked.connect(slot)
            row.addWidget(button)
        layout.addLayout(row)
        scope.currentIndexChanged.connect(refresh)
        refresh()
        dialog.exec()

    @editable
    def _insert_citation(self):
        if not self._reference_selection_editable():
            return
        sources = references.registry(self.state).setdefault("sources", {})
        if not sources:
            self._manage_sources()
        if not sources:
            return
        keys = list(sources)
        labels = [f"{index + 1}. {sources[key].get('author', '')} — {sources[key]['title']}" for index, key in enumerate(keys)]
        selected, accepted = QInputDialog.getItem(self, "Insert Citation", "Source:", labels, editable=False)
        if not accepted:
            return
        locator, accepted = QInputDialog.getText(self, "Citation Locator", "Page or page range (optional):")
        if accepted:
            references.insert_citation(self.editor, self.state, keys[labels.index(selected)], locator)
            self._refresh_references()
            self._feature_changed()

    @editable
    def _insert_caption(self):
        if not self._reference_selection_editable():
            return
        selected = self._selected_node("caption")
        node = selected[1] if selected else {}
        label, accepted = QInputDialog.getItem(self, "Caption", "Label:", references.LABELS,
                                             references.LABELS.index(node.get("label", "Figure")), False)
        if not accepted:
            return
        text, accepted = QInputDialog.getText(self, "Caption", "Description:", text=node.get("text", ""))
        if accepted:
            if selected:
                node.update(label=label, text=text)
            else:
                cursor = self.editor.textCursor()
                if cursor.block().text():
                    cursor.movePosition(QTextCursor.MoveOperation.EndOfBlock)
                    cursor.insertBlock()
                    self.editor.setTextCursor(cursor)
                features.insert_node(self.editor, self.state, {"kind": "caption", "label": label, "text": text}, "Caption")
                cursor = self.editor.textCursor()
                cursor.insertBlock()
                self.editor.setTextCursor(cursor)
            self._refresh_references()
            self._feature_changed()

    @editable
    def _insert_reference_collection(self, collection):
        if not self._reference_selection_editable():
            return
        label = "All"
        if collection == "figures":
            label, accepted = QInputDialog.getItem(self, "Table of Figures", "Include:", ["All", *references.LABELS], editable=False)
            if not accepted:
                return
        references.ensure_collection(self.editor, self.state, collection, label)
        self._refresh_references()
        self._feature_changed()

    @editable
    def _mark_index(self):
        self._mark_reference_term(False)

    @editable
    def _mark_authority(self):
        self._mark_reference_term(True)

    def _mark_reference_term(self, authority):
        if not self._reference_selection_editable():
            return
        cursor = self.editor.textCursor()
        if not cursor.hasSelection():
            QMessageBox.information(self, "Mark Entry", "Select the text to mark first.")
            return
        term, accepted = QInputDialog.getText(self, "Mark Authority" if authority else "Mark Index Entry",
                                            "Entry:", text=cursor.selectedText())
        if not accepted:
            return
        category = None
        if authority:
            category, accepted = QInputDialog.getItem(self, "Authority Category", "Category:",
                                                     ["Cases", "Statutes", "Regulations", "Treaties", "Other authorities"], editable=True)
            if not accepted:
                return
        try:
            references.tag_selection(self.editor, self.state, term, category)
        except ValueError as exc:
            QMessageBox.warning(self, "Mark Entry", str(exc))
            return
        self._refresh_references()
        self._feature_changed()

    @editable
    def _cross_reference(self):
        if not self._reference_selection_editable():
            return
        choices = []
        for position, key, node in references.ordered_nodes(self.editor.document(), self.state):
            if node["kind"] in ("bookmark", "caption", "note"):
                label = node.get("name", node.get("text", key))
                choices.append((f"{node['kind'].title()}: {label[:100]}", key))
        for position, level, text in self.editor.headings():
            choices.append((f"Heading {level}: {text}", (position, self.editor.document().findBlock(position).length() - 1)))
        if not choices:
            QMessageBox.information(self, "Cross-reference", "Add a heading, bookmark, caption, or note first.")
            return
        labels = [f"{index + 1}. {label}" for index, (label, _) in enumerate(choices)]
        choice, accepted = QInputDialog.getItem(self, "Cross-reference", "Target:", labels, editable=False)
        if not accepted:
            return
        display, accepted = QInputDialog.getItem(self, "Cross-reference", "Display:", ["Text / note number", "Page number"], editable=False)
        if not accepted:
            return
        target = choices[labels.index(choice)][1]
        if isinstance(target, tuple):
            position, length = target
            cursor = QTextCursor(self.editor.document())
            cursor.setPosition(position)
            cursor.setPosition(position + length, QTextCursor.MoveMode.KeepAnchor)
            target = features.uid()
            features.mark_cursor(cursor, target)
            self.state.features["nodes"][target] = {"kind": "bookmark", "name": "Heading_" + target}
        features.insert_node(self.editor, self.state, {"kind": "reference", "target": target,
                                                     "display": "page" if display == "Page number" else "text"}, "Reference")
        self._refresh_references()
        self._feature_changed()

    def _reference_selection_editable(self):
        cursor = self.editor.textCursor()
        if not self.tracker.allows_edit(cursor.selectionStart(), cursor.selectionEnd(), formatting=True):
            self.statusBar().showMessage("This range is protected, or formatting is restricted.", 5000)
            return False
        return True
