from __future__ import annotations

from PySide6.QtCore import QRectF, Qt, QTimer
from PySide6.QtGui import QColor, QKeySequence, QPainter, QTextCursor, QTextDocumentFragment
from PySide6.QtWidgets import (QComboBox, QGraphicsItem, QGraphicsScene,
                              QGraphicsView, QHBoxLayout, QLabel, QPushButton,
                              QMenu, QSpinBox, QStackedWidget, QTextEdit, QTreeWidget,
                              QTreeWidgetItem, QVBoxLayout, QWidget)

from .editor import SafeDocument
from .publishing import Publication


class LiveTextView(QTextEdit):
    def __init__(self, owner, draft=False, parent=None):
        super().__init__(parent)
        self.owner = owner
        self.draft = draft
        self._syncing = False
        self.setDocument(SafeDocument(self))
        self.setAcceptRichText(not draft)
        self.pull()
        self.document().contentsChange.connect(self.push)
        owner.editor.document().contentsChanged.connect(self.pull)

    def pull(self):
        if self._syncing:
            return
        self._syncing = True
        try:
            position = self.textCursor().position()
            scroll = self.verticalScrollBar().value()
            source = self.owner.editor.document()
            if self.draft:
                self.setPlainText(source.toPlainText())
            else:
                self.setHtml(source.toHtml())
            self.document().folio_state = self.owner.state
            self.document().setDefaultFont(source.defaultFont())
            cursor = self.textCursor()
            cursor.setPosition(min(position, self.document().characterCount() - 1))
            self.setTextCursor(cursor)
            self.verticalScrollBar().setValue(scroll)
            self.setReadOnly(self.owner.editor.isReadOnly())
        finally:
            self._syncing = False

    def push(self, position, removed, added):
        if self._syncing:
            return
        if self.owner.editor.isReadOnly() or not self.owner._can_edit():
            self.pull()
            return
        source = self.owner.editor.document()
        if position > source.characterCount() - 1:
            self.pull()
            return
        formatting = removed == added and source.toPlainText() == self.toPlainText()
        if not self.owner.tracker.allows_edit(position, min(position + removed, source.characterCount() - 1),
                                              formatting=formatting):
            self.pull()
            return
        self._syncing = True
        try:
            target = QTextCursor(source)
            target.setPosition(position)
            target.setPosition(min(position + removed, source.characterCount() - 1),
                               QTextCursor.MoveMode.KeepAnchor)
            fragment = QTextCursor(self.document())
            fragment.setPosition(position)
            fragment.setPosition(min(position + added, self.document().characterCount() - 1),
                                 QTextCursor.MoveMode.KeepAnchor)
            target.beginEditBlock()
            if added == 0:
                target.removeSelectedText()
            elif self.draft:
                target.insertText(fragment.selectedText().replace("\u2029", "\n"))
            else:
                target.insertFragment(QTextDocumentFragment(fragment))
            target.endEditBlock()
            self.owner.editor.setTextCursor(target)
        finally:
            self._syncing = False
        if source.toPlainText() != self.toPlainText():
            self.pull()

    def keyPressEvent(self, event):
        from .document_features import detach_completed_node

        detach_completed_node(self)
        if event.matches(QKeySequence.StandardKey.Undo):
            self.undo()
        elif event.matches(QKeySequence.StandardKey.Redo):
            self.redo()
        else:
            super().keyPressEvent(event)

    def insertPlainText(self, text):
        from .document_features import detach_completed_node

        detach_completed_node(self)
        super().insertPlainText(text)

    def inputMethodEvent(self, event):
        from .document_features import detach_completed_node

        detach_completed_node(self)
        super().inputMethodEvent(event)

    def insertFromMimeData(self, source):
        if self.owner.editor.isReadOnly() or not self.owner._can_edit():
            return
        selection = self.textCursor()
        cursor = QTextCursor(self.owner.editor.document())
        cursor.setPosition(selection.anchor())
        cursor.setPosition(selection.position(), QTextCursor.MoveMode.KeepAnchor)
        self.owner.editor.setTextCursor(cursor)
        self.owner.editor.insert_mime_data(source, "text" if self.draft else "keep_formatting")
        position = self.owner.editor.textCursor().position()
        self.pull()
        cursor = self.textCursor()
        cursor.setPosition(min(position, self.document().characterCount() - 1))
        self.setTextCursor(cursor)

    def createMimeDataFromSelection(self):
        if self.draft:
            return super().createMimeDataFromSelection()
        saved = self.owner.editor.textCursor()
        selection = self.textCursor()
        cursor = QTextCursor(self.owner.editor.document())
        cursor.setPosition(selection.anchor())
        cursor.setPosition(selection.position(), QTextCursor.MoveMode.KeepAnchor)
        self.owner.editor.setTextCursor(cursor)
        try:
            return self.owner.editor.createMimeDataFromSelection()
        finally:
            self.owner.editor.setTextCursor(saved)

    def undo(self):
        if not self.owner.editor.isReadOnly() and self.owner._can_edit():
            self.owner.editor.undo()

    def redo(self):
        if not self.owner.editor.isReadOnly() and self.owner._can_edit():
            self.owner.editor.redo()

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        editable = not self.owner.editor.isReadOnly() and self.owner._can_edit()
        for label, callback, enabled in (
                ("Undo", self.undo, editable and self.owner.editor.document().isUndoAvailable()),
                ("Redo", self.redo, editable and self.owner.editor.document().isRedoAvailable()),
                ("Cut", self.cut, editable and self.textCursor().hasSelection()),
                ("Copy", self.copy, self.textCursor().hasSelection()),
                ("Paste", self.paste, editable), ("Select All", self.selectAll, True)):
            menu.addAction(label, callback).setEnabled(enabled)
        menu.exec(event.globalPos())


class PublicationPage(QGraphicsItem):
    def __init__(self, publication, index):
        super().__init__()
        self.publication = publication
        self.index = index
        width, height = publication.page_size(index)
        self.bounds = QRectF(0, 0, width, height)

    def boundingRect(self):
        return self.bounds

    def paint(self, painter, option, widget=None):
        painter.fillRect(self.bounds, QColor("white"))
        self.publication.paint_page(painter, self.index)


class PageView(QGraphicsView):
    def __init__(self, owner, parent=None):
        super().__init__(parent)
        self.owner = owner
        self.horizontal = False
        self.zoom_mode = "Page Width"
        self.percent = 100
        self.setScene(QGraphicsScene(self))
        self.setBackgroundBrush(QColor("#e7ebf0"))
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)
        self.publication = None
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(150)
        self.timer.timeout.connect(self.refresh)
        owner.editor.document().contentsChanged.connect(self.timer.start)
        self.refresh()

    def refresh(self):
        self.publication = Publication(self.owner.editor.document(), self.owner.state.page,
                                       self.owner._title, state=self.owner.state)
        self.scene().clear()
        x, y = 24.0, 24.0
        for index in range(self.publication.page_count()):
            item = PublicationPage(self.publication, index)
            item.setPos(x, y)
            self.scene().addItem(item)
            if self.horizontal:
                x += item.bounds.width() + 24
            else:
                y += item.bounds.height() + 24
        self.scene().setSceneRect(self.scene().itemsBoundingRect().adjusted(-24, -24, 24, 24))
        self.apply_zoom()

    def apply_zoom(self):
        if self.publication is None:
            return
        width, height = self.publication.page_size(0)
        available_w = max(1, self.viewport().width() - 48)
        available_h = max(1, self.viewport().height() - 48)
        if self.zoom_mode == "Percent":
            factor = self.percent / 100
        elif self.zoom_mode == "Single Page":
            factor = min(available_w / width, available_h / height)
        elif self.zoom_mode == "Multiple Pages":
            factor = min(available_w / (2 * width + 24), available_h / height)
        else:
            factor = available_w / width
        factor = max(0.1, min(5.0, factor))
        self.resetTransform()
        self.scale(factor, factor)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.apply_zoom()

    def mouseDoubleClickEvent(self, event):
        item = self.itemAt(event.position().toPoint())
        if isinstance(item, PublicationPage):
            position = self.publication.first_position_on_page(item.index)
            cursor = self.owner.editor.textCursor()
            cursor.setPosition(min(position, self.owner.editor.document().characterCount() - 1))
            self.owner.editor.setTextCursor(cursor)
            self.owner.raise_()
            self.owner.editor.setFocus()
        super().mouseDoubleClickEvent(event)


class DocumentViewport(QWidget):
    def __init__(self, owner, mode="Print Layout", parent=None):
        super().__init__(parent)
        self.owner = owner
        self.views = {}
        self.setWindowTitle(f"{owner._title} — document view")
        self.resize(1000, 750)
        layout = QVBoxLayout(self)
        controls = QHBoxLayout()
        self.mode = QComboBox()
        self.mode.addItems(["Print Layout", "Web Layout", "Read Mode", "Outline Mode", "Draft Mode"])
        controls.addWidget(self.mode)
        self.zoom = QComboBox()
        self.zoom.addItems(["Page Width", "Single Page", "Multiple Pages", "Percent"])
        controls.addWidget(self.zoom)
        self.percent = QSpinBox()
        self.percent.setRange(10, 500)
        self.percent.setValue(100)
        self.percent.setSuffix("%")
        controls.addWidget(self.percent)
        self.sync = QPushButton("Sync scrolling")
        self.sync.setCheckable(True)
        controls.addWidget(self.sync)
        controls.addStretch()
        layout.addLayout(controls)
        self.hint = QLabel()
        layout.addWidget(self.hint)
        self.stack = QStackedWidget()
        layout.addWidget(self.stack)
        self.pages = PageView(owner)
        self.stack.addWidget(self.pages)
        self.outline = QTreeWidget()
        self.outline.setHeaderLabels(["Document outline"])
        self.outline.itemActivated.connect(self.jump)
        self.stack.addWidget(self.outline)
        self.mode.currentTextChanged.connect(self.set_mode)
        self.zoom.currentTextChanged.connect(self.set_zoom)
        self.percent.valueChanged.connect(self.set_zoom)
        owner.editor.document().contentsChanged.connect(self.refresh_outline)
        owner.canvas.verticalScrollBar().valueChanged.connect(self.scroll_from_owner)
        self.pages.verticalScrollBar().valueChanged.connect(self.scroll_to_owner)
        self.mode.setCurrentText(mode)
        self.set_mode(mode)

    def set_mode(self, mode):
        self.zoom.setVisible(mode in ("Print Layout", "Read Mode"))
        self.percent.setVisible(mode in ("Print Layout", "Read Mode"))
        if mode in ("Print Layout", "Read Mode"):
            if mode == "Read Mode" and self.zoom.currentText() == "Page Width":
                self.zoom.setCurrentText("Single Page")
            self.pages.horizontal = mode == "Read Mode" or self.zoom.currentText() == "Multiple Pages"
            self.pages.refresh()
            self.stack.setCurrentWidget(self.pages)
            self.hint.setText("Double-click a page to edit it in the main window." if mode == "Print Layout"
                              else "Scroll horizontally to read the document.")
        elif mode == "Outline Mode":
            self.refresh_outline()
            self.stack.setCurrentWidget(self.outline)
            self.hint.setText("Double-click a heading to navigate to its text.")
        else:
            if mode not in self.views:
                view = LiveTextView(self.owner, draft=mode == "Draft Mode")
                self.views[mode] = view
                self.stack.addWidget(view)
                view.verticalScrollBar().valueChanged.connect(self.scroll_to_owner)
            self.views[mode].pull()
            self.stack.setCurrentWidget(self.views[mode])
            self.hint.setText("Edits update the active document. Undo uses its shared history.")

    def set_zoom(self, *_args):
        self.pages.zoom_mode = self.zoom.currentText()
        self.pages.percent = self.percent.value()
        horizontal = self.mode.currentText() == "Read Mode" or self.zoom.currentText() == "Multiple Pages"
        if self.pages.horizontal != horizontal:
            self.pages.horizontal = horizontal
            self.pages.refresh()
        else:
            self.pages.apply_zoom()

    def refresh_outline(self):
        self.outline.clear()
        parents = {}
        for position, level, text in self.owner.editor.headings():
            parent = next((parents[key] for key in sorted(parents, reverse=True) if key < level), self.outline)
            item = QTreeWidgetItem(parent, [text or "(empty heading)"])
            item.setData(0, Qt.ItemDataRole.UserRole, position)
            parents = {key: value for key, value in parents.items() if key < level}
            parents[level] = item
        self.outline.expandAll()

    def jump(self, item, _column):
        cursor = self.owner.editor.textCursor()
        cursor.setPosition(item.data(0, Qt.ItemDataRole.UserRole))
        self.owner.editor.setTextCursor(cursor)
        self.owner.editor.setFocus()

    def active_scrollbar(self):
        return self.stack.currentWidget().verticalScrollBar()

    def scroll_from_owner(self, value):
        if not self.sync.isChecked() or getattr(self, "_scrolling", False):
            return
        self._scrolling = True
        try:
            source = self.owner.canvas.verticalScrollBar()
            target = self.active_scrollbar()
            target.setValue(round(value / max(1, source.maximum()) * target.maximum()))
        finally:
            self._scrolling = False

    def scroll_to_owner(self, value):
        if not self.sync.isChecked() or getattr(self, "_scrolling", False):
            return
        self._scrolling = True
        try:
            source = self.active_scrollbar()
            target = self.owner.canvas.verticalScrollBar()
            target.setValue(round(value / max(1, source.maximum()) * target.maximum()))
        finally:
            self._scrolling = False
