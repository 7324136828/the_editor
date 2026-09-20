from __future__ import annotations

import html

from PySide6.QtCore import QPointF, Qt, QSize
from PySide6.QtGui import QColor, QImage, QPainter, QPen, QTextCursor
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QHBoxLayout,
    QInputDialog, QLabel, QListWidget, QListWidgetItem, QMessageBox,
    QPushButton, QSpinBox, QTextBrowser, QVBoxLayout, QWidget,
)


class CommentThreadsDialog(QDialog):
    def __init__(self, tracker, editor, parent=None, editable: bool = True):
        super().__init__(parent)
        self.setWindowTitle("Comment threads")
        self.resize(860, 620)
        self.tracker = tracker
        self.editor = editor
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Local reviewer names identify authors; no accounts or network are used."))
        filters = QHBoxLayout()
        self.author = QComboBox()
        self.author.addItem("All authors", None)
        authors = {item.author for item in tracker.state.comments}
        authors.update(reply.author for item in tracker.state.comments for reply in item.replies)
        for author in sorted(authors):
            self.author.addItem(author, author)
        self.resolved = QCheckBox("Include resolved")
        self.resolved.setChecked(True)
        filters.addWidget(self.author)
        filters.addWidget(self.resolved)
        layout.addLayout(filters)
        self.items = QListWidget()
        layout.addWidget(self.items, 1)
        self.thread = QTextBrowser()
        self.thread.setOpenExternalLinks(False)
        layout.addWidget(self.thread, 1)
        buttons = QHBoxLayout()
        for label, action, mutates in (
                ("Previous", lambda: self.navigate(-1), False),
                ("Next", lambda: self.navigate(1), False),
                ("Go to anchor", self.jump, False),
                ("Reply", self.reply, True),
                ("Resolve / Reopen", self.resolve, True),
                ("Delete thread", self.delete, True)):
            button = QPushButton(label)
            button.clicked.connect(action)
            button.setEnabled(editable or not mutates)
            buttons.addWidget(button)
        layout.addLayout(buttons)
        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.rejected.connect(self.reject)
        layout.addWidget(close)
        self.author.currentIndexChanged.connect(self.refresh)
        self.resolved.toggled.connect(self.refresh)
        self.items.currentItemChanged.connect(self.show_thread)
        self.items.itemDoubleClicked.connect(lambda *_: self.jump())
        self.tracker.changed.connect(self.refresh)
        self.finished.connect(lambda *_: self.tracker.changed.disconnect(self.refresh))
        self.refresh()

    def current(self):
        item = self.items.currentItem()
        identifier = item.data(Qt.ItemDataRole.UserRole) if item else None
        return next((comment for comment in self.tracker.state.comments if comment.id == identifier), None)

    def refresh(self, *_):
        current = self.current()
        selected = current.id if current else None
        self.items.clear()
        for comment in self.tracker.chronological_comments(
                self.author.currentData(), self.resolved.isChecked()):
            prefix = "[Resolved] " if comment.resolved else ""
            item = QListWidgetItem(f"{prefix}{comment.created[:16]} · {comment.author}: {comment.text}")
            item.setData(Qt.ItemDataRole.UserRole, comment.id)
            self.items.addItem(item)
            if comment.id == selected:
                self.items.setCurrentItem(item)
        if self.items.currentRow() < 0 and self.items.count():
            self.items.setCurrentRow(0)
        self.show_thread()

    def show_thread(self, *_):
        comment = self.current()
        if comment is None:
            self.thread.clear()
            return
        parts = [f"<p><b>Anchor:</b> {html.escape(comment.quote)}</p>"]
        if comment.orphaned:
            parts.append("<p>The original anchor was edited or deleted.</p>")
        for entry in [comment] + sorted(comment.replies, key=lambda value: value.created):
            parts.append(f"<p><b>{html.escape(entry.author)}</b> · {html.escape(entry.created[:19])}<br>"
                         + html.escape(entry.text).replace("\n", "<br>") + "</p>")
        self.thread.setHtml("".join(parts))

    def navigate(self, step):
        if self.items.count():
            self.items.setCurrentRow((self.items.currentRow() + step) % self.items.count())
            self.jump()

    def jump(self):
        comment = self.current()
        if comment is None:
            return
        cursor = self.editor.textCursor()
        bound = self.editor.document().characterCount() - 1
        cursor.setPosition(min(comment.start, bound))
        cursor.setPosition(min(comment.end, bound), QTextCursor.MoveMode.KeepAnchor)
        self.editor.setTextCursor(cursor)
        if hasattr(self.parent(), "canvas"):
            self.parent().canvas.scroll_to_cursor()

    def reply(self):
        comment = self.current()
        if comment is None:
            return
        text, accepted = QInputDialog.getMultiLineText(self, "Reply to comment", "Reply:")
        if accepted and text.strip():
            self.tracker.reply_to_comment(comment.id, text)

    def resolve(self):
        comment = self.current()
        if comment is not None:
            self.tracker.resolve_comment(comment.id, not comment.resolved)

    def delete(self):
        comment = self.current()
        if comment is not None:
            self.tracker.delete_comment(comment.id)


class InkCanvas(QWidget):
    def __init__(self, parent=None, ink=None):
        super().__init__(parent)
        self.setMinimumSize(720, 360)
        self.strokes = []
        self.drawing_size = QSize(720, 360)
        self.color = "#17365d"
        self.pen_width = 3
        self._drawing = False
        self.setCursor(Qt.CursorShape.CrossCursor)
        if ink is not None:
            self.load_ink(ink)

    def load_ink(self, data):
        from .review_services import checked_ink

        data = checked_ink(data)
        self.drawing_size = QSize(data["width"], data["height"])
        self.strokes = [(stroke["color"], stroke["width"],
                         [QPointF(x, y) for x, y in stroke["points"]]) for stroke in data["strokes"]]
        self.update()

    def ink(self):
        from .review_services import checked_ink

        return checked_ink({"width": self.drawing_size.width(), "height": self.drawing_size.height(),
                            "strokes": [{"color": color, "width": width,
                                         "points": [[point.x(), point.y()] for point in points]}
                                        for color, width, points in self.strokes]})

    def drawing_point(self, event):
        position = event.position()
        return QPointF(min(self.drawing_size.width(), max(0, position.x() * self.drawing_size.width() / self.width())),
                       min(self.drawing_size.height(), max(0, position.y() * self.drawing_size.height() / self.height())))

    def paint_strokes(self, painter):
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        for color, width, points in self.strokes:
            painter.setPen(QPen(QColor(color), width, Qt.PenStyle.SolidLine,
                                Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
            if len(points) == 1:
                painter.drawPoint(points[0])
            for first, second in zip(points, points[1:]):
                painter.drawLine(first, second)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("white"))
        painter.scale(self.width() / self.drawing_size.width(), self.height() / self.drawing_size.height())
        self.paint_strokes(painter)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drawing = True
            self.strokes.append((self.color, self.pen_width, [self.drawing_point(event)]))
            self.update()

    def mouseMoveEvent(self, event):
        if self._drawing:
            self.strokes[-1][2].append(self.drawing_point(event))
            self.update()

    def mouseReleaseEvent(self, event):
        self._drawing = False

    def undo(self):
        if self.strokes:
            self.strokes.pop()
            self.update()

    def clear(self):
        self.strokes.clear()
        self.update()

    def image(self):
        result = QImage(self.drawing_size, QImage.Format.Format_ARGB32_Premultiplied)
        result.fill(Qt.GlobalColor.transparent)
        painter = QPainter(result)
        self.paint_strokes(painter)
        painter.end()
        return result


class InkDialog(QDialog):
    def __init__(self, parent=None, ink=None):
        super().__init__(parent)
        self.setWindowTitle("Digital ink")
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Draw with your mouse or pen. Folio retains editable strokes and embeds an image for Word."))
        self.canvas = InkCanvas(ink=ink)
        toolbar = QHBoxLayout()
        colors = QComboBox()
        for name, color in (("Blue", "#17365d"), ("Black", "#111111"), ("Red", "#ad2337"), ("Green", "#13795b")):
            colors.addItem(name, color)
        colors.currentIndexChanged.connect(lambda *_: setattr(self.canvas, "color", colors.currentData()))
        width = QSpinBox()
        width.setRange(1, 24)
        width.setValue(3)
        width.setSuffix(" px")
        width.valueChanged.connect(lambda value: setattr(self.canvas, "pen_width", value))
        toolbar.addWidget(colors)
        toolbar.addWidget(width)
        for label, action in (("Undo stroke", self.canvas.undo), ("Clear", self.canvas.clear)):
            button = QPushButton(label)
            button.clicked.connect(action)
            toolbar.addWidget(button)
        layout.addLayout(toolbar)
        layout.addWidget(self.canvas)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Insert drawing")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
