from __future__ import annotations

import csv
from pathlib import Path

from PySide6.QtCore import QPointF, QRectF, Qt, Signal
from PySide6.QtGui import (
    QColor,
    QImage,
    QPainter,
    QPainterPath,
    QPen,
    QTextCursor,
    QTextDocument,
)
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from . import mailmerge
from .models import PageSettings, ValidationError


class FindReplaceDialog(QDialog):
    def __init__(self, editor, parent=None):
        super().__init__(parent)
        self.editor = editor
        self.setWindowTitle("Find and Replace")
        self.setModal(False)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.find_edit = QLineEdit()
        self.replace_edit = QLineEdit()
        form.addRow("Find what:", self.find_edit)
        form.addRow("Replace with:", self.replace_edit)
        layout.addLayout(form)
        options = QHBoxLayout()
        self.case_check = QCheckBox("Match case")
        self.word_check = QCheckBox("Whole words")
        options.addWidget(self.case_check)
        options.addWidget(self.word_check)
        options.addStretch(1)
        layout.addLayout(options)
        buttons = QHBoxLayout()
        self.prev_button = QPushButton("Previous")
        self.next_button = QPushButton("Next")
        self.replace_button = QPushButton("Replace")
        self.replace_all_button = QPushButton("Replace All")
        close_button = QPushButton("Close")
        for button in (self.prev_button, self.next_button,
                       self.replace_button, self.replace_all_button,
                       close_button):
            buttons.addWidget(button)
        layout.addLayout(buttons)
        self.status = QLabel("")
        layout.addWidget(self.status)
        self.next_button.clicked.connect(lambda: self.find(forward=True))
        self.prev_button.clicked.connect(lambda: self.find(forward=False))
        self.replace_button.clicked.connect(self.replace)
        self.replace_all_button.clicked.connect(self.replace_all)
        close_button.clicked.connect(self.close)
        self.find_edit.returnPressed.connect(lambda: self.find(forward=True))

    def _flags(self):
        flags = QTextDocument.FindFlag(0)
        if self.case_check.isChecked():
            flags |= QTextDocument.FindFlag.FindCaseSensitively
        if self.word_check.isChecked():
            flags |= QTextDocument.FindFlag.FindWholeWords
        return flags

    def find(self, forward: bool = True) -> bool:
        text = self.find_edit.text()
        if not text:
            self.status.setText("")
            return False
        flags = self._flags()
        if not forward:
            flags |= QTextDocument.FindFlag.FindBackward
        found = self.editor.document().find(text, self.editor.textCursor(), flags)
        if found.isNull():
            found = self.editor.document().find(text, QTextCursor(), flags)
        if found.isNull():
            self.status.setText("No matches found.")
            return False
        self.editor.setTextCursor(found)
        self.editor.ensureCursorVisible()
        self.status.setText("")
        return True

    def replace(self) -> None:
        if self.editor.isReadOnly():
            return
        cursor = self.editor.textCursor()
        text = self.find_edit.text()
        if cursor.hasSelection() and self._matches(cursor.selectedText(), text):
            cursor.insertText(self.replace_edit.text())
        self.find(forward=True)

    def _matches(self, selected: str, needle: str) -> bool:
        if self.case_check.isChecked():
            return selected == needle
        return selected.lower() == needle.lower()

    def replace_all(self) -> int:
        if self.editor.isReadOnly():
            return 0
        text = self.find_edit.text()
        if not text:
            return 0
        doc = self.editor.document()
        flags = self._flags()
        cursor = QTextCursor(doc)
        cursor.beginEditBlock()
        count = 0
        while True:
            found = doc.find(text, cursor, flags)
            if found.isNull():
                break
            found.insertText(self.replace_edit.text())
            cursor = found
            count += 1
        cursor.endEditBlock()
        self.status.setText(f"Replaced {count} occurrence(s).")
        return count


class PasteSpecialDialog(QDialog):
    def __init__(self, mime_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Paste Special")
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Paste clipboard content as:"))
        self.keep_formatting = QRadioButton("Formatted text")
        self.plain_text = QRadioButton("Unformatted text")
        self.image = QRadioButton("Picture")
        self.keep_formatting.setEnabled(mime_data.hasHtml())
        self.plain_text.setEnabled(mime_data.hasText())
        self.image.setEnabled(mime_data.hasImage())
        for button in (self.keep_formatting, self.plain_text, self.image):
            layout.addWidget(button)
        if mime_data.hasHtml():
            self.keep_formatting.setChecked(True)
        elif mime_data.hasText():
            self.plain_text.setChecked(True)
        else:
            self.image.setChecked(True)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def mode(self) -> str:
        if self.image.isChecked():
            return "image"
        if self.plain_text.isChecked():
            return "text"
        return "keep_formatting"


class PageSetupDialog(QDialog):
    def __init__(self, settings: PageSettings, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Page Setup")
        self._result = None
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.paper = QComboBox()
        self.paper.addItems(["A4", "Letter", "Legal"])
        self.paper.setCurrentText(settings.paper)
        self.orientation = QComboBox()
        self.orientation.addItems(["Portrait", "Landscape"])
        self.orientation.setCurrentIndex(1 if settings.landscape else 0)
        form.addRow("Paper size:", self.paper)
        form.addRow("Orientation:", self.orientation)
        self.top = self._mm(settings.top_mm)
        self.right = self._mm(settings.right_mm)
        self.bottom = self._mm(settings.bottom_mm)
        self.left = self._mm(settings.left_mm)
        form.addRow("Top margin (mm):", self.top)
        form.addRow("Right margin (mm):", self.right)
        form.addRow("Bottom margin (mm):", self.bottom)
        form.addRow("Left margin (mm):", self.left)
        self.columns = QSpinBox()
        self.columns.setRange(1, 3)
        self.columns.setValue(settings.columns)
        self.gutter = self._mm(settings.gutter_mm)
        form.addRow("Columns:", self.columns)
        form.addRow("Column gutter (mm):", self.gutter)
        self.balance_columns = QCheckBox("Balance the final columns")
        self.balance_columns.setChecked(settings.balance_columns)
        form.addRow("", self.balance_columns)
        self.header = QLineEdit(settings.header)
        self.footer = QLineEdit(settings.footer)
        form.addRow("Header:", self.header)
        form.addRow("Footer:", self.footer)
        self.line_numbering = QComboBox()
        for label, value in (
                ("None", "none"),
                ("Continuous", "continuous"),
                ("Restart each page", "restart_page"),
                ("Restart each section", "restart_section")):
            self.line_numbering.addItem(label, value)
        line_index = self.line_numbering.findData(settings.line_numbering)
        self.line_numbering.setCurrentIndex(max(0, line_index))
        self.line_start = QSpinBox()
        self.line_start.setRange(1, 999999)
        self.line_start.setValue(settings.line_number_start)
        self.line_count_by = QSpinBox()
        self.line_count_by.setRange(1, 1000)
        self.line_count_by.setValue(settings.line_number_count_by)
        self.line_distance = self._mm(settings.line_number_distance_mm)
        form.addRow("Line numbers:", self.line_numbering)
        form.addRow("Line numbers start at:", self.line_start)
        form.addRow("Line numbers count by:", self.line_count_by)
        form.addRow("Line number distance (mm):", self.line_distance)
        self.hyphenation = QCheckBox("Automatically hyphenate document")
        self.hyphenation.setChecked(settings.hyphenation == "automatic")
        self.hyphenate_caps = QCheckBox("Hyphenate words in CAPS")
        self.hyphenate_caps.setChecked(settings.hyphenate_caps)
        self.hyphen_limit = QSpinBox()
        self.hyphen_limit.setRange(0, 99)
        self.hyphen_limit.setSpecialValueText("No limit")
        self.hyphen_limit.setValue(settings.consecutive_hyphen_limit)
        form.addRow("Hyphenation:", self.hyphenation)
        form.addRow("", self.hyphenate_caps)
        form.addRow("Consecutive hyphens:", self.hyphen_limit)
        layout.addLayout(form)
        hint = QLabel("Use {CurrentPage}, {TotalPages}, and {title} in headers/footers "
                      "({page} and {pages} also work). "
                      "Multi-column layout and running headers are applied "
                      "in the publishing preview and exported output.")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    @staticmethod
    def _mm(value: float) -> QDoubleSpinBox:
        box = QDoubleSpinBox()
        box.setRange(0.0, 200.0)
        box.setDecimals(1)
        box.setValue(value)
        return box

    def _accept(self) -> None:
        candidate = PageSettings(
            paper=self.paper.currentText(),
            landscape=self.orientation.currentIndex() == 1,
            top_mm=self.top.value(),
            right_mm=self.right.value(),
            bottom_mm=self.bottom.value(),
            left_mm=self.left.value(),
            columns=self.columns.value(),
            gutter_mm=self.gutter.value(),
            header=self.header.text(),
            footer=self.footer.text(),
            line_numbering=str(self.line_numbering.currentData()),
            line_number_start=self.line_start.value(),
            line_number_count_by=self.line_count_by.value(),
            line_number_distance_mm=self.line_distance.value(),
            hyphenation=("automatic" if self.hyphenation.isChecked()
                         else "none"),
            hyphenate_caps=self.hyphenate_caps.isChecked(),
            consecutive_hyphen_limit=self.hyphen_limit.value(),
            balance_columns=self.balance_columns.isChecked(),
        )
        try:
            candidate.validate()
        except ValidationError as exc:
            QMessageBox.warning(self, "Page Setup", str(exc))
            return
        self._result = candidate
        self.accept()

    def settings(self) -> PageSettings | None:
        return self._result


class ParagraphDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Paragraph")
        self._result = None
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.before = self._pt(0.0)
        self.after = self._pt(8.0)
        self.spacing = QComboBox()
        for label, value in (("Single", 1.0), ("1.15", 1.15),
                             ("1.5", 1.5), ("Double", 2.0)):
            self.spacing.addItem(label, value)
        self.spacing.setCurrentIndex(1)
        self.left = self._mm(0.0)
        self.right = self._mm(0.0)
        self.first = self._mm(0.0)
        self.first.setRange(-50.0, 100.0)
        form.addRow("Space before (pt):", self.before)
        form.addRow("Space after (pt):", self.after)
        form.addRow("Line spacing:", self.spacing)
        form.addRow("Left indent (mm):", self.left)
        form.addRow("Right indent (mm):", self.right)
        form.addRow("First line (mm):", self.first)
        layout.addLayout(form)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    @staticmethod
    def _pt(value: float) -> QDoubleSpinBox:
        box = QDoubleSpinBox()
        box.setRange(0.0, 500.0)
        box.setValue(value)
        return box

    @staticmethod
    def _mm(value: float) -> QDoubleSpinBox:
        box = QDoubleSpinBox()
        box.setRange(0.0, 100.0)
        box.setDecimals(1)
        box.setValue(value)
        return box

    def _accept(self) -> None:
        self._result = (
            self.before.value(),
            self.after.value(),
            float(self.spacing.currentData()),
            self.left.value(),
            self.right.value(),
            self.first.value(),
        )
        self.accept()

    def values(self):
        return self._result


class TableDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Table")
        self._result = None
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.rows = QSpinBox()
        self.rows.setRange(1, 100)
        self.rows.setValue(3)
        self.cols = QSpinBox()
        self.cols.setRange(1, 20)
        self.cols.setValue(3)
        form.addRow("Rows:", self.rows)
        form.addRow("Columns:", self.cols)
        layout.addLayout(form)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def _accept(self) -> None:
        self._result = (self.rows.value(), self.cols.value())
        self.accept()

    def values(self):
        return self._result


class CoverPageDialog(QDialog):
    def __init__(self, parent=None, title: str = ""):
        super().__init__(parent)
        self.setWindowTitle("Insert Cover Page")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.style = QComboBox()
        self.style.addItems(["ViewMaster", "Facet", "Ion"])
        self.title = QLineEdit(title)
        self.subtitle = QLineEdit()
        form.addRow("Design:", self.style)
        form.addRow("Title:", self.title)
        form.addRow("Subtitle:", self.subtitle)
        layout.addLayout(form)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self.accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def values(self) -> tuple[str, str, str]:
        return (self.style.currentText(), self.title.text(),
                self.subtitle.text())


class ChartDialog(QDialog):
    COLORS = ("#2f6fd0", "#4aa3df", "#43a047", "#f9a825", "#8e5bb7",
              "#d65a4a")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Chart")
        self._result = None
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.chart_type = QComboBox()
        self.chart_type.addItems(["Bar", "Line", "Pie"])
        self.data = QPlainTextEdit()
        self.data.setPlainText("North,42\nSouth,31\nEast,27\nWest,36")
        self.data.setPlaceholderText("One label,value pair per line")
        form.addRow("Chart type:", self.chart_type)
        form.addRow("Data:", self.data)
        layout.addLayout(form)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def _accept(self) -> None:
        rows: list[tuple[str, float]] = []
        try:
            for raw in self.data.toPlainText().splitlines():
                if not raw.strip():
                    continue
                label, value = raw.rsplit(",", 1)
                rows.append((label.strip() or "Item", float(value.strip())))
        except (ValueError, TypeError):
            QMessageBox.warning(
                self, "Insert Chart",
                "Enter one label and numeric value per line, separated by a comma.")
            return
        if not rows or len(rows) > 20:
            QMessageBox.warning(
                self, "Insert Chart", "Charts require between 1 and 20 values.")
            return
        self._result = (self.chart_type.currentText(), rows)
        self.accept()

    def values(self):
        return self._result

    @classmethod
    def render(cls, chart_type: str, rows: list[tuple[str, float]]) -> QImage:
        image = QImage(800, 480, QImage.Format.Format_ARGB32)
        image.fill(QColor("white"))
        painter = QPainter(image)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        plot = QRectF(80, 40, 650, 350)
        painter.setPen(QPen(QColor("#cbd5e1"), 1))
        if chart_type == "Pie":
            total = sum(max(0.0, value) for _label, value in rows) or 1.0
            angle = 0
            pie = QRectF(210, 45, 340, 340)
            for index, (label, value) in enumerate(rows):
                span = int(round(max(0.0, value) / total * 5760))
                painter.setPen(Qt.PenStyle.NoPen)
                painter.setBrush(QColor(cls.COLORS[index % len(cls.COLORS)]))
                painter.drawPie(pie, angle, span)
                angle += span
                painter.setPen(QColor("#334155"))
                painter.drawText(585, 70 + index * 24, label)
        else:
            painter.drawLine(plot.bottomLeft(), plot.topLeft())
            painter.drawLine(plot.bottomLeft(), plot.bottomRight())
            values = [value for _label, value in rows]
            low = min(0.0, min(values))
            high = max(values)
            span = max(1.0, high - low)
            step = plot.width() / max(1, len(rows))
            points = []
            for index, (label, value) in enumerate(rows):
                x = plot.left() + step * (index + 0.5)
                y = plot.bottom() - ((value - low) / span) * plot.height()
                color = QColor(cls.COLORS[index % len(cls.COLORS)])
                if chart_type == "Bar":
                    painter.fillRect(
                        QRectF(x - step * 0.3, y, step * 0.6,
                               plot.bottom() - y), color)
                points.append((x, y, color))
                painter.setPen(QColor("#334155"))
                painter.drawText(
                    QRectF(x - step * 0.48, plot.bottom() + 8,
                           step * 0.96, 34),
                    Qt.AlignmentFlag.AlignHCenter
                    | Qt.AlignmentFlag.AlignTop, label)
            if chart_type == "Line":
                painter.setPen(QPen(QColor("#2f6fd0"), 3))
                for first, second in zip(points, points[1:]):
                    painter.drawLine(QPointF(first[0], first[1]),
                                     QPointF(second[0], second[1]))
                painter.setPen(Qt.PenStyle.NoPen)
                for x, y, color in points:
                    painter.setBrush(color)
                    painter.drawEllipse(QRectF(x - 5, y - 5, 10, 10))
        painter.end()
        return image


class LinkDialog(QDialog):
    def __init__(self, parent=None, text: str = ""):
        super().__init__(parent)
        self.setWindowTitle("Insert Link")
        self._result = None
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.url = QLineEdit()
        self.url.setPlaceholderText("https://example.com")
        self.text = QLineEdit(text)
        form.addRow("Address:", self.url)
        form.addRow("Text:", self.text)
        layout.addLayout(form)
        note = QLabel("Only http, https, and mailto links are allowed.")
        layout.addWidget(note)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def _accept(self) -> None:
        from PySide6.QtCore import QUrl
        from .editor import ALLOWED_LINK_SCHEMES

        url = self.url.text().strip()
        scheme = QUrl(url).scheme().lower()
        if scheme not in ALLOWED_LINK_SCHEMES:
            QMessageBox.warning(
                self, "Insert Link",
                "Links must start with http://, https://, or mailto:."
            )
            return
        self._result = (url, self.text.text().strip() or url)
        self.accept()

    def values(self):
        return self._result


class ImageSizeDialog(QDialog):
    def __init__(self, current_width: float, max_width: float, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Image Size")
        self._result = None
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.width = QDoubleSpinBox()
        self.width.setRange(8.0, max(8.0, max_width))
        self.width.setSuffix(" px")
        self.width.setValue(min(current_width, max_width))
        form.addRow("Width:", self.width)
        layout.addLayout(form)
        note = QLabel("Height is adjusted automatically to keep the "
                      "image's aspect ratio.")
        note.setWordWrap(True)
        layout.addWidget(note)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def _accept(self) -> None:
        self._result = self.width.value()
        self.accept()

    def values(self):
        return self._result


class ShapeDialog(QDialog):
    SHAPES = ["Rectangle", "Rounded rectangle", "Ellipse", "Line", "Arrow"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert Shape")
        self._result = None
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.shape = QComboBox()
        self.shape.addItems(self.SHAPES)
        self.color = QLineEdit("#185abd")
        self.label = QLineEdit()
        form.addRow("Shape:", self.shape)
        form.addRow("Outline color:", self.color)
        form.addRow("Label (optional):", self.label)
        layout.addLayout(form)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def _accept(self) -> None:
        color = QColor(self.color.text().strip())
        if not color.isValid():
            QMessageBox.warning(self, "Insert Shape",
                                "Enter a color like #185abd.")
            return
        self._result = (self.shape.currentText(), color,
                        self.label.text().strip())
        self.accept()

    def values(self):
        return self._result

    @staticmethod
    def render(shape: str, color: QColor, label: str) -> QImage:
        image = QImage(360, 140, QImage.Format.Format_ARGB32)
        image.fill(Qt.GlobalColor.transparent)
        painter = QPainter(image)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        pen = QPen(color, 3)
        painter.setPen(pen)
        painter.setBrush(QColor(color.red(), color.green(), color.blue(), 40))
        rect = image.rect().adjusted(12, 20, -12, -40)
        if shape == "Rectangle":
            painter.drawRect(rect)
        elif shape == "Rounded rectangle":
            painter.drawRoundedRect(rect, 12, 12)
        elif shape == "Ellipse":
            painter.drawEllipse(rect)
        elif shape == "Line":
            painter.drawLine(12, image.height() // 2,
                             image.width() - 12, image.height() // 2)
        elif shape == "Arrow":
            y = image.height() // 2
            painter.drawLine(12, y, image.width() - 40, y)
            path = QPainterPath()
            path.moveTo(image.width() - 12, y)
            path.lineTo(image.width() - 44, y - 14)
            path.lineTo(image.width() - 44, y + 14)
            path.closeSubpath()
            painter.setBrush(color)
            painter.drawPath(path)
        if label:
            painter.setPen(QColor("#263244"))
            font = painter.font()
            font.setPointSize(11)
            painter.setFont(font)
            painter.drawText(image.rect().adjusted(8, 8, -8, -8),
                             Qt.AlignmentFlag.AlignHCenter
                             | Qt.AlignmentFlag.AlignVCenter, label)
        painter.end()
        return image


class DiagramDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Process Diagram")
        self._result = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Enter 3–5 step labels, one per line:"))
        self.steps = QPlainTextEdit()
        self.steps.setPlainText("Plan\nBuild\nReview\nShip")
        layout.addWidget(self.steps)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def _accept(self) -> None:
        steps = [s.strip() for s in self.steps.toPlainText().splitlines()
                 if s.strip()]
        if not 3 <= len(steps) <= 5:
            QMessageBox.warning(self, "Process Diagram",
                                "Enter between 3 and 5 steps.")
            return
        self._result = steps
        self.accept()

    def values(self):
        return self._result

    @staticmethod
    def render(steps: list[str]) -> QImage:
        count = len(steps)
        width, height = 960, 220
        image = QImage(width, height, QImage.Format.Format_ARGB32)
        image.fill(Qt.GlobalColor.transparent)
        painter = QPainter(image)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        accent = QColor("#185abd")
        box_w = (width - 40) // count - 24
        box_h = 96
        top = (height - box_h) // 2
        font = painter.font()
        font.setPointSize(12)
        font.setBold(True)
        painter.setFont(font)
        for index, step in enumerate(steps):
            x = 20 + index * ((width - 40) // count)
            rect = image.rect().adjusted(0, 0, 0, 0)
            del rect
            painter.setPen(QPen(accent, 2.5))
            painter.setBrush(QColor(220, 232, 249))
            painter.drawRoundedRect(x, top, box_w, box_h, 10, 10)
            painter.setPen(QColor("#17365d"))
            painter.drawText(x, top, box_w, box_h,
                             Qt.AlignmentFlag.AlignCenter
                             | Qt.TextFlag.TextWordWrap, step)
            if index < count - 1:
                ax = x + box_w + 2
                ay = top + box_h // 2
                painter.setPen(QPen(accent, 2.5))
                painter.drawLine(ax, ay, ax + 20, ay)
                path = QPainterPath()
                path.moveTo(ax + 22, ay)
                path.lineTo(ax + 14, ay - 5)
                path.lineTo(ax + 14, ay + 5)
                path.closeSubpath()
                painter.setBrush(accent)
                painter.drawPath(path)
        painter.end()
        return image


class TemplateGalleryDialog(QDialog):
    DESCRIPTIONS = {
        "blank": "Blank — an empty document",
        "welcome": "Welcome — a short tour of Folio",
        "report": "Project report — headings, table, and lists",
        "resume": "Resume — sections with merge placeholders",
        "invoice": "Invoice — table with {{placeholders}}",
        "essay": "Academic essay — title, sections, and a quote",
    }

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("New Document")
        self._result = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Choose a starting point:"))
        self.list = QListWidget()
        from .templates import TEMPLATE_NAMES
        for name in TEMPLATE_NAMES:
            item = QListWidgetItem(self.DESCRIPTIONS.get(name, name))
            item.setData(Qt.ItemDataRole.UserRole, name)
            self.list.addItem(item)
        self.list.setCurrentRow(0)
        layout.addWidget(self.list)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)
        self.list.itemDoubleClicked.connect(lambda _i: self._accept())

    def _accept(self) -> None:
        item = self.list.currentItem()
        if item is None:
            return
        self._result = item.data(Qt.ItemDataRole.UserRole)
        self.accept()

    def choice(self) -> str | None:
        return self._result


class AuthorDialog(QDialog):
    def __init__(self, current: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Author Name")
        self._result = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Name used for comments and tracked changes:"))
        self.name = QLineEdit(current)
        layout.addWidget(self.name)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                               | QDialogButtonBox.StandardButton.Cancel)
        box.accepted.connect(self._accept)
        box.rejected.connect(self.reject)
        layout.addWidget(box)

    def _accept(self) -> None:
        self._result = self.name.text().strip() or "You"
        self.accept()

    def value(self) -> str | None:
        return self._result


class RecoveryDialog(QDialog):
    def __init__(self, entries: list[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Recover Documents")
        self._result = None
        self._discard = False
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Folio found unsaved work from a previous session:"))
        self.list = QListWidget()
        for entry in entries:
            title = entry.get("title") or "Untitled"
            path = entry.get("path") or "(not saved yet)"
            updated = entry.get("updated") or ""
            item = QListWidgetItem(f"{title} — {path} — {updated}")
            item.setData(Qt.ItemDataRole.UserRole, entry.get("doc_id"))
            self.list.addItem(item)
        if self.list.count():
            self.list.setCurrentRow(0)
        layout.addWidget(self.list)
        buttons = QHBoxLayout()
        restore = QPushButton("Recover selected")
        discard = QPushButton("Discard selected")
        skip = QPushButton("Skip")
        buttons.addWidget(restore)
        buttons.addWidget(discard)
        buttons.addStretch(1)
        buttons.addWidget(skip)
        layout.addLayout(buttons)
        restore.clicked.connect(self._restore)
        discard.clicked.connect(self._discard_clicked)
        skip.clicked.connect(self.reject)

    def _restore(self) -> None:
        item = self.list.currentItem()
        if item is None:
            return
        self._result = item.data(Qt.ItemDataRole.UserRole)
        self.accept()

    def _discard_clicked(self) -> None:
        item = self.list.currentItem()
        if item is None:
            return
        self._result = item.data(Qt.ItemDataRole.UserRole)
        self._discard = True
        self.accept()

    def choice(self):
        return self._result, self._discard


class MailMergeDialog(QDialog):
    def __init__(self, document, state, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Mail Merge")
        self.document = document
        self.state = state
        self.records: list[dict[str, str]] = []
        self.outputs: list[Path] = []
        self._cancelled = False
        layout = QVBoxLayout(self)
        row = QHBoxLayout()
        self.source_label = QLabel("No merge source selected.")
        pick = QPushButton("Choose CSV/XLSX…")
        self._pick_source_button = pick
        row.addWidget(self.source_label, 1)
        row.addWidget(pick)
        layout.addLayout(row)
        self.preview = QLabel("")
        self.preview.setWordWrap(True)
        layout.addWidget(self.preview)
        out_row = QHBoxLayout()
        self.dir_label = QLabel("Output folder: (not set)")
        pick_dir = QPushButton("Choose folder…")
        self._pick_dir_button = pick_dir
        out_row.addWidget(self.dir_label, 1)
        out_row.addWidget(pick_dir)
        layout.addLayout(out_row)
        self.directory: Path | None = None
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)
        buttons = QHBoxLayout()
        self.run_button = QPushButton("Generate documents")
        self.run_button.setEnabled(False)
        cancel = QPushButton("Close")
        buttons.addStretch(1)
        buttons.addWidget(self.run_button)
        buttons.addWidget(cancel)
        layout.addLayout(buttons)
        pick.clicked.connect(self._pick_source)
        pick_dir.clicked.connect(self._pick_dir)
        self.run_button.clicked.connect(self._run)
        cancel.clicked.connect(self._close)

    def _pick_source(self) -> None:
        from PySide6.QtWidgets import QFileDialog

        path, _ = QFileDialog.getOpenFileName(
            self, "Choose merge source", "",
            "Data sources (*.csv *.xlsx)")
        if not path:
            return
        try:
            self.records = mailmerge.read_records(Path(path))
        except (mailmerge.MergeError, OSError, UnicodeError, csv.Error) as exc:
            QMessageBox.warning(self, "Mail Merge",
                                f"Could not read the merge source:\n{exc}")
            return
        self.source_label.setText(Path(path).name)
        headers = list(self.records[0].keys()) if self.records else []
        fields = mailmerge.discover_fields(self.document)
        missing = [f for f in fields if f not in headers]
        preview = (f"{len(self.records)} record(s); "
                   f"columns: {', '.join(headers) or '(none)'}\n"
                   f"Fields in document: {', '.join(fields) or '(none)'}")
        if self.records:
            first = ", ".join(f"{k}={v}" for k, v in
                              list(self.records[0].items())[:6])
            preview += f"\nFirst record: {first}"
        if missing:
            preview += f"\nMissing columns: {', '.join(missing)}"
        self.preview.setText(preview)
        self._refresh()

    def _pick_dir(self) -> None:
        from PySide6.QtWidgets import QFileDialog

        directory = QFileDialog.getExistingDirectory(
            self, "Choose output folder")
        if directory:
            self.directory = Path(directory)
            self.dir_label.setText(f"Output folder: {directory}")
        self._refresh()

    def _refresh(self) -> None:
        self.run_button.setEnabled(bool(self.records) and self.directory is not None)

    def _close(self) -> None:
        self._cancelled = True
        self.close()

    def _run(self) -> None:
        if self.directory is None:
            return
        self.progress.setVisible(True)
        self.progress.setRange(0, len(self.records))
        self.progress.setValue(0)
        self._completed = 0

        def progress(done, total):
            self._completed = done
            self.progress.setValue(done)
            from PySide6.QtWidgets import QApplication
            QApplication.processEvents()

        controls = (self.run_button, self._pick_source_button,
                    self._pick_dir_button)
        for widget in controls:
            widget.setEnabled(False)
        try:
            self.outputs = mailmerge.merge_to_directory(
                self.document, self.state, self.records, self.directory,
                progress=progress, cancelled=lambda: self._cancelled)
        except mailmerge.MergeError as exc:
            QMessageBox.warning(self, "Mail Merge", str(exc))
            return
        except Exception as exc:
            QMessageBox.warning(
                self, "Mail Merge",
                f"Merge failed after {self._completed} document(s) were "
                f"generated:\n{exc}")
            return
        finally:
            for widget in controls:
                widget.setEnabled(True)
            self._refresh()
        QMessageBox.information(
            self, "Mail Merge",
            f"Created {len(self.outputs)} document(s) in {self.directory}."
        )
        self.accept()


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About Folio")
        layout = QVBoxLayout(self)
        title = QLabel("<b>Folio</b> — local-first word processor")
        title.setStyleSheet("font-size:16px;")
        layout.addWidget(title)
        text = QLabel(
            "Folio edits and saves real .docx files entirely on this "
            "computer. No account or network connection is required.\n\n"
            "Honest limits:\n"
            "• Spelling, grammar, and style checks use basic offline "
            "heuristics — not a full grammar engine.\n"
            "• Tracked changes cover text edits and optional formatting edits, "
            "not structural changes, and are Folio metadata rather than "
            "Word-native revisions.\n"
            "• External .docx files are imported with best effort; "
            "advanced Word features may not survive.\n"
            "• Editing shows a single flowing page-width canvas; columns "
            "section geometry, object wrapping, and running headers/footers are rendered in the publishing "
            "preview and exported PDF/print output.\n"
            "• Equations, diagrams, SVG graphics, and 3D models export as pictures; "
            "equation/diagram sources remain editable, and supported static 3D mesh data is retained in Folio.\n"
            "• Typed signatures are local consent records, not verified digital signatures. "
            "Embedded files are passive attachments, not executable OLE controls.\n"
            "• Transcription supplies local integration hooks, not a built-in audio recognizer. "
            "Preview hyphenation uses a small bundled English dictionary.\n"
            "• Embedded Folio metadata is limited to 16 MiB per file; "
            "documents that exceed it cannot be saved.\n"
            "• Citations use common style patterns; accessibility checks assist "
            "human review and do not certify compliance.\n"
            "• Translation is local glossary substitution, not machine translation. "
            "Email drafts and feedback are exported locally, not sent.\n"
            "• JavaScript macros use Folio's document API; VBA is not supported. "
            "Editing restrictions are local safeguards, not file encryption."
        )
        text.setWordWrap(True)
        layout.addWidget(text)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        box.accepted.connect(self.accept)
        layout.addWidget(box)


class VersionHistoryDialog(QDialog):
    restoreRequested = Signal(int)

    def __init__(self, versions: list[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Version History")
        self._selected = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Local snapshots for this document:"))
        self.list = QListWidget()
        for version in versions:
            item = QListWidgetItem(
                f"{version['created']} — {version['label']}")
            item.setData(Qt.ItemDataRole.UserRole, version["id"])
            self.list.addItem(item)
        if self.list.count():
            self.list.setCurrentRow(0)
        layout.addWidget(self.list)
        buttons = QHBoxLayout()
        restore = QPushButton("Restore selected")
        close = QPushButton("Close")
        buttons.addStretch(1)
        buttons.addWidget(restore)
        buttons.addWidget(close)
        layout.addLayout(buttons)
        restore.clicked.connect(self._restore)
        close.clicked.connect(self.reject)

    def _restore(self) -> None:
        item = self.list.currentItem()
        if item is None:
            return
        self._selected = item.data(Qt.ItemDataRole.UserRole)
        self.accept()

    def selected(self):
        return self._selected
