from __future__ import annotations

import difflib
import html
import re
from urllib.parse import urlparse

from PySide6.QtCore import QPointF, QRectF, QSize, Qt
from PySide6.QtGui import (
    QColor, QFont, QIcon, QImage, QPainter, QPainterPath, QPen, QPixmap,
    QTextCharFormat, QTextCursor, QTextFormat,
)
from PySide6.QtWidgets import (
    QCheckBox, QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMessageBox,
    QPushButton, QScrollArea, QSpinBox, QTabWidget, QTextEdit, QVBoxLayout, QWidget,
)

from .editor import COLUMN_BREAK_MARKER, SafeDocument
from .publishing import Publication


VIDEO_HOSTS = {
    "youtube.com": "YouTube", "youtu.be": "YouTube",
    "vimeo.com": "Vimeo", "slideshare.net": "SlideShare", "ted.com": "TED",
}
PALETTES = {
    "Office Blue": ("#17365d", "#263244"),
    "Forest": ("#166534", "#24362c"),
    "Plum": ("#6b21a8", "#36243e"),
    "Monochrome": ("#222222", "#333333"),
}
STYLE_SETS = {"Modern": (22, 16, 12, 11), "Compact": (18, 14, 12, 10),
              "Large Print": (28, 22, 18, 16)}
SYNONYMS = {
    "good": ("excellent", "fine", "helpful", "sound"),
    "bad": ("poor", "harmful", "unfavorable"),
    "big": ("large", "substantial", "considerable"),
    "small": ("little", "compact", "minor"),
    "important": ("significant", "essential", "notable"),
    "show": ("display", "demonstrate", "reveal"),
    "make": ("create", "produce", "build"),
    "use": ("apply", "employ", "utilize"),
    "help": ("assist", "support", "aid"),
    "clear": ("plain", "distinct", "evident"),
    "fast": ("quick", "rapid", "swift"),
    "slow": ("gradual", "leisurely", "unhurried"),
    "happy": ("glad", "pleased", "joyful"),
    "sad": ("unhappy", "sorrowful", "dejected"),
    "start": ("begin", "commence", "launch"),
    "end": ("finish", "conclude", "complete"),
    "idea": ("concept", "notion", "proposal"),
    "change": ("alter", "modify", "adjust"),
    "think": ("consider", "reflect", "reason"),
    "say": ("state", "express", "mention"),
    "need": ("require", "want", "necessity"),
    "plan": ("proposal", "strategy", "outline"),
    "simple": ("plain", "straightforward", "uncomplicated"),
    "difficult": ("hard", "challenging", "demanding"),
    "beautiful": ("lovely", "attractive", "elegant"),
}


def video_provider(url: str) -> str:
    parsed = urlparse(url.strip())
    host = (parsed.hostname or "").lower()
    if parsed.scheme != "https" or parsed.username or parsed.password:
        raise ValueError("Enter an HTTPS video address without login information.")
    for domain, provider in VIDEO_HOSTS.items():
        if host == domain or host.endswith("." + domain):
            if parsed.path.strip("/"):
                return provider
    raise ValueError("Enter a YouTube, Vimeo, SlideShare, or TED video address.")


def document_statistics(document, page, title: str) -> dict[str, int]:
    text = document.toPlainText().replace(COLUMN_BREAK_MARKER, "").replace("\ufffc", "")
    publication = Publication(document, page, title)
    lines = 0
    block = publication.document.begin()
    while block.isValid():
        lines += block.layout().lineCount()
        block = block.next()
    return {
        "Pages": publication.page_count(),
        "Words": len(re.findall(r"\b\w+(?:['’\-]\w+)*\b", text)),
        "Characters (with spaces)": len(text.replace("\n", "")),
        "Characters (without spaces)": sum(not ch.isspace() for ch in text),
        "Paragraphs": sum(bool(p.strip()) for p in text.splitlines()),
        "Lines": lines,
    }


def comparison_html(original: str, revised: str) -> str:
    chunks = ["<p>Text comparison: <span style='background:#fee2e2'>deletions</span> "
              "and <span style='background:#dcfce7'>insertions</span>. "
              "Formatting and images are not compared.</p>"]
    before = original.splitlines(keepends=True)
    after = revised.splitlines(keepends=True)
    matcher = difflib.SequenceMatcher(None, before, after, autojunk=False)
    changes = 0
    for tag, a, b, c, d in matcher.get_opcodes():
        if tag == "equal":
            chunks.append("<pre>" + html.escape("".join(before[a:b])) + "</pre>")
            continue
        changes += 1
        if tag in ("delete", "replace"):
            chunks.append("<pre style='background:#fee2e2'>− "
                          + html.escape("".join(before[a:b])) + "</pre>")
        if tag in ("insert", "replace"):
            chunks.append("<pre style='background:#dcfce7'>+ "
                          + html.escape("".join(after[c:d])) + "</pre>")
    if not changes:
        chunks.insert(0, "<p>No text differences.</p>")
    return "".join(chunks)


def apply_design(editor, palette: str, fonts: tuple[str, str], style_set: str) -> None:
    heading_color, body_color = PALETTES[palette]
    sizes = STYLE_SETS[style_set]
    cursor = QTextCursor(editor.document())
    cursor.beginEditBlock()
    block = editor.document().begin()
    while block.isValid():
        level = block.blockFormat().headingLevel()
        fmt = QTextCharFormat()
        fmt.setFontFamilies([fonts[0] if level else fonts[1]])
        fmt.setFontPointSize(sizes[min(level - 1, 2)] if level else sizes[3])
        fmt.setForeground(QColor(heading_color if level else body_color))
        iterator = block.begin()
        while not iterator.atEnd():
            fragment = iterator.fragment()
            if (fragment.isValid() and not fragment.charFormat().isImageFormat()
                    and not fragment.charFormat().isAnchor()):
                edit = QTextCursor(editor.document())
                edit.setPosition(fragment.position())
                edit.setPosition(fragment.position() + fragment.length(),
                                 QTextCursor.MoveMode.KeepAnchor)
                edit.mergeCharFormat(fmt)
            iterator += 1
        block = block.next()
    cursor.endEditBlock()


def synonyms(word: str) -> list[str]:
    key = word.lower().strip()
    if key in SYNONYMS:
        return list(SYNONYMS[key])
    for base, values in SYNONYMS.items():
        if key in values:
            return [base, *(value for value in values if value != key)]
    return []


def syllable_guide(text: str) -> str:
    def split(match):
        word = match.group()
        if len(word) < 6:
            return word
        return re.sub(r"([aeiouyAEIOUY])([bcdfghjklmnpqrstvwxz])(?=[aeiouy])",
                      r"\1·\2", word)
    return re.sub(r"[A-Za-z]+", split, text)


def buttons(dialog, layout):
    box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok
                           | QDialogButtonBox.StandardButton.Cancel)
    box.accepted.connect(dialog.accept)
    box.rejected.connect(dialog.reject)
    layout.addWidget(box)
    return box


class VideoDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Online Video Link")
        layout = QVBoxLayout(self)
        note = QLabel("YouTube • Vimeo • SlideShare • TED\n"
                      "Insert a clickable link. Folio does not download or play videos.")
        note.setWordWrap(True)
        layout.addWidget(note)
        self.url = QLineEdit()
        self.url.setPlaceholderText("https://www.youtube.com/watch?v=…")
        self.title = QLineEdit()
        self.title.setPlaceholderText("Video title (optional)")
        layout.addWidget(self.url)
        layout.addWidget(self.title)
        self.error = QLabel()
        layout.addWidget(self.error)
        buttons(self, layout)

    def accept(self):
        try:
            video_provider(self.url.text())
        except ValueError as exc:
            self.error.setText(str(exc))
            return
        super().accept()


class TextInsertDialog(QDialog):
    SYMBOLS = {
        "Common": "© ® ™ ° ± × ÷ € £ ¥ § ¶ • … – — ✓ ∞",
        "Greek": "α β γ δ ε θ λ μ π σ φ ψ ω Γ Δ Θ Λ Σ Φ Ω",
        "Mathematics": "∀ ∂ ∃ ∅ ∇ ∈ ∉ ∑ √ ∫ ≈ ≠ ≤ ≥ ⊂ ⊃ ∩ ∪",
    }
    EQUATIONS = ("a² + b² = c²", "x = (−b ± √(b² − 4ac)) / 2a",
                 "E = mc²", "∑ᵢ₌₁ⁿ i = n(n + 1) / 2", "∫ xⁿ dx = xⁿ⁺¹ / (n + 1) + C")

    def __init__(self, kind: str, parent=None):
        super().__init__(parent)
        self.kind = kind
        self.setWindowTitle("Insert " + kind)
        layout = QVBoxLayout(self)
        self.text = QLineEdit()
        self.style = QComboBox()
        if kind == "WordArt":
            self.style.addItems(["Blue Outline", "Plum Bold", "Gold Italic"])
            self.text.setText("Your text here")
        elif kind == "Equation":
            layout.addWidget(QLabel("Editable Unicode equations; no equation layout engine."))
            self.style.addItems(self.EQUATIONS)
            self.text.setText(self.EQUATIONS[0])
            self.style.currentTextChanged.connect(self.text.setText)
        else:
            self.style.addItems(list(self.SYMBOLS))
            self.symbols = QListWidget()
            self.symbols.setFlow(QListWidget.Flow.LeftToRight)
            self.symbols.setWrapping(True)
            self.symbols.setGridSize(QSize(45, 40))
            self.symbols.itemClicked.connect(lambda item: self.text.setText(item.text()))
            self.style.currentTextChanged.connect(self._symbols)
            self._symbols(self.style.currentText())
            layout.addWidget(self.symbols)
            self.text.setPlaceholderText("Select a symbol or enter a Unicode character")
        layout.addWidget(self.style)
        layout.addWidget(self.text)
        self.box = buttons(self, layout)
        self.text.textChanged.connect(lambda value: self.box.button(
            QDialogButtonBox.StandardButton.Ok).setEnabled(bool(value.strip())))
        self.box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(
            bool(self.text.text().strip()))

    def _symbols(self, category):
        self.symbols.clear()
        self.symbols.addItems(self.SYMBOLS[category].split())

    def insert(self, editor):
        fmt = QTextCharFormat(editor.textCursor().charFormat())
        if self.kind == "WordArt":
            fmt.setFontPointSize(32)
            fmt.setFontWeight(QFont.Weight.Bold)
            colors = ["#185abd", "#7c3aed", "#996515"]
            color = QColor(colors[self.style.currentIndex()])
            fmt.setForeground(color)
            if self.style.currentIndex() == 0:
                fmt.setTextOutline(QPen(color, 1))
                fmt.setForeground(QColor("#e0ecff"))
            fmt.setFontItalic(self.style.currentIndex() == 2)
        elif self.kind == "Equation":
            fmt.setFontFamilies(["Cambria Math"])
            fmt.setFontPointSize(16)
        cursor = editor.textCursor()
        previous = cursor.charFormat()
        cursor.insertText(self.text.text(), fmt)
        cursor.setCharFormat(previous)
        editor.setTextCursor(cursor)


class SignatureDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Signature Fields")
        layout = QVBoxLayout(self)
        note = QLabel("Insert printable signer, signature, and date fields.\n"
                      "Electronic signing, identity checks, certificates, and an audit "
                      "trail require a signing service; this field is not a digital signature.")
        note.setWordWrap(True)
        layout.addWidget(note)
        form = QFormLayout()
        self.name = QLineEdit()
        self.role = QLineEdit()
        self.email = QLineEdit()
        form.addRow("Signer name:", self.name)
        form.addRow("Title / role:", self.role)
        form.addRow("Email (optional):", self.email)
        layout.addLayout(form)
        buttons(self, layout)

    def field_html(self):
        values = [self.name.text().strip() or "Signer name", self.role.text().strip(),
                  self.email.text().strip()]
        return ("<p>Signature: ______________________________</p><p>"
                + "<br/>".join(html.escape(value) for value in values if value)
                + "</p><p>Date: ____________________</p>")


def stock_art(category: str, variant: int) -> QImage:
    image = QImage(480, 320, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    colors = ["#2563eb", "#059669", "#9333ea", "#ea580c"]
    color = QColor(colors[variant % 4])
    painter.setPen(QPen(color, 10))
    painter.setBrush(color)
    if category == "Icons":
        if variant == 0:
            path = QPainterPath(QPointF(140, 165))
            path.lineTo(210, 230)
            path.lineTo(345, 90)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawPath(path)
        elif variant == 1:
            painter.drawEllipse(QRectF(150, 60, 180, 180))
            painter.setPen(QPen(QColor("white"), 10))
            painter.drawLine(240, 90, 240, 150)
            painter.drawLine(240, 150, 290, 180)
        elif variant == 2:
            painter.drawRoundedRect(QRectF(100, 70, 280, 180), 12, 12)
            painter.setPen(QPen(QColor("white"), 8))
            painter.drawLine(110, 80, 240, 175)
            painter.drawLine(240, 175, 370, 80)
        else:
            path = QPainterPath(QPointF(240, 35))
            for x, y in [(280, 125), (380, 135), (305, 205), (325, 300),
                         (240, 250), (155, 300), (175, 205), (100, 135), (200, 125)]:
                path.lineTo(x, y)
            path.closeSubpath()
            painter.drawPath(path)
    elif category == "Cartoon People":
        painter.drawRoundedRect(QRectF(145, 155, 190, 145), 45, 45)
        painter.setBrush(QColor(["#f1c7a5", "#9b6846", "#dfad88", "#63402d"][variant]))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(QRectF(185, 40, 110, 125))
        painter.setBrush(QColor("#263244"))
        painter.drawEllipse(QRectF(208, 88, 9, 9))
        painter.drawEllipse(QRectF(263, 88, 9, 9))
        painter.setPen(QPen(QColor("#263244"), 4))
        painter.drawArc(QRectF(220, 105, 40, 25), 180 * 16, 180 * 16)
    elif category == "Stickers":
        painter.setBrush(QColor("#fef3c7"))
        painter.drawRoundedRect(QRectF(55, 55, 370, 210), 65, 65)
        painter.setFont(QFont("Segoe UI", 30, QFont.Weight.Bold))
        painter.drawText(image.rect(), Qt.AlignmentFlag.AlignCenter,
                         ["Great work!", "Thank you", "Well done", "Let's go!"][variant])
    else:
        image.fill(QColor(["#dbeafe", "#d1fae5", "#f3e8ff", "#ffedd5"][variant]))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor("#fbbf24"))
        painter.drawEllipse(QRectF(335, 30, 75, 75))
        path = QPainterPath(QPointF(0, 320))
        for x, y in [(0, 240), (150, 70), (270, 230), (350, 135), (480, 240), (480, 320)]:
            path.lineTo(x, y)
        path.closeSubpath()
        painter.setBrush(color)
        painter.drawPath(path)
    painter.end()
    return image


class StockMediaDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Stock Media — Local Collection")
        self.resize(650, 420)
        self.selected_image = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Built-in Folio artwork and images from your computer."))
        self.tabs = QTabWidget()
        for category in ("Icons", "Stickers", "Illustrations", "Cartoon People", "Images"):
            page = QWidget()
            page_layout = QVBoxLayout(page)
            if category == "Images":
                choose = QPushButton("Choose image from computer…")
                choose.clicked.connect(self._choose_image)
                page_layout.addWidget(choose)
                self.filename = QLabel("No image selected")
                page_layout.addWidget(self.filename)
                page_layout.addStretch()
            else:
                items = QListWidget()
                items.setViewMode(QListWidget.ViewMode.IconMode)
                items.setIconSize(QSize(140, 95))
                for index in range(4):
                    preview = stock_art(category, index)
                    item = QListWidgetItem(QIcon(QPixmap.fromImage(preview)),
                                           f"{category} {index + 1}")
                    item.setData(Qt.ItemDataRole.UserRole, (category, index))
                    items.addItem(item)
                items.itemClicked.connect(self._select)
                page_layout.addWidget(items)
            self.tabs.addTab(page, category)
        layout.addWidget(self.tabs)
        self.box = buttons(self, layout)
        self.box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)
        self.tabs.currentChanged.connect(self._clear_selection)

    def _clear_selection(self, _index):
        self.selected_image = None
        self.box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(False)

    def _select(self, item):
        self.selected_image = stock_art(*item.data(Qt.ItemDataRole.UserRole))
        self.box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(True)

    def _choose_image(self):
        filename, _ = QFileDialog.getOpenFileName(
            self, "Choose Image", "", "Images (*.png *.jpg *.jpeg *.bmp *.webp)")
        if not filename:
            return
        image = QImage(filename)
        if image.isNull():
            QMessageBox.warning(self, "Image", "This image could not be read.")
            return
        self.selected_image = image
        self.filename.setText(filename)
        self.box.button(QDialogButtonBox.StandardButton.Ok).setEnabled(True)


class ReadingText(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.focus_lines = 0

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self.focus_lines:
            return
        start = self.textCursor()
        start.clearSelection()
        start.movePosition(QTextCursor.MoveOperation.StartOfLine)
        end = QTextCursor(start)
        end.movePosition(QTextCursor.MoveOperation.Down,
                         QTextCursor.MoveMode.MoveAnchor, self.focus_lines - 1)
        end.movePosition(QTextCursor.MoveOperation.EndOfLine)
        top = max(0, self.cursorRect(start).top() - 4)
        bottom = max(0, self.cursorRect(end).bottom() + 4)
        painter = QPainter(self.viewport())
        color = self.palette().base().color()
        color.setAlpha(190)
        painter.fillRect(0, 0, self.viewport().width(), top, color)
        painter.fillRect(0, bottom, self.viewport().width(),
                         max(0, self.viewport().height() - bottom), color)
        painter.end()


class ReadingDialog(QDialog):
    def __init__(self, document, parent=None, web=False):
        super().__init__(parent)
        self.setWindowTitle("Web Layout" if web else "Immersive Reader")
        self.resize(1000, 750)
        self.original_text = document.toPlainText().replace(COLUMN_BREAK_MARKER, "")
        self.source_html = document.toHtml()
        self.speech = None
        layout = QVBoxLayout(self)
        tools = QHBoxLayout()
        self.syllables = QCheckBox("Syllables (English heuristic)")
        tools.addWidget(self.syllables)
        self.line_focus = QComboBox()
        for label, value in (("Line Focus: Off", 0), ("1 line", 1), ("3 lines", 3), ("5 lines", 5)):
            self.line_focus.addItem(label, value)
        tools.addWidget(self.line_focus)
        self.color = QComboBox()
        self.color.addItems(["White", "Sepia", "Dark"])
        tools.addWidget(self.color)
        self.width = QComboBox()
        for label, value in (("Narrow", 500), ("Moderate", 720), ("Wide", 1100)):
            self.width.addItem(label, value)
        self.width.setCurrentIndex(1)
        tools.addWidget(self.width)
        layout.addLayout(tools)
        row = QHBoxLayout()
        for label, slot in (("Read Aloud", self.read_aloud), ("Pause", self.pause_speech),
                            ("Resume", self.resume_speech), ("Stop", self.stop_speech),
                            ("Full Screen", self._fullscreen)):
            button = QPushButton(label)
            button.clicked.connect(slot)
            row.addWidget(button)
        close = QPushButton("Close")
        close.clicked.connect(self.close)
        row.addWidget(close)
        layout.addLayout(row)
        self.reader = ReadingText()
        self.reader.setReadOnly(True)
        self.reader.setDocument(SafeDocument(self.reader))
        self.reader.document().setDefaultFont(document.defaultFont())
        self.reader.setHtml(self.source_html)
        self.reader.document().setDocumentMargin(24)
        holder = QHBoxLayout()
        holder.addStretch()
        holder.addWidget(self.reader, 1)
        holder.addStretch()
        layout.addLayout(holder, 1)
        self.status = QLabel("Reading copy — changes here do not change the document.")
        layout.addWidget(self.status)
        self.syllables.toggled.connect(self._syllables)
        self.line_focus.currentIndexChanged.connect(self._focus_lines)
        self.reader.cursorPositionChanged.connect(self._focus_lines)
        self.color.currentTextChanged.connect(self._color)
        self.width.currentIndexChanged.connect(self._width)
        self._width()

    def _syllables(self, enabled):
        if enabled:
            self.reader.setPlainText(syllable_guide(self.original_text))
        else:
            self.reader.setHtml(self.source_html)
        self.reader.document().setDocumentMargin(24)
        self._color(self.color.currentText())

    def _color(self, name):
        background, foreground = {"White": ("#ffffff", "#263244"),
                                  "Sepia": ("#f4ecd8", "#423528"),
                                  "Dark": ("#202630", "#f4f4f4")}[name]
        self.reader.setStyleSheet(f"QTextEdit {{ background:{background}; color:{foreground}; }}")
        cursor = QTextCursor(self.reader.document())
        cursor.select(QTextCursor.SelectionType.Document)
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(foreground))
        cursor.mergeCharFormat(fmt)
        self._focus_lines()

    def _width(self, _index=0):
        self.reader.setMaximumWidth(int(self.width.currentData()))

    def _focus_lines(self, *_args):
        count = int(self.line_focus.currentData())
        self.reader.focus_lines = count
        self.reader.viewport().update()
        if not count:
            self.reader.setExtraSelections([])
            return
        cursor = self.reader.textCursor()
        cursor.clearSelection()
        cursor.movePosition(QTextCursor.MoveOperation.StartOfLine)
        cursor.movePosition(QTextCursor.MoveOperation.Down,
                            QTextCursor.MoveMode.KeepAnchor, count - 1)
        cursor.movePosition(QTextCursor.MoveOperation.EndOfLine,
                            QTextCursor.MoveMode.KeepAnchor)
        selection = QTextEdit.ExtraSelection()
        selection.cursor = cursor
        selection.format.setBackground(QColor("#fff0a8"))
        selection.format.setForeground(QColor("#202630"))
        selection.format.setProperty(QTextFormat.Property.FullWidthSelection, True)
        self.reader.setExtraSelections([selection])

    def read_aloud(self):
        try:
            from PySide6.QtTextToSpeech import QTextToSpeech
            if self.speech is None:
                engines = QTextToSpeech.availableEngines()
                engine = next((name for name in ("winrt", "sapi", "speechd", "darwin")
                               if name in engines), None)
                if engine is None:
                    self.status.setText("No local speech engine is installed.")
                    return
                self.speech = QTextToSpeech(engine, self)
                self.speech.errorOccurred.connect(
                    lambda _reason, message: self.status.setText(message))
            if self.speech.state() == QTextToSpeech.State.Error:
                self.status.setText(self.speech.errorString() or "The local voice is unavailable.")
                return
            self.speech.say(self.original_text)
        except (ImportError, RuntimeError) as exc:
            self.status.setText(f"Read Aloud is unavailable: {exc}")

    def pause_speech(self):
        if self.speech is not None:
            self.speech.pause()

    def resume_speech(self):
        if self.speech is not None:
            self.speech.resume()

    def stop_speech(self):
        if self.speech is not None:
            self.speech.stop()

    def _fullscreen(self):
        self.showNormal() if self.isFullScreen() else self.showFullScreen()

    def done(self, result):
        self.stop_speech()
        super().done(result)

    def closeEvent(self, event):
        self.stop_speech()
        super().closeEvent(event)


class ZoomDialog(QDialog):
    def __init__(self, percent, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Zoom")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.mode = QComboBox()
        self.mode.addItems(["Percent", "Page Width", "One Page", "Two Pages", "Multiple Pages"])
        self.percent = QSpinBox()
        self.percent.setRange(10, 500)
        self.percent.setValue(percent)
        form.addRow("View:", self.mode)
        form.addRow("Percent:", self.percent)
        layout.addLayout(form)
        buttons(self, layout)


class PagePreview(QWidget):
    def __init__(self, publication, columns, parent=None):
        super().__init__(parent)
        self.publication = publication
        self.columns = columns
        self.page_width = {1: 650, 2: 420, 3: 270}[columns]
        sizes = [publication.page_size(index) for index in range(publication.page_count())]
        self.scale = self.page_width / max(width for width, height in sizes)
        self.page_height = max(round(height * self.scale) for width, height in sizes)
        self.page_rects = []
        y = 24
        for first in range(0, len(sizes), columns):
            row_height = max(round(height * self.scale) for width, height in sizes[first:first + columns])
            for column, (width, height) in enumerate(sizes[first:first + columns]):
                x = 24 + column * (self.page_width + 24)
                self.page_rects.append(QRectF(x, y, width * self.scale, height * self.scale))
            y += row_height + 42
        self.setFixedSize(columns * (self.page_width + 24) + 24,
                          y)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(event.rect(), QColor("#e2e8f0"))
        for page, rect in enumerate(self.page_rects):
            if not rect.adjusted(0, 0, 0, 24).intersects(QRectF(event.rect())):
                continue
            painter.fillRect(rect, QColor("white"))
            painter.save()
            painter.translate(rect.topLeft())
            self.publication.paint_page(painter, page, self.scale)
            painter.restore()
            painter.setPen(QColor("#334155"))
            painter.drawText(round(rect.left()), round(rect.bottom() + 20), f"Page {page + 1}")
        painter.end()


class PagePreviewDialog(QDialog):
    def __init__(self, publication, columns=1, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Page Preview")
        self.resize(960, 760)
        layout = QVBoxLayout(self)
        self.mode = QComboBox()
        self.mode.addItems(["One Page", "Two Pages", "Multiple Pages"])
        self.mode.setCurrentIndex(columns - 1)
        layout.addWidget(self.mode)
        self.scroll = QScrollArea()
        self.scroll.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.publication = publication
        layout.addWidget(self.scroll)
        close = QPushButton("Close")
        close.clicked.connect(self.accept)
        layout.addWidget(close)
        self.mode.currentIndexChanged.connect(self._layout_pages)
        self._layout_pages(columns - 1)

    def _layout_pages(self, index):
        self.pages = PagePreview(self.publication, index + 1)
        self.scroll.setWidget(self.pages)
