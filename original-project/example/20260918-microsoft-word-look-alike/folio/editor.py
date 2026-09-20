from __future__ import annotations

import base64
import html
import json
import re
import uuid
from dataclasses import dataclass

from PySide6.QtCore import QByteArray, QBuffer, QIODevice, QMimeData, QSizeF, Qt, QUrl
from PySide6.QtGui import (
    QBrush,
    QColor,
    QFont,
    QGuiApplication,
    QImage,
    QKeyEvent,
    QPainter,
    QPen,
    QTextBlockFormat,
    QTextCharFormat,
    QTextCursor,
    QTextDocument,
    QTextDocumentFragment,
    QTextFormat,
    QTextFrameFormat,
    QTextImageFormat,
    QTextLength,
    QTextListFormat,
    QTextTableFormat,
)
from PySide6.QtWidgets import QTextEdit

MM_TO_PX = 96.0 / 25.4
PT_TO_PX = 96.0 / 72.0


def mm_to_px(mm: float) -> float:
    return mm * MM_TO_PX


def pt_to_px(points: float) -> float:
    return points * PT_TO_PX


@dataclass(frozen=True)
class StylePreset:
    name: str
    heading_level: int
    size_pt: float
    bold: bool
    color: str
    before_pt: float
    after_pt: float
    italic: bool = False
    left_margin_px: float = 0.0


STYLE_PRESETS: dict[str, StylePreset] = {
    p.name: p
    for p in (
        StylePreset("Normal", 0, 11, False, "#263244", 0, 8),
        StylePreset("Title", 0, 32, False, "#17365d", 0, 18),
        StylePreset("Subtitle", 0, 16, False, "#64748b", 0, 16),
        StylePreset("Heading 1", 1, 22, True, "#17365d", 20, 10),
        StylePreset("Heading 2", 2, 16, True, "#245c9f", 16, 8),
        StylePreset("Heading 3", 3, 12, True, "#245c9f", 12, 6),
        StylePreset("Quote", 0, 12, False, "#64748b", 10, 10, italic=True, left_margin_px=24),
    )
}

STYLE_ORDER = ["Normal", "Title", "Subtitle", "Heading 1", "Heading 2", "Heading 3", "Quote"]

ALLOWED_LINK_SCHEMES = {"http", "https", "mailto"}
OBJECT_REPLACEMENT = "￼"
COLUMN_BREAK_MARKER = "\u2063"
COLUMN_BREAK_URL = "folio:column-break"
_FIELD_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def block_break_kind(block) -> str | None:
    """Return the flow break represented by *block*, if any."""
    if block.blockFormat().pageBreakPolicy() != (
            QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore):
        return None
    iterator = block.begin()
    while not iterator.atEnd():
        fragment = iterator.fragment()
        if fragment.isValid():
            fmt = fragment.charFormat()
            if (fmt.isAnchor() and fmt.anchorHref() == COLUMN_BREAK_URL
                    and COLUMN_BREAK_MARKER in fragment.text()):
                return "column"
        iterator += 1
    return "page"


def _image_to_data_uri(image: QImage) -> str:
    data = QByteArray()
    buffer = QBuffer(data)
    buffer.open(QIODevice.OpenModeFlag.WriteOnly)
    image.save(buffer, "PNG")
    buffer.close()
    encoded = base64.b64encode(bytes(data.data())).decode("ascii")
    return "data:image/png;base64," + encoded


class SafeDocument(QTextDocument):
    def toHtml(self) -> str:
        from .formatting import serialize_block_metadata

        return serialize_block_metadata(self, super().toHtml())

    def setHtml(self, source: str) -> None:
        from .formatting import restore_block_metadata

        super().setHtml(source)
        restore_block_metadata(self, source)
        self.clearUndoRedoStacks()

    def drawContents(self, painter, *args) -> None:
        from .formatting import draw_paragraph_borders

        super().drawContents(painter, *args)
        draw_paragraph_borders(painter, self)

    def loadResource(self, resource_type, url):
        if isinstance(url, QUrl):
            raw = url.toString()
        else:
            raw = str(url)
        if raw.startswith("data:image/"):
            marker = ";base64,"
            idx = raw.find(marker)
            if idx != -1:
                try:
                    payload = bytes(QByteArray.fromPercentEncoding(
                        raw[idx + len(marker):].encode("utf-8"))).decode("ascii")
                    image = QImage.fromData(QByteArray.fromBase64(payload.encode("ascii")))
                except Exception:
                    return None
                if not image.isNull():
                    return image
        return None


class RichEditor(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        doc = SafeDocument(self)
        doc.setDefaultFont(QFont("Calibri", 11))
        self.setDocument(doc)
        self.setAcceptRichText(True)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setTabChangesFocus(False)
        self._default_char = QTextCharFormat()
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(STYLE_PRESETS["Normal"].color))
        self.setCurrentCharFormat(fmt)
        self._in_sync_height = False
        self._line_numbering = "none"
        self._line_number_start = 1
        self._line_number_count_by = 1
        self._line_number_distance_px = mm_to_px(5.0)
        doc.contentsChanged.connect(self._auto_resize)

    def safe_document(self) -> SafeDocument:
        return self.document()

    def _auto_resize(self) -> None:
        self.updateGeometry()
        self._sync_height()

    def _sync_height(self) -> None:
        if self._in_sync_height:
            return
        self._in_sync_height = True
        try:
            doc = self.document()
            page = doc.pageSize()
            layout_height = doc.documentLayout().documentSize().height()
            root = doc.rootFrame().frameFormat()
            if page.height() > 0:
                height = max(page.height(), layout_height + root.topMargin() + root.bottomMargin() + 8)
            else:
                height = layout_height + 16
            if int(self.height()) != int(height):
                self.setFixedHeight(int(height) + 2)
        finally:
            self._in_sync_height = False

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        self._sync_height()

    def content_width_px(self) -> float:
        doc = self.document()
        page = doc.pageSize()
        if page.width() > 0:
            root = doc.rootFrame().frameFormat()
            return page.width() - root.leftMargin() - root.rightMargin()
        return max(200.0, self.width() - 40.0)

    def apply_page_settings(self, settings) -> None:
        from .models import PageSettings

        if not isinstance(settings, PageSettings):
            raise TypeError("settings must be a PageSettings")
        doc = self.document()
        width_mm, height_mm = settings.size_mm()
        doc.setPageSize(QSizeF(mm_to_px(width_mm), mm_to_px(height_mm)))
        root = doc.rootFrame()
        fmt = root.frameFormat()
        fmt.setTopMargin(mm_to_px(settings.top_mm))
        fmt.setRightMargin(mm_to_px(settings.right_mm))
        fmt.setBottomMargin(mm_to_px(settings.bottom_mm))
        fmt.setLeftMargin(mm_to_px(settings.left_mm))
        root.setFrameFormat(fmt)
        self._line_numbering = settings.line_numbering
        self._line_number_start = settings.line_number_start
        self._line_number_count_by = settings.line_number_count_by
        self._line_number_distance_px = mm_to_px(
            settings.line_number_distance_mm)
        self._sync_height()
        self.viewport().update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        from .formatting import draw_paragraph_borders

        painter = QPainter(self.viewport())
        painter.translate(-self.horizontalScrollBar().value(),
                          -self.verticalScrollBar().value())
        draw_paragraph_borders(painter, self.document())
        painter.end()
        if self._line_numbering == "none":
            return
        painter = QPainter(self.viewport())
        font = QFont(self.document().defaultFont())
        font.setPointSizeF(8.0)
        painter.setFont(font)
        painter.setPen(QColor("#7a8699"))
        number = self._line_number_start
        page_height = max(1.0, self.document().pageSize().height())
        previous_page = 0
        block = self.document().begin()
        while block.isValid():
            layout = block.layout()
            block_top = self.document().documentLayout().blockBoundingRect(
                block).top()
            page = int(block_top / page_height)
            if (self._line_numbering == "restart_page"
                    and page != previous_page):
                number = self._line_number_start
            previous_page = page
            line_count = max(1, layout.lineCount())
            for line_index in range(line_count):
                line = layout.lineAt(line_index) if layout.lineCount() else None
                cursor = QTextCursor(block)
                if line is not None:
                    cursor.setPosition(block.position() + line.textStart())
                rect = self.cursorRect(cursor)
                if ((number - self._line_number_start)
                        % self._line_number_count_by == 0):
                    right = rect.left() - self._line_number_distance_px
                    painter.drawText(
                        0, rect.top() - 2, max(0, int(right)),
                        max(12, rect.height() + 4),
                        Qt.AlignmentFlag.AlignRight
                        | Qt.AlignmentFlag.AlignVCenter,
                        str(number))
                number += 1
            block = block.next()
        painter.end()

    def _selected_blocks(self):
        cursor = self.textCursor()
        doc = self.document()
        start = cursor.selectionStart()
        end = cursor.selectionEnd()
        if cursor.hasSelection() and end > start:
            end -= 1
        return doc.findBlock(start), doc.findBlock(end)

    def apply_style(self, name: str) -> None:
        from .design import DesignSettings, bind_style

        if name not in STYLE_PRESETS:
            raise ValueError(f"Unknown style: {name}")
        bind_style(self.textCursor(), name,
                   getattr(self, "design_settings", DesignSettings()))

    def current_style_name(self) -> str:
        from .design import STYLE_PROPERTY, style_name

        cursor = self.textCursor()
        block_fmt = cursor.blockFormat()
        if block_fmt.hasProperty(STYLE_PROPERTY):
            return style_name(cursor.block())
        char_fmt = cursor.charFormat()
        level = block_fmt.headingLevel()
        for name in STYLE_ORDER:
            preset = STYLE_PRESETS[name]
            if preset.heading_level == level and level > 0:
                return name
        if level > 0:
            return f"Heading {level}"
        size = char_fmt.fontPointSize()
        for name in ("Title", "Subtitle", "Quote", "Normal"):
            preset = STYLE_PRESETS[name]
            if abs(size - preset.size_pt) < 0.5:
                if name == "Quote" and block_fmt.leftMargin() > 1:
                    return name
                if name != "Quote":
                    return name
        return "Normal"

    def _merge_char(self, fmt: QTextCharFormat) -> None:
        cursor = self.textCursor()
        tracker = getattr(self.document(), "folio_review_tracker", None)
        if tracker is not None and not tracker.allows_edit(
                cursor.selectionStart(), cursor.selectionEnd(), formatting=True):
            tracker.edit_blocked.emit("Formatting is restricted in this range.")
            return
        if cursor.hasSelection():
            cursor.beginEditBlock()
            cursor.mergeCharFormat(fmt)
            cursor.endEditBlock()
        else:
            self.mergeCurrentCharFormat(fmt)

    def set_font_family(self, name: str) -> None:
        fmt = QTextCharFormat()
        fmt.setFontFamilies([name])
        self._merge_char(fmt)
        self.setFocus()

    def set_font_size(self, points: float) -> None:
        if points <= 0:
            return
        fmt = QTextCharFormat()
        fmt.setFontPointSize(float(points))
        self._merge_char(fmt)
        self.setFocus()

    def _selected_text_fragments(self):
        cursor = self.textCursor()
        start, end = cursor.selectionStart(), cursor.selectionEnd()
        block = self.document().findBlock(start)
        result = []
        while block.isValid() and block.position() < end:
            iterator = block.begin()
            while not iterator.atEnd():
                fragment = iterator.fragment()
                if fragment.isValid():
                    fmt = fragment.charFormat()
                    lo = max(start, fragment.position())
                    hi = min(end, fragment.position() + fragment.length())
                    if (lo < hi and not fmt.isImageFormat()
                            and fmt.objectType() == QTextFormat.ObjectTypes.NoObject):
                        result.append((lo, hi, QTextCharFormat(fmt)))
                iterator += 1
            block = block.next()
        return result

    def sample_inline_format(self) -> QTextCharFormat:
        source = self.textCursor().charFormat()
        fragments = self._selected_text_fragments()
        if fragments:
            source = fragments[0][2]
        sample = QTextCharFormat()
        sample.setFont(source.font().resolve(self.document().defaultFont()))
        sample.setForeground(source.foreground())
        sample.setBackground(source.background())
        sample.setUnderlineStyle(source.underlineStyle())
        sample.setUnderlineColor(source.underlineColor())
        sample.setVerticalAlignment(source.verticalAlignment())
        sample.setTextOutline(source.textOutline())
        return sample

    def apply_inline_format(self, sample: QTextCharFormat) -> None:
        clean = QTextCharFormat()
        clean.setFont(sample.font())
        clean.setForeground(sample.foreground())
        clean.setBackground(sample.background())
        clean.setUnderlineStyle(sample.underlineStyle())
        clean.setUnderlineColor(sample.underlineColor())
        clean.setVerticalAlignment(sample.verticalAlignment())
        clean.setTextOutline(sample.textOutline())
        cursor = self.textCursor()
        if not cursor.hasSelection():
            self.mergeCurrentCharFormat(clean)
            return
        cursor.beginEditBlock()
        for start, end, _fmt in self._selected_text_fragments():
            edit = QTextCursor(self.document())
            edit.setPosition(start)
            edit.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            edit.mergeCharFormat(clean)
        cursor.endEditBlock()
        self.setFocus()

    def step_font_size(self, direction: int) -> None:
        from .formatting import stepped_font_size

        cursor = self.textCursor()
        default = self.document().defaultFont().pointSizeF()
        if not cursor.hasSelection():
            size = cursor.charFormat().fontPointSize() or default
            self.set_font_size(stepped_font_size(size, direction))
            return
        cursor.beginEditBlock()
        for start, end, current in self._selected_text_fragments():
            size = current.fontPointSize() or default
            fmt = QTextCharFormat()
            fmt.setFontPointSize(stepped_font_size(size, direction))
            edit = QTextCursor(self.document())
            edit.setPosition(start)
            edit.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            edit.mergeCharFormat(fmt)
        cursor.endEditBlock()
        self.setFocus()

    def change_case(self, mode: str) -> None:
        if mode not in ("sentence", "lower", "upper", "title"):
            raise ValueError("Unknown case transformation")
        cursor = self.textCursor()
        if not cursor.hasSelection():
            return
        start, end = cursor.selectionStart(), cursor.selectionEnd()
        selected = cursor.selectedText()
        lowered = selected.lower()
        lower_offset = 0
        position = start
        sentence_start = True
        word_start = True
        replacements = []
        fragments = self._selected_text_fragments()
        fragment_index = 0
        for char in selected:
            width = len(char.encode("utf-16-le")) // 2
            lower_width = len(char.lower())
            lower = lowered[lower_offset:lower_offset + lower_width]
            lower_offset += lower_width
            replacement = lower
            if mode == "upper":
                replacement = char.upper()
            elif mode == "title" and word_start:
                replacement = char.title()
            elif mode == "sentence" and sentence_start and char.isalpha():
                replacement = char.upper()
            if char.isalpha():
                sentence_start = False
            elif char in ".!?\u2029\n":
                sentence_start = True
            word_start = not char.isalpha()
            while fragment_index < len(fragments) and fragments[fragment_index][1] <= position:
                fragment_index += 1
            if (replacement != char and fragment_index < len(fragments)
                    and fragments[fragment_index][0] <= position
                    and position + width <= fragments[fragment_index][1]):
                replacements.append((position, width, replacement,
                                     fragments[fragment_index][2]))
            position += width
        delta = 0
        for index, (position, width, replacement, fmt) in enumerate(reversed(replacements)):
            edit = QTextCursor(self.document())
            edit.setPosition(position)
            edit.setPosition(position + width, QTextCursor.MoveMode.KeepAnchor)
            if index:
                edit.joinPreviousEditBlock()
            else:
                edit.beginEditBlock()
            edit.insertText(replacement, fmt)
            edit.endEditBlock()
            delta += len(replacement.encode("utf-16-le")) // 2 - width
        cursor.setPosition(start)
        cursor.setPosition(end + delta, QTextCursor.MoveMode.KeepAnchor)
        self.setTextCursor(cursor)
        self.setFocus()

    def set_foreground(self, color: QColor) -> None:
        fmt = QTextCharFormat()
        fmt.setForeground(color)
        self._merge_char(fmt)
        self.setFocus()

    def set_highlight(self, color: QColor | None) -> None:
        fmt = QTextCharFormat()
        if color is None:
            fmt.setBackground(QBrush(Qt.BrushStyle.NoBrush))
        else:
            fmt.setBackground(color)
        self._merge_char(fmt)
        self.setFocus()

    def toggle_bold(self) -> None:
        fmt = QTextCharFormat()
        current = self.textCursor().charFormat().fontWeight()
        fmt.setFontWeight(
            QFont.Weight.Normal if current >= QFont.Weight.Bold else QFont.Weight.Bold
        )
        self._merge_char(fmt)

    def toggle_italic(self) -> None:
        fmt = QTextCharFormat()
        fmt.setFontItalic(not self.textCursor().charFormat().fontItalic())
        self._merge_char(fmt)

    def toggle_underline(self) -> None:
        fmt = QTextCharFormat()
        fmt.setFontUnderline(not self.textCursor().charFormat().fontUnderline())
        self._merge_char(fmt)

    def toggle_strikethrough(self) -> None:
        fmt = QTextCharFormat()
        fmt.setFontStrikeOut(not self.textCursor().charFormat().fontStrikeOut())
        self._merge_char(fmt)

    def toggle_text_shadow(self) -> None:
        """Toggle a subtle text effect that exports as Word's shadow flag."""
        current = self.textCursor().charFormat().textOutline()
        fmt = QTextCharFormat()
        if current.style() == Qt.PenStyle.NoPen:
            fmt.setTextOutline(QPen(QColor("#78869a"), 0.65))
        else:
            fmt.setTextOutline(QPen(Qt.PenStyle.NoPen))
        self._merge_char(fmt)
        self.setFocus()

    def toggle_subscript(self) -> None:
        self._toggle_vertical_alignment(QTextCharFormat.VerticalAlignment.AlignSubScript)

    def toggle_superscript(self) -> None:
        self._toggle_vertical_alignment(QTextCharFormat.VerticalAlignment.AlignSuperScript)

    def _toggle_vertical_alignment(self, alignment) -> None:
        fmt = QTextCharFormat()
        current = self.textCursor().charFormat().verticalAlignment()
        fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignNormal
                                 if current == alignment else alignment)
        self._merge_char(fmt)
        self.setFocus()

    def set_alignment(self, alignment: Qt.AlignmentFlag) -> None:
        self.setAlignment(alignment)
        self.setFocus()

    def clear_formatting(self) -> None:
        cursor = self.textCursor()
        cursor.beginEditBlock()
        doc = self.document()
        preset = STYLE_PRESETS["Normal"]
        sel_start, sel_end = cursor.selectionStart(), cursor.selectionEnd()
        char_fmt = QTextCharFormat()
        char_fmt.setFontFamilies(["Calibri"])
        char_fmt.setFontPointSize(preset.size_pt)
        char_fmt.setFontWeight(QFont.Weight.Normal)
        char_fmt.setFontItalic(False)
        char_fmt.setFontUnderline(False)
        char_fmt.setFontStrikeOut(False)
        char_fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignNormal)
        char_fmt.setTextOutline(QPen(Qt.PenStyle.NoPen))
        char_fmt.setForeground(QColor(preset.color))
        char_fmt.setBackground(QBrush(Qt.BrushStyle.NoBrush))
        char_fmt.setAnchor(False)
        char_fmt.setAnchorHref("")
        char_fmt.setToolTip("")
        if cursor.hasSelection():
            start, end = self._selected_blocks()
            block = start
            while block.isValid():
                iterator = block.begin()
                while not iterator.atEnd():
                    fragment = iterator.fragment()
                    if fragment.isValid():
                        frag_fmt = fragment.charFormat()
                        if not frag_fmt.isImageFormat() and frag_fmt.objectType() == QTextFormat.ObjectTypes.NoObject:
                            f_start = fragment.position()
                            f_end = f_start + fragment.length()
                            lo = max(f_start, sel_start)
                            hi = min(f_end, sel_end)
                            if lo < hi:
                                edit = QTextCursor(doc)
                                edit.setPosition(lo)
                                edit.setPosition(hi, QTextCursor.MoveMode.KeepAnchor)
                                edit.setCharFormat(char_fmt)
                    iterator += 1
                if block == end:
                    break
                block = block.next()
            block = start
            while block.isValid():
                edit = QTextCursor(block)
                fmt = QTextBlockFormat()
                fmt.setBottomMargin(pt_to_px(preset.after_pt))
                edit.setBlockFormat(fmt)
                if block == end:
                    break
                block = block.next()
        else:
            self.setCurrentCharFormat(char_fmt)
        cursor.endEditBlock()

    def set_list(self, ordered: bool) -> None:
        cursor = self.textCursor()
        style = (QTextListFormat.Style.ListDecimal if ordered
                 else QTextListFormat.Style.ListDisc)
        start, end = self._selected_blocks()
        all_listed = True
        block = start
        while block.isValid():
            lst = block.textList()
            if lst is None:
                all_listed = False
                break
            lst_style = lst.format().style()
            is_decimal = lst_style in (
                QTextListFormat.Style.ListDecimal,
                QTextListFormat.Style.ListLowerAlpha,
                QTextListFormat.Style.ListUpperAlpha,
                QTextListFormat.Style.ListLowerRoman,
                QTextListFormat.Style.ListUpperRoman,
            )
            if is_decimal != ordered:
                all_listed = False
                break
            if block == end:
                break
            block = block.next()
        cursor.beginEditBlock()
        if all_listed and start.isValid():
            block = start
            while block.isValid():
                edit = QTextCursor(block)
                fmt = edit.blockFormat()
                fmt.setObjectIndex(-1)
                edit.setBlockFormat(fmt)
                if block == end:
                    break
                block = block.next()
        else:
            fmt = QTextListFormat()
            fmt.setStyle(style)
            fmt.setIndent(1)
            block = start
            while block.isValid():
                self._join_list(block, fmt)
                if block == end:
                    break
                block = block.next()
        cursor.endEditBlock()
        self.setFocus()

    def _join_list(self, block, fmt: QTextListFormat) -> None:
        from .formatting import MULTILEVEL_LIST_PROPERTY

        level = fmt.indent()
        neighbors = (block.previous(), block.next())
        target = None
        for index, neighbor in enumerate(neighbors):
            while neighbor.isValid():
                candidate = neighbor.textList()
                if candidate is None or candidate.format().indent() < level:
                    break
                other = candidate.format()
                if other.indent() == level:
                    if (other.style() == fmt.style()
                            and bool(other.property(MULTILEVEL_LIST_PROPERTY))
                            == bool(fmt.property(MULTILEVEL_LIST_PROPERTY))):
                        target = candidate
                    break
                neighbor = neighbor.previous() if index == 0 else neighbor.next()
            if target is not None:
                break
        if target is None:
            QTextCursor(block).createList(fmt)
        else:
            target.add(block)

    def set_multilevel_list(self) -> None:
        from .formatting import MULTILEVEL_LIST_PROPERTY

        cursor = self.textCursor()
        start, end = self._selected_blocks()
        cursor.beginEditBlock()
        block = start
        while block.isValid():
            fmt = QTextListFormat()
            level = block.textList().format().indent() if block.textList() else 1
            fmt.setIndent(level)
            fmt.setStyle(self._multilevel_style(level))
            fmt.setProperty(MULTILEVEL_LIST_PROPERTY, True)
            self._join_list(block, fmt)
            if block == end:
                break
            block = block.next()
        cursor.endEditBlock()
        self.setFocus()

    @staticmethod
    def _multilevel_style(level: int):
        return (QTextListFormat.Style.ListDecimal,
                QTextListFormat.Style.ListLowerAlpha,
                QTextListFormat.Style.ListLowerRoman)[(level - 1) % 3]

    def indent_paragraph(self, delta: int) -> None:
        from .formatting import MULTILEVEL_LIST_PROPERTY

        cursor = self.textCursor()
        start, end = self._selected_blocks()
        changes = []
        block = start
        while block.isValid():
            lst = block.textList()
            changes.append((block, QTextListFormat(lst.format()) if lst else None))
            if block == end:
                break
            block = block.next()
        cursor.beginEditBlock()
        for block, list_fmt in changes:
            if list_fmt is not None:
                edit = QTextCursor(block)
                level = min(9, list_fmt.indent() + delta)
                if level < 1:
                    fmt = edit.blockFormat()
                    fmt.setObjectIndex(-1)
                    fmt.setIndent(0)
                    edit.setBlockFormat(fmt)
                else:
                    list_fmt.setIndent(level)
                    if list_fmt.property(MULTILEVEL_LIST_PROPERTY):
                        list_fmt.setStyle(self._multilevel_style(level))
                    self._join_list(block, list_fmt)
            else:
                edit = QTextCursor(block)
                fmt = edit.blockFormat()
                fmt.setLeftMargin(max(0.0, fmt.leftMargin() + delta * 24.0))
                edit.setBlockFormat(fmt)
        cursor.endEditBlock()
        self.setFocus()

    def set_paragraph_shading(self, color: QColor | None) -> None:
        cursor = self.textCursor()
        start, end = self._selected_blocks()
        cursor.beginEditBlock()
        block = start
        while block.isValid():
            edit = QTextCursor(block)
            fmt = edit.blockFormat()
            if color is None:
                fmt.clearBackground()
            else:
                fmt.setBackground(color)
            edit.setBlockFormat(fmt)
            if block == end:
                break
            block = block.next()
        cursor.endEditBlock()
        self.viewport().update()

    def set_paragraph_border(self, kind: str, color: QColor | str = "#263244",
                             width: float = 1.0) -> None:
        from .formatting import PARAGRAPH_BORDER_PROPERTY, border_spec

        spec = border_spec(kind, QColor(color).name(), width)
        cursor = self.textCursor()
        start, end = self._selected_blocks()
        cursor.beginEditBlock()
        block = start
        while block.isValid():
            edit = QTextCursor(block)
            fmt = edit.blockFormat()
            if spec["sides"]:
                fmt.setProperty(PARAGRAPH_BORDER_PROPERTY, json.dumps(spec))
            else:
                fmt.clearProperty(PARAGRAPH_BORDER_PROPERTY)
            edit.setBlockFormat(fmt)
            if block == end:
                break
            block = block.next()
        cursor.endEditBlock()
        self.viewport().update()

    def set_paragraph(self, before_pt: float, after_pt: float, line_multiple: float,
                      left_mm: float, right_mm: float, first_mm: float) -> None:
        cursor = self.textCursor()
        start, end = self._selected_blocks()
        cursor.beginEditBlock()
        block = start
        while block.isValid():
            edit = QTextCursor(block)
            fmt = edit.blockFormat()
            fmt.setTopMargin(pt_to_px(before_pt))
            fmt.setBottomMargin(pt_to_px(after_pt))
            fmt.setLeftMargin(mm_to_px(left_mm))
            fmt.setRightMargin(mm_to_px(right_mm))
            fmt.setTextIndent(mm_to_px(first_mm))
            if line_multiple > 0:
                fmt.setLineHeight(line_multiple * 100.0,
                                  QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
            else:
                fmt.setLineHeight(0, QTextBlockFormat.LineHeightTypes.SingleHeight.value)
            edit.setBlockFormat(fmt)
            if block == end:
                break
            block = block.next()
        cursor.endEditBlock()
        self.setFocus()

    def set_line_spacing(self, multiple: float) -> None:
        cursor = self.textCursor()
        start, end = self._selected_blocks()
        cursor.beginEditBlock()
        block = start
        while block.isValid():
            edit = QTextCursor(block)
            fmt = edit.blockFormat()
            fmt.setLineHeight(multiple * 100.0,
                              QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
            edit.setBlockFormat(fmt)
            if block == end:
                break
            block = block.next()
        cursor.endEditBlock()
        self.setFocus()

    def insert_image(self, image: QImage, width: float | None = None) -> None:
        if image.isNull():
            raise ValueError("Cannot insert an empty image")
        uri = _image_to_data_uri(image)
        doc = self.document()
        doc.addResource(QTextDocument.ResourceType.ImageResource, QUrl(uri), image)
        max_width = self.content_width_px()
        if width is None:
            width = min(float(image.width()), max_width)
        width = max(8.0, min(float(width), max_width))
        height = width * image.height() / max(1, image.width())
        fmt = QTextImageFormat()
        fmt.setName(uri)
        fmt.setWidth(width)
        fmt.setHeight(height)
        cursor = self.textCursor()
        cursor.insertImage(fmt)
        self.setFocus()

    def selected_image_format(self) -> QTextImageFormat | None:
        cursor = self.textCursor()
        if not cursor.hasSelection():
            block = cursor.block()
            it = block.begin()
            while not it.atEnd():
                frag = it.fragment()
                if frag.isValid() and frag.charFormat().isImageFormat():
                    fmt = QTextImageFormat(frag.charFormat().toImageFormat())
                    if frag.position() <= cursor.position() <= frag.position() + frag.length():
                        return fmt
                it += 1
            return None
        fmt = cursor.charFormat()
        if fmt.isImageFormat():
            return QTextImageFormat(fmt.toImageFormat())
        return None

    def resize_selected_image(self, width: float) -> bool:
        cursor = self.textCursor()
        target = None
        doc = self.document()
        if cursor.hasSelection():
            block = doc.findBlock(cursor.selectionStart())
            it = block.begin()
            while not it.atEnd():
                frag = it.fragment()
                if frag.isValid() and frag.charFormat().isImageFormat():
                    if frag.position() >= cursor.selectionStart():
                        target = frag
                        break
                it += 1
        else:
            block = cursor.block()
            it = block.begin()
            while not it.atEnd():
                frag = it.fragment()
                if frag.isValid() and frag.charFormat().isImageFormat():
                    if frag.position() <= cursor.position() <= frag.position() + frag.length():
                        target = frag
                        break
                it += 1
        if target is None:
            return False
        fmt = QTextImageFormat(target.charFormat().toImageFormat())
        width = max(8.0, min(float(width), self.content_width_px()))
        old_w = fmt.width()
        old_h = fmt.height()
        fmt.setWidth(width)
        if old_w > 0 and old_h > 0:
            fmt.setHeight(old_h * width / old_w)
        edit = QTextCursor(doc)
        edit.setPosition(target.position())
        edit.setPosition(target.position() + target.length(),
                         QTextCursor.MoveMode.KeepAnchor)
        edit.beginEditBlock()
        edit.setCharFormat(fmt)
        edit.endEditBlock()
        return True

    def insert_image_file(self, path, width: float | None = None) -> None:
        image = QImage(str(path))
        if image.isNull():
            raise ValueError(f"Could not read image file: {path}")
        self.insert_image(image, width)

    def insert_table(self, rows: int, columns: int):
        rows = max(1, int(rows))
        columns = max(1, int(columns))
        fmt = QTextTableFormat()
        fmt.setCellSpacing(0.0)
        fmt.setCellPadding(6.0)
        fmt.setBorder(0.75)
        fmt.setBorderBrush(QColor("#b8c4d4"))
        fmt.setBorderStyle(QTextFrameFormat.BorderStyle.BorderStyle_Solid)
        fmt.setHeaderRowCount(1)
        fmt.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
        fmt.setAlignment(Qt.AlignmentFlag.AlignLeft)
        cursor = self.textCursor()
        cursor.beginEditBlock()
        table = cursor.insertTable(rows, columns, fmt)
        header = QColor("#dce8f9")
        for col in range(columns):
            cell = table.cellAt(0, col)
            cell_fmt = cell.format()
            cell_fmt.setBackground(header)
            cell.setFormat(cell_fmt)
        cursor.endEditBlock()
        self.setFocus()
        return table

    def insert_cover_page(self, title: str, subtitle: str = "",
                          style: str = "ViewMaster") -> None:
        palettes = {
            "ViewMaster": ("#17365d", "#dce8f9", "#245c9f"),
            "Facet": ("#134e4a", "#ccfbf1", "#0f766e"),
            "Ion": ("#4c1d95", "#ede9fe", "#7c3aed"),
        }
        dark, pale, accent = palettes.get(style, palettes["ViewMaster"])
        safe_title = html.escape(title.strip() or "Document title")
        safe_subtitle = html.escape(subtitle.strip())
        cover = (
            f'<div style="background-color:{pale};padding:48px;">'
            f'<p style="font-size:11pt;color:{accent};">{html.escape(style)}</p>'
            f'<p style="font-size:34pt;font-weight:600;color:{dark};'
            f'margin-top:72px;">{safe_title}</p>'
            f'<p style="font-size:16pt;color:{accent};margin-top:18px;">'
            f'{safe_subtitle}</p></div>')
        cursor = QTextCursor(self.document())
        cursor.setPosition(0)
        cursor.beginEditBlock()
        cursor.insertHtml(cover)
        cursor.insertBlock()
        fmt = cursor.blockFormat()
        fmt.setPageBreakPolicy(
            QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
        cursor.setBlockFormat(fmt)
        cursor.endEditBlock()
        self.setTextCursor(cursor)
        self.setFocus()

    def current_table(self):
        return self.textCursor().currentTable()

    def table_insert_row(self, below: bool = True) -> bool:
        table = self.current_table()
        if table is None:
            return False
        cursor = self.textCursor()
        cell = table.cellAt(cursor)
        row = cell.row() + (1 if below else 0)
        table.insertRows(row, 1)
        return True

    def table_insert_column(self, right: bool = True) -> bool:
        table = self.current_table()
        if table is None:
            return False
        cursor = self.textCursor()
        cell = table.cellAt(cursor)
        col = cell.column() + (1 if right else 0)
        table.insertColumns(col, 1)
        return True

    def table_delete_row(self) -> bool:
        table = self.current_table()
        if table is None or table.rows() <= 1:
            return False
        cell = table.cellAt(self.textCursor())
        table.removeRows(cell.row(), 1)
        return True

    def table_delete_column(self) -> bool:
        table = self.current_table()
        if table is None or table.columns() <= 1:
            return False
        cell = table.cellAt(self.textCursor())
        table.removeColumns(cell.column(), 1)
        return True

    def table_merge_cells(self) -> bool:
        table = self.current_table()
        if table is None:
            return False
        cursor = self.textCursor()
        if not cursor.hasSelection() or not cursor.hasComplexSelection():
            return False
        table.mergeCells(cursor)
        return True

    def insert_page_break(self) -> None:
        cursor = self.textCursor()
        cursor.beginEditBlock()
        cursor.insertBlock()
        fmt = cursor.blockFormat()
        fmt.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
        cursor.setBlockFormat(fmt)
        cursor.endEditBlock()
        self.setFocus()

    def insert_column_break(self) -> None:
        """Move following content to the next publishing column."""
        cursor = self.textCursor()
        following_format = QTextCharFormat(cursor.charFormat())
        following_format.setAnchor(False)
        following_format.setAnchorHref("")
        cursor.beginEditBlock()
        cursor.insertBlock()
        block_fmt = cursor.blockFormat()
        block_fmt.setPageBreakPolicy(
            QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
        cursor.setBlockFormat(block_fmt)
        marker = QTextCharFormat()
        marker.setAnchor(True)
        marker.setAnchorHref(COLUMN_BREAK_URL)
        marker.setFontPointSize(0.1)
        marker.setForeground(QColor(Qt.GlobalColor.transparent))
        cursor.insertText(COLUMN_BREAK_MARKER, marker)
        cursor.setCharFormat(following_format)
        cursor.endEditBlock()
        self.setTextCursor(cursor)
        self.setFocus()

    def apply_manual_hyphenation(self, hyphenate_caps: bool = False) -> int:
        """Insert discretionary hyphens in long words and return the count."""
        text = self.document().toPlainText()
        candidates: list[int] = []
        vowels = set("aeiouyAEIOUY")
        for match in re.finditer(r"[A-Za-z]{12,}", text):
            word = match.group(0)
            if not hyphenate_caps and word.isupper():
                continue
            middle = len(word) // 2
            choices = sorted(range(4, len(word) - 3),
                             key=lambda value: abs(value - middle))
            split = next((value for value in choices
                          if (word[value - 1] in vowels)
                          != (word[value] in vowels)), middle)
            candidates.append(match.start() + split)
        if not candidates:
            return 0
        cursor = QTextCursor(self.document())
        cursor.beginEditBlock()
        for py_position in reversed(candidates):
            utf16_position = len(
                text[:py_position].encode("utf-16-le")) // 2
            cursor.setPosition(utf16_position)
            cursor.insertText("\u00ad")
        cursor.endEditBlock()
        return len(candidates)

    def insert_horizontal_rule(self) -> None:
        cursor = self.textCursor()
        cursor.insertHtml("<hr/>")
        self.setFocus()

    def insert_link(self, url: str, text: str | None = None) -> None:
        parsed = QUrl(url)
        scheme = parsed.scheme().lower()
        if scheme not in ALLOWED_LINK_SCHEMES:
            raise ValueError(f"Link scheme not allowed: {scheme or '(none)'}")
        cursor = self.textCursor()
        fmt = QTextCharFormat()
        fmt.setAnchor(True)
        fmt.setAnchorHref(url)
        fmt.setForeground(QColor("#185abd"))
        fmt.setFontUnderline(True)
        label = text or url
        cursor.beginEditBlock()
        cursor.insertText(label, fmt)
        cursor.endEditBlock()
        self.setFocus()

    def insert_merge_field(self, field_name: str) -> None:
        name = field_name.strip()
        if not _FIELD_NAME_RE.match(name):
            raise ValueError("Field name must be a single identifier (letters, digits, _)")
        self.textCursor().insertText("{{" + name + "}}")
        self.setFocus()

    def headings(self) -> list[tuple[int, int, str]]:
        doc = self.document()
        result: list[tuple[int, int, str]] = []
        block = doc.begin()
        while block.isValid():
            level = block.blockFormat().headingLevel()
            if level > 0:
                result.append((block.position(), level, block.text()))
            block = block.next()
        return result

    def insertPlainText(self, text: str) -> None:
        from .document_features import detach_completed_node

        detach_completed_node(self)
        super().insertPlainText(text)

    def inputMethodEvent(self, event) -> None:
        from .document_features import detach_completed_node

        detach_completed_node(self)
        super().inputMethodEvent(event)

    def keyPressEvent(self, event: QKeyEvent) -> None:
        if self.isReadOnly():
            super().keyPressEvent(event)
            return
        from .document_features import detach_completed_node

        detach_completed_node(self)
        cursor = self.textCursor()
        key = event.key()
        mods = event.modifiers()
        if key == Qt.Key.Key_Backtab:
            if cursor.block().textList() is not None:
                self.indent_paragraph(-1)
                return
        elif key == Qt.Key.Key_Tab:
            if cursor.block().textList() is not None:
                self.indent_paragraph(1)
                return
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and mods & Qt.KeyboardModifier.ControlModifier:
            self.insert_page_break()
            return
        if (key in (Qt.Key.Key_Return, Qt.Key.Key_Enter)
                and not mods & (Qt.KeyboardModifier.ShiftModifier
                                | Qt.KeyboardModifier.AltModifier
                                | Qt.KeyboardModifier.MetaModifier)
                and not cursor.hasSelection()):
            previous_block = cursor.block()
            had_break = (cursor.blockFormat().pageBreakPolicy()
                         == QTextFormat.PageBreakFlag.PageBreak_AlwaysBefore)
            super().keyPressEvent(event)
            new_cursor = self.textCursor()
            if had_break and new_cursor.block() != previous_block:
                fmt = new_cursor.blockFormat()
                if fmt.pageBreakPolicy() == (QTextFormat.PageBreakFlag
                                             .PageBreak_AlwaysBefore):
                    fmt.setPageBreakPolicy(
                        QTextFormat.PageBreakFlag.PageBreak_Auto)
                    new_cursor.setBlockFormat(fmt)
            return
        if mods & Qt.KeyboardModifier.ControlModifier:
            if key == Qt.Key.Key_1:
                self.apply_style("Heading 1")
                return
            if key == Qt.Key.Key_2:
                self.apply_style("Heading 2")
                return
            if key == Qt.Key.Key_3:
                self.apply_style("Heading 3")
                return
            if key == Qt.Key.Key_0:
                self.apply_style("Normal")
                return
        super().keyPressEvent(event)

    def insert_mime_data(self, source, mode: str = "keep_formatting") -> None:
        if self.isReadOnly():
            return
        from .document_features import detach_completed_node

        detach_completed_node(self)
        if mode not in ("keep_formatting", "text", "image"):
            raise ValueError(f"Unknown paste mode: {mode}")
        if mode == "image" and source.hasImage():
            image = source.imageData()
            if isinstance(image, QImage) and not image.isNull():
                self.insert_image(image)
                return
        if mode == "text" and source.hasText():
            self.textCursor().insertText(source.text())
            return
        if source.hasImage():
            image = source.imageData()
            if isinstance(image, QImage) and not image.isNull():
                self.insert_image(image)
                return
        if mode == "keep_formatting" and source.hasHtml():
            from .document_features import (NODE_MIME_TYPE, capture_node_bindings,
                                            prepare_pasted_nodes, restore_node_bindings)

            doc = SafeDocument()
            doc.setDefaultFont(self.document().defaultFont())
            doc.setHtml(source.html())
            payload = bytes(source.data(NODE_MIME_TYPE)) if source.hasFormat(NODE_MIME_TYPE) else None
            prepare_pasted_nodes(doc, getattr(self.document(), "folio_state", None), payload)
            bindings = capture_node_bindings(doc)
            fragment = QTextDocumentFragment(doc)
            cursor = self.textCursor()
            offset = cursor.selectionStart()
            cursor.beginEditBlock()
            cursor.insertFragment(fragment)
            restore_node_bindings(self.document(), [
                {"start": item["start"] + offset, "end": item["end"] + offset,
                 "keys": item["keys"]} for item in bindings], (offset, cursor.position()))
            cursor.endEditBlock()
            from .document_features import _outside_node_format

            self.setTextCursor(cursor)
            self.setCurrentCharFormat(_outside_node_format(cursor.charFormat()))
            return
        if source.hasText():
            self.textCursor().insertText(source.text())
            return
        super().insertFromMimeData(source)

    def paste_special(self, mode: str = "keep_formatting") -> None:
        clipboard = QGuiApplication.clipboard()
        self.insert_mime_data(clipboard.mimeData(), mode)

    def createMimeDataFromSelection(self):
        from .document_features import (NODE_MIME_TYPE, capture_node_bindings,
                                        clipboard_feature_payload, restore_node_bindings)

        original = super().createMimeDataFromSelection()
        mime = QMimeData()
        for name in original.formats():
            mime.setData(name, original.data(name))
        if original.hasImage():
            mime.setImageData(original.imageData())
        if not self.textCursor().hasSelection():
            return mime
        document = SafeDocument()
        document.setDefaultFont(self.document().defaultFont())
        QTextCursor(document).insertFragment(QTextDocumentFragment(self.textCursor()))
        start, end = self.textCursor().selectionStart(), self.textCursor().selectionEnd()
        bindings = [{"start": max(start, item["start"]) - start,
                     "end": min(end, item["end"]) - start, "keys": item["keys"]}
                    for item in capture_node_bindings(self.document())
                    if item["start"] < end and item["end"] > start]
        restore_node_bindings(document, bindings)
        mime.setHtml(document.toHtml())
        state = getattr(self.document(), "folio_state", None)
        if state is not None:
            mime.setData(NODE_MIME_TYPE, QByteArray(clipboard_feature_payload(document, state)))
        return mime

    def insertFromMimeData(self, source) -> None:
        self.insert_mime_data(source, "keep_formatting")


def fragment_text(cursor: QTextCursor) -> str:
    return cursor.selectedText().replace(" ", "\n").replace(" ", "\n")
