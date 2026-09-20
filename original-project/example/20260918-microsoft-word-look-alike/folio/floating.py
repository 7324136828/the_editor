from __future__ import annotations

import math
from dataclasses import dataclass

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import (
    QFontMetricsF, QImage, QPainter, QTextCursor, QTextDocumentFragment,
    QTextFormat, QTextImageFormat,
)

from .document_features import object_image, object_rect, visible_objects
from .editor import SafeDocument, block_break_kind, mm_to_px


@dataclass
class FlowLine:
    page: int
    column: int
    start: int
    end: int
    rect: QRectF
    document: SafeDocument
    source_y: float = 0.0
    marker: str = ""


class FloatingLayout:
    def __init__(self, document, page_settings, objects, body_height=None):
        page_settings.validate()
        self.source = document
        self.settings = page_settings
        self.objects = list(visible_objects(objects))
        self.lines: list[FlowLine] = []
        self.column_width = mm_to_px(page_settings.column_width_mm())
        self.body_height = body_height or mm_to_px(page_settings.content_height_mm())
        self.margin_left = mm_to_px(page_settings.left_mm)
        self.margin_top = mm_to_px(page_settings.top_mm)
        self.bottom = self.margin_top + self.body_height
        self.gutter = mm_to_px(page_settings.gutter_mm)
        self._page = 0
        self._column = 0
        self._y = self.margin_top
        self._silhouettes = {}
        self._previous_margin = 0.0
        self._device = QImage(16, 16, QImage.Format.Format_ARGB32)
        self._device.setDotsPerMeterX(round(96 / 0.0254))
        self._device.setDotsPerMeterY(round(96 / 0.0254))
        iterator = document.rootFrame().begin()
        while not iterator.atEnd():
            frame = iterator.currentFrame()
            block = iterator.currentBlock()
            if frame is not None:
                self._layout_frame(frame)
            elif block.isValid():
                self._layout_block(block)
            iterator += 1
        self._pages = max(1, self._page + 1,
                          max((obj.get("page", 0) + 1 for obj in self.objects), default=1))

    def page_count(self) -> int:
        return self._pages

    def page_for_position(self, position: int) -> int:
        for line in self.lines:
            if line.start <= position < line.end or line.start == position:
                return line.page
        before = [line for line in self.lines if line.start <= position]
        return before[-1].page if before else 0

    def _advance_column(self, page=False):
        if page:
            self._page += 1
            self._column = 0
        else:
            self._column += 1
            if self._column >= self.settings.columns:
                self._page += 1
                self._column = 0
        self._y = self.margin_top
        self._previous_margin = 0.0

    def _column_x(self):
        return self.margin_left + self._column * (self.column_width + self.gutter)

    def _outline(self, obj):
        key = obj["id"]
        if key in self._silhouettes:
            return self._silhouettes[key]
        image = object_image(obj)
        if image.isNull():
            result = None
        else:
            if image.width() > 512 or image.height() > 512:
                image = image.scaled(512, 512, Qt.AspectRatioMode.KeepAspectRatio,
                                     Qt.TransformationMode.SmoothTransformation)
            rows = []
            for y in range(image.height()):
                opaque = [x for x in range(image.width()) if image.pixelColor(x, y).alpha() > 16]
                rows.append((min(opaque), max(opaque) + 1) if opaque else None)
            result = image.width(), image.height(), rows
        self._silhouettes[key] = result
        return result

    def excluded_intervals(self, page, y, height, padding=6.0):
        intervals = []
        for obj in self.objects:
            if obj.get("page", 0) != page or obj.get("wrap") not in ("square", "tight"):
                continue
            rect = object_rect(obj)
            if y + height <= rect.top() - padding or y >= rect.bottom() + padding:
                continue
            left, right = rect.left(), rect.right()
            if obj.get("wrap") == "tight" and obj["kind"] == "image":
                outline = self._outline(obj)
                if outline is not None:
                    width, image_height, rows = outline
                    first = max(0, math.floor((y - padding - rect.top()) / rect.height() * image_height))
                    last = min(image_height, math.ceil((y + height + padding - rect.top())
                                                       / rect.height() * image_height))
                    occupied = [row for row in rows[first:last] if row]
                    if not occupied:
                        continue
                    left += min(row[0] for row in occupied) / width * rect.width()
                    right = rect.left() + max(row[1] for row in occupied) / width * rect.width()
            intervals.append((left - padding, right + padding))
        merged = []
        for left, right in sorted(intervals):
            if merged and left <= merged[-1][1]:
                merged[-1] = (merged[-1][0], max(right, merged[-1][1]))
            else:
                merged.append((left, right))
        return merged

    def free_intervals(self, page, y, height, left, right):
        intervals = []
        current = left
        for lo, hi in self.excluded_intervals(page, y, height):
            if hi <= current or lo >= right:
                continue
            if lo > current:
                intervals.append((current, min(lo, right)))
            current = max(current, hi)
        if current < right:
            intervals.append((current, right))
        return [(lo, hi) for lo, hi in intervals if hi - lo >= 16]

    def _skip_obstruction(self, height):
        ends = [obj["y"] + obj["height"] + 6 for obj in self.objects
                if obj.get("page", 0) == self._page and obj.get("wrap") in ("square", "tight")
                and obj["y"] + obj["height"] + 6 > self._y]
        next_y = min(ends, default=self._y + max(1, height))
        self._y = max(self._y + max(1, height), next_y)
        if self._y + height > self.bottom:
            self._advance_column()

    def _fragment(self, start, end, width, block_format=None):
        cursor = QTextCursor(self.source)
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        fragment = QTextDocumentFragment(cursor)
        document = SafeDocument()
        document.setDefaultFont(self.source.defaultFont())
        document.setDefaultTextOption(self.source.defaultTextOption())
        document.setDocumentMargin(0)
        document.documentLayout().setPaintDevice(self._device)
        edit = QTextCursor(document)
        edit.insertFragment(fragment)
        if block_format is not None:
            fmt = edit.blockFormat()
            fmt.setAlignment(block_format.alignment())
            fmt.setTopMargin(0)
            fmt.setBottomMargin(0)
            fmt.setLeftMargin(0)
            fmt.setRightMargin(0)
            fmt.setIndent(0)
            fmt.setTextIndent(0)
            fmt.setNonBreakableLines(False)
            fmt.setPageBreakPolicy(QTextFormat.PageBreakFlag.PageBreak_Auto)
            edit.setBlockFormat(fmt)
            if edit.currentList():
                edit.currentList().remove(edit.block())
            edit.setBlockCharFormat(self.source.findBlock(start).charFormat())
        block = document.begin()
        while block.isValid():
            iterator = block.begin()
            while not iterator.atEnd():
                item = iterator.fragment()
                if item.isValid() and item.charFormat().isImageFormat():
                    image_format = QTextImageFormat(item.charFormat().toImageFormat())
                    image_width, image_height = image_format.width(), image_format.height()
                    if image_width > 0 and image_height > 0:
                        factor = min(1, self.column_width / image_width, self.body_height / image_height)
                        if factor < 1:
                            image_format.setWidth(image_width * factor)
                            image_format.setHeight(image_height * factor)
                            image_cursor = QTextCursor(document)
                            image_cursor.setPosition(item.position())
                            image_cursor.setPosition(item.position() + item.length(),
                                                     QTextCursor.MoveMode.KeepAnchor)
                            image_cursor.setCharFormat(image_format)
                iterator += 1
            block = block.next()
        document.setTextWidth(max(1, width))
        document.documentLayout().documentSize()
        return document

    def _row(self, block, position, end, first, height):
        fmt = block.blockFormat()
        list_indent = block.textList().format().indent() if block.textList() else fmt.indent()
        left = self._column_x() + fmt.leftMargin() + list_indent * self.source.indentWidth()
        right = self._column_x() + self.column_width - fmt.rightMargin()
        if first:
            left += fmt.textIndent()
        right = max(left + 16, right)
        for attempt in range(6):
            spans = self.free_intervals(self._page, self._y, height, left, right)
            if block.textDirection() == Qt.LayoutDirection.RightToLeft:
                spans.reverse()
            row = []
            consumed = position
            row_height = height
            for lo, hi in spans:
                probe_end = min(end, consumed + 1024)
                while True:
                    document = self._fragment(consumed, probe_end, hi - lo, fmt)
                    qt_block = document.firstBlock()
                    layout = qt_block.layout()
                    if not layout.lineCount():
                        break
                    line = layout.lineAt(0)
                    if probe_end == end or consumed + line.textLength() < probe_end:
                        break
                    probe_end = min(end, consumed + (probe_end - consumed) * 2)
                if not layout.lineCount():
                    continue
                if line.naturalTextWidth() > hi - lo + 1:
                    continue
                line_height = fmt.lineHeight(line.height(), 1.0)
                line_height = max(1, line.height(), line_height)
                row_height = max(row_height, line_height)
                source_y = document.documentLayout().blockBoundingRect(qt_block).top() + line.y()
                marker = block.textList().itemText(block) if first and not row and block.textList() else ""
                row.append(FlowLine(self._page, self._column, consumed,
                                    min(end, consumed + line.textLength()),
                                    QRectF(lo, self._y, hi - lo, line_height), document, source_y, marker))
                consumed += line.textLength()
                if consumed >= end:
                    break
            if row_height <= height + 0.01:
                return row, consumed, row_height
            height = row_height
        return [], position, height

    def _layout_block(self, block):
        break_kind = block_break_kind(block)
        if break_kind and (self.lines or self._y > self.margin_top):
            self._advance_column(page=break_kind == "page")
        fmt = block.blockFormat()
        self._y += max(0, fmt.topMargin() - self._previous_margin)
        font = block.charFormat().font().resolve(self.source.defaultFont())
        height = max(1, QFontMetricsF(font, self._device).height())
        position, end = block.position(), block.position() + block.length() - 1
        first = True
        while position < end or first:
            if self._y + height > self.bottom and self._y > self.margin_top:
                self._advance_column()
            row, consumed, actual_height = self._row(block, position, end, first, height)
            if actual_height > self.body_height:
                actual_height = self.body_height
            if self._y + actual_height > self.bottom and self._y > self.margin_top:
                self._advance_column()
                height = actual_height
                continue
            if not row:
                self._skip_obstruction(actual_height)
                continue
            self.lines.extend(row)
            self._y += actual_height
            first = False
            if consumed <= position:
                break
            position = consumed
        self._y += fmt.bottomMargin()
        self._previous_margin = fmt.bottomMargin()
        if fmt.pageBreakPolicy() & QTextFormat.PageBreakFlag.PageBreak_AlwaysAfter:
            self._advance_column(page=True)

    def _layout_frame(self, frame):
        start = max(0, frame.firstPosition() - 1)
        end = min(self.source.characterCount() - 1, frame.lastPosition() + 1)
        document = self._fragment(start, end, self.column_width)
        height = document.documentLayout().documentSize().height()
        source_y = 0.0
        while source_y < height:
            if self._y >= self.bottom - 1:
                self._advance_column()
            x = self._column_x()
            available = self.bottom - self._y
            band = min(height - source_y, available)
            spans = self.free_intervals(self._page, self._y, band, x, x + self.column_width)
            if not spans or max(hi - lo for lo, hi in spans) < self.column_width - 0.1:
                self._skip_obstruction(min(20, band))
                continue
            self.lines.append(FlowLine(self._page, self._column, start, end,
                                        QRectF(x, self._y, self.column_width, band), document, source_y))
            self._y += band
            source_y += band
        self._previous_margin = 0

    def draw_page(self, painter: QPainter, page_index: int) -> None:
        if not 0 <= page_index < self.page_count():
            raise IndexError("Page index out of range")
        for line in self.lines:
            if line.page != page_index:
                continue
            painter.save()
            painter.setClipRect(line.rect, Qt.ClipOperation.IntersectClip)
            painter.translate(line.rect.left(), line.rect.top() - line.source_y)
            line.document.drawContents(painter, QRectF(0, line.source_y,
                                                       line.rect.width(), line.rect.height()).adjusted(
                                                           0, 0.01, 0, -0.01))
            painter.restore()
            if line.marker:
                painter.save()
                font = self.source.findBlock(line.start).charFormat().font().resolve(self.source.defaultFont())
                painter.setFont(font)
                painter.drawText(line.rect.adjusted(-40, 0, -line.rect.width() - 7, 0),
                                 Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                                 line.marker)
                painter.restore()
