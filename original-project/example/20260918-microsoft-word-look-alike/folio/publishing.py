from __future__ import annotations

import bisect
import copy
import math
import os
import tempfile
from pathlib import Path

from PySide6.QtCore import QMarginsF, QRectF, QSizeF, Qt
from PySide6.QtGui import (
    QColor,
    QFont,
    QImage,
    QPageLayout,
    QPageSize,
    QPainter,
    QPdfWriter,
    QTextCursor,
    QTextDocumentFragment,
    QTextFormat,
    QTextImageFormat,
    QTextListFormat,
)

from .editor import SafeDocument, block_break_kind, mm_to_px
from .models import PAPER_SIZES_MM, PageSettings
from .design import DesignSettings, paint_page_decoration
from .document_features import paint_objects, section_ranges, visible_objects
from .hyphenation import apply_dictionary, layout_position

LAYOUT_DPI = 96


def page_layout_for(settings: PageSettings) -> QPageLayout:
    width_mm, height_mm = PAPER_SIZES_MM[settings.paper]
    orientation = (QPageLayout.Orientation.Landscape if settings.landscape
                   else QPageLayout.Orientation.Portrait)
    return QPageLayout(
        QPageSize(QSizeF(width_mm, height_mm), QPageSize.Unit.Millimeter),
        orientation,
        QMarginsF(0, 0, 0, 0),
        QPageLayout.Unit.Millimeter,
    )


class Publication:
    def __init__(self, document, settings: PageSettings, title: str, state=None,
                 _body_height=None, _compose=True, _balance_tail=True, _note_numbers=None):
        settings.validate()
        self.settings = settings
        self.title = title
        self.state = state if state is not None else getattr(document, "folio_state", None)
        self.design = DesignSettings.from_dict(self.state.design if self.state else {})
        self.objects = self.state.features["objects"] if self.state else []
        self._sections = []
        self._physical_settings = []
        self._flow = None
        self._tail = None
        self._inserted_hyphens = []
        from .references import note_entries

        notes = note_entries(document, self.state, "footnote") if self.state else []
        self._note_numbers = _note_numbers or {key: number for _, key, number, _ in notes}
        self._footnotes = {}
        self._footnote_height = 0.0
        self.document = SafeDocument()
        self.document.setDefaultFont(document.defaultFont())
        self.document.setHtml(document.toHtml())
        self._layout_device = QImage(16, 16, QImage.Format.Format_ARGB32)
        self._layout_device.setDotsPerMeterX(int(LAYOUT_DPI / 0.0254))
        self._layout_device.setDotsPerMeterY(int(LAYOUT_DPI / 0.0254))
        self.document.documentLayout().setPaintDevice(self._layout_device)
        self.column_w = mm_to_px(settings.column_width_mm())
        self.body_h = math.floor((_body_height or mm_to_px(settings.content_height_mm())) * 64) / 64
        if notes and not (self.state and _compose and self.state.features["sections"]):
            self._footnote_height = max(48.0, self.body_h * 0.25)
            self.body_h -= self._footnote_height
        self.gutter = mm_to_px(settings.gutter_mm)
        self.margin_left = mm_to_px(settings.left_mm)
        self.margin_top = mm_to_px(settings.top_mm)
        self.margin_right = mm_to_px(settings.right_mm)
        self.margin_bottom = mm_to_px(settings.bottom_mm)
        width_mm, height_mm = settings.size_mm()
        self.page_w = mm_to_px(width_mm)
        self.page_h = mm_to_px(height_mm)
        root = self.document.rootFrame()
        fmt = root.frameFormat()
        fmt.setTopMargin(0.0)
        fmt.setRightMargin(0.0)
        fmt.setBottomMargin(0.0)
        fmt.setLeftMargin(0.0)
        root.setFrameFormat(fmt)
        if settings.hyphenation == "automatic":
            self._inserted_hyphens = apply_dictionary(self.document, settings.hyphenate_caps)
        self.document.setPageSize(QSizeF(self.column_w, self.body_h))
        self._bound_oversized_images()
        wrapping = any(obj.get("wrap") in ("square", "tight") for obj in visible_objects(self.objects))
        if settings.columns > 1 and settings.balance_columns and not wrapping and not notes:
            self._balance_columns()
        if self._inserted_hyphens and settings.consecutive_hyphen_limit:
            self._enforce_hyphen_limit()
        self._break_pages = self._find_break_pages()
        self._slots: list[tuple[int, int]] = []
        self._map_pages()
        self._numbered_lines = self._map_line_numbers()
        if self.state and _compose and self.state.features["sections"]:
            self._compose_sections(document)
        elif wrapping:
            from .floating import FloatingLayout
            self._flow = FloatingLayout(self.document, settings, self.objects, self.body_h)
        elif _balance_tail and settings.columns > 1 and settings.balance_columns and not notes:
            self._balance_last_page()
        self._numbered_lines = self._map_line_numbers()
        if notes and not self._sections:
            self._layout_footnotes(notes)

    def _layout_footnotes(self, notes):
        page = 0
        used = 0.0
        available = self._footnote_height - 18.0
        for position, key, number, node in notes:
            anchor_page = self.page_for_position(position) - 1
            if anchor_page > page:
                page, used = anchor_page, 0.0
            note = SafeDocument()
            font = QFont("Calibri")
            font.setPixelSize(12)
            note.setDefaultFont(font)
            note.setDocumentMargin(0)
            note.setPlainText(f"{self._note_numbers.get(key, number)}. {node['text']}")
            note.setTextWidth(self.page_w - self.margin_left - self.margin_right)
            note.documentLayout().documentSize()
            block = note.begin()
            while block.isValid():
                rect = note.documentLayout().blockBoundingRect(block)
                for index in range(block.layout().lineCount()):
                    line = block.layout().lineAt(index)
                    height = line.height()
                    if used + height > available and used:
                        page, used = page + 1, 0.0
                    self._footnotes.setdefault(page, []).append((note, rect.top() + line.y(), height, used))
                    used += height
                block = block.next()
            used += 4.0

    def _paint_footnotes(self, painter, physical_index):
        entries = self._footnotes.get(physical_index, [])
        if not entries:
            return
        y = self.margin_top + self.body_h
        width = self.page_w - self.margin_left - self.margin_right
        painter.save()
        painter.setPen(QColor("#64748b"))
        painter.drawLine(self.margin_left, y + 4, self.margin_left + min(width / 3, 160), y + 4)
        for document, source_y, height, offset in entries:
            painter.save()
            painter.setClipRect(QRectF(self.margin_left, y + 18 + offset, width, height))
            painter.translate(self.margin_left, y + 18 + offset - source_y)
            document.drawContents(painter, QRectF(0, source_y, width, height))
            painter.restore()
        painter.restore()

    def _source_position(self, position):
        inserted = [value + index for index, value in enumerate(self._inserted_hyphens)]
        return position - bisect.bisect_left(inserted, position)

    def _enforce_hyphen_limit(self):
        limit = self.settings.consecutive_hyphen_limit
        while self._inserted_hyphens:
            self.document.documentLayout().documentSize()
            actual = [position + index for index, position in enumerate(self._inserted_hyphens)]
            removable = set(actual)
            remove = []
            consecutive = 0
            block = self.document.begin()
            while block.isValid():
                layout = block.layout()
                for index in range(layout.lineCount()):
                    line = layout.lineAt(index)
                    last = block.position() + line.textStart() + line.textLength() - 1
                    if (index < layout.lineCount() - 1 and last in removable
                            and self.document.characterAt(last) == "\u00ad"):
                        consecutive += 1
                        if consecutive > limit:
                            remove.append(last)
                    else:
                        consecutive = 0
                block = block.next()
            if not remove:
                return
            edit = QTextCursor(self.document)
            edit.beginEditBlock()
            for position in reversed(remove):
                edit.setPosition(position)
                edit.deleteChar()
            edit.endEditBlock()
            removed = set(remove)
            self._inserted_hyphens = [position for position, actual_position
                                     in zip(self._inserted_hyphens, actual)
                                     if actual_position not in removed]

    def _balance_columns(self):
        logical = self.document.pageCount()
        block = self.document.begin()
        while block.isValid():
            if block_break_kind(block):
                return
            block = block.next()
        if logical > self.settings.columns:
            return
        low, high = 1.0, self.body_h
        for _ in range(18):
            middle = (low + high) / 2
            self.document.setPageSize(QSizeF(self.column_w, middle))
            if self.document.pageCount() <= self.settings.columns:
                high = middle
            else:
                low = middle
        self.body_h = math.floor(min(self.body_h, high + 1) * 64) / 64
        self.document.setPageSize(QSizeF(self.column_w, self.body_h))

    def _balance_last_page(self):
        if self.page_count() <= 1 or self._find_break_pages():
            return
        block = self.document.begin()
        while block.isValid():
            if block_break_kind(block):
                return
            block = block.next()
        physical = max(page for page, slot in self._slots)
        if physical == 0:
            return
        source_start = self.first_position_on_page(physical)
        start = layout_position(source_start, self._inserted_hyphens)
        cursor = QTextCursor(self.document)
        cursor.setPosition(start)
        if cursor.currentTable():
            return
        fragment = self._slice(self.document, start, self.document.characterCount() - 1)
        settings = copy.copy(self.settings)
        settings.hyphenation = "none"
        child = Publication(fragment, settings, self.title, _body_height=self.body_h,
                            _compose=False, _balance_tail=False)
        if child.page_count() == 1:
            self._tail = (start, physical, child)

    @staticmethod
    def _slice(document, start, end):
        cursor = QTextCursor(document)
        cursor.setPosition(start)
        cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
        result = SafeDocument()
        result.setDefaultFont(document.defaultFont())
        QTextCursor(result).insertFragment(QTextDocumentFragment(cursor))
        if start != document.findBlock(start).position():
            edit = QTextCursor(result)
            fmt = edit.blockFormat()
            fmt.setTopMargin(0)
            fmt.setTextIndent(0)
            edit.setBlockFormat(fmt)
        return result

    def used_height(self, physical=None):
        physical = self.body_page_count() - 1 if physical is None else physical
        if self._tail and physical == self._tail[1]:
            return self._tail[2].used_height()
        if self._flow:
            return max((line.rect.bottom() - self.margin_top for line in self._flow.lines
                        if line.page == physical), default=0)
        height = 0.0
        layout = self.document.documentLayout()
        block = self.document.begin()
        while block.isValid():
            rect = layout.blockBoundingRect(block)
            for index in range(block.layout().lineCount()):
                line = block.layout().lineAt(index)
                y = rect.top() + line.y()
                logical = min(int(y / self.body_h + 0.00001), len(self._slots) - 1)
                if logical >= 0 and self._slots[logical][0] == physical:
                    height = max(height, y - logical * self.body_h + line.height()
                                 + block.blockFormat().bottomMargin())
            block = block.next()
        return height

    def _compose_sections(self, source):
        previous = None
        used = 0.0
        for start, end, settings, kind in section_ranges(source, self.state):
            fragment = self._slice(source, start, end)
            offset = 0.0
            first_page = len(self._physical_settings)
            geometry = lambda s: (s.size_mm(), s.top_mm, s.bottom_mm, s.left_mm, s.right_mm)
            from .references import note_entries

            section_notes = note_entries(fragment, self.state, "footnote")
            preceding_notes = self._sections and self._sections[-1][4]._footnotes
            if (kind == "continuous" and previous and not section_notes and not preceding_notes and geometry(previous) == geometry(settings)
                    and used + 30 < mm_to_px(settings.content_height_mm())):
                first_page -= 1
                offset = used
            elif kind in ("even", "odd"):
                parity = 0 if kind == "even" else 1
                if (first_page + 1) % 2 != parity:
                    self._physical_settings.append(previous or settings)
                    first_page += 1
            remaining = mm_to_px(settings.content_height_mm()) - offset
            local_state = copy.copy(self.state)
            local_state.features = {"nodes": self.state.features["nodes"], "sections": [],
                                    "objects": self._local_objects(first_page, offset)}
            child = Publication(fragment, settings, self.title, local_state,
                                _body_height=remaining, _compose=False, _note_numbers=self._note_numbers)
            if offset and child.body_page_count() > 1:
                split = child.first_position_on_page(1)
                if split > 0:
                    first = self._slice(fragment, 0, split)
                    child = Publication(first, settings, self.title, local_state,
                                        _body_height=remaining, _compose=False, _note_numbers=self._note_numbers)
                    self._sections.append((start, start + split, first_page, offset, child))
                    start += split
                    fragment = self._slice(source, start, end)
                    first_page = len(self._physical_settings)
                else:
                    first_page = len(self._physical_settings)
                offset = 0.0
                local_state = copy.copy(local_state)
                local_state.features = {**local_state.features,
                                        "objects": self._local_objects(first_page, offset)}
                child = Publication(fragment, settings, self.title, local_state, _compose=False,
                                    _note_numbers=self._note_numbers)
            self._sections.append((start, end, first_page, offset, child))
            child_pages = max(child.body_page_count(), max(child._footnotes, default=-1) + 1)
            for _ in range(first_page + child_pages - len(self._physical_settings)):
                self._physical_settings.append(settings)
            used = offset + child.used_height()
            previous = settings

    def _local_objects(self, first_page, offset):
        objects = []
        for original in self.objects:
            if original.get("page", 0) < first_page:
                continue
            obj = copy.copy(original)
            obj["page"] = original.get("page", 0) - first_page
            if obj["page"] == 0:
                obj["y"] -= offset
            objects.append(obj)
        parents = {obj.get("group") for obj in objects}
        for original in self.objects:
            if original["kind"] == "group" and original["id"] in parents:
                if not any(obj["id"] == original["id"] for obj in objects):
                    objects.append(copy.copy(original))
        return objects

    def first_position_on_page(self, physical):
        if self._flow:
            return min((self._source_position(line.start) for line in self._flow.lines
                        if line.page >= physical), default=self.document.characterCount() - 1)
        if self._tail and physical == self._tail[1]:
            return self._source_position(self._tail[0])
        layout = self.document.documentLayout()
        block = self.document.begin()
        while block.isValid():
            rect = layout.blockBoundingRect(block)
            for index in range(block.layout().lineCount()):
                line = block.layout().lineAt(index)
                logical = min(int((rect.top() + line.y()) / self.body_h + 0.00001), len(self._slots) - 1)
                if self._slots[logical][0] >= physical:
                    return self._source_position(block.position() + line.textStart())
            block = block.next()
        return self._source_position(self.document.characterCount() - 1)

    def page_settings(self, physical):
        if self._physical_settings:
            return self._physical_settings[min(physical, len(self._physical_settings) - 1)]
        return self.settings

    def page_size(self, physical):
        return tuple(mm_to_px(value) for value in self.page_settings(physical).size_mm())

    def _bound_oversized_images(self) -> None:
        doc = self.document
        block = doc.begin()
        while block.isValid():
            iterator = block.begin()
            while not iterator.atEnd():
                fragment = iterator.fragment()
                if fragment.isValid():
                    fmt = fragment.charFormat()
                    if fmt.isImageFormat():
                        image_fmt = QTextImageFormat(fmt.toImageFormat())
                        old_w = image_fmt.width()
                        old_h = image_fmt.height()
                        if old_w <= 0 or old_h <= 0:
                            iterator += 1
                            continue
                        scale = min(
                            1.0,
                            self.column_w / old_w,
                            self.body_h / old_h,
                        )
                        if scale < 1.0:
                            image_fmt.setWidth(old_w * scale)
                            image_fmt.setHeight(old_h * scale)
                            cursor = QTextCursor(doc)
                            cursor.setPosition(fragment.position())
                            cursor.setPosition(
                                fragment.position() + fragment.length(),
                                QTextCursor.MoveMode.KeepAnchor,
                            )
                            cursor.setCharFormat(image_fmt)
                iterator += 1
            block = block.next()

    def _find_break_pages(self) -> set[int]:
        breaks: set[int] = set()
        layout = self.document.documentLayout()
        _ = self.document.pageCount()
        block = self.document.begin()
        while block.isValid():
            if block_break_kind(block) == "page":
                top = layout.blockBoundingRect(block).top()
                index = int(top / self.body_h + 0.001)
                breaks.add(index)
            block = block.next()
        return breaks

    def _map_pages(self) -> None:
        columns = self.settings.columns
        logical = max(1, self.document.pageCount())
        physical = 0
        slot = 0
        for k in range(logical):
            if slot != 0 and k in self._break_pages:
                physical += 1
                slot = 0
            self._slots.append((physical, slot))
            slot += 1
            if slot == columns:
                physical += 1
                slot = 0

    def _line_locations(self):
        if self._flow:
            return list(dict.fromkeys((line.page, line.column, line.rect.top() - self.margin_top)
                                      for line in self._flow.lines))
        locations = []
        layout = self.document.documentLayout()
        block = self.document.begin()
        while block.isValid():
            block_rect = layout.blockBoundingRect(block)
            text_layout = block.layout()
            line_count = max(1, text_layout.lineCount())
            for index in range(line_count):
                line = (text_layout.lineAt(index)
                        if text_layout.lineCount() else None)
                y = block_rect.top() + (line.y() if line is not None else 0.0)
                logical = max(0, min(int(y / self.body_h + 0.00001),
                                     len(self._slots) - 1))
                physical, slot = self._slots[logical]
                if not self._tail or physical < self._tail[1]:
                    locations.append((physical, slot, y - logical * self.body_h))
            block = block.next()
        if self._tail:
            locations.extend((page + self._tail[1], slot, y)
                             for page, slot, y in self._tail[2]._line_locations())
        return locations

    def _map_line_numbers(self, start=None) -> list[tuple[int, int, float, int]]:
        if self.settings.line_numbering == "none":
            return []
        result: list[tuple[int, int, float, int]] = []
        number = self.settings.line_number_start if start is None else start
        previous_physical = 0
        for physical, slot, local_y in self._line_locations():
            if (self.settings.line_numbering == "restart_page"
                    and physical != previous_physical):
                number = self.settings.line_number_start
            previous_physical = physical
            if ((number - self.settings.line_number_start)
                    % self.settings.line_number_count_by == 0):
                result.append((physical, slot, local_y, number))
            number += 1
        return result

    def body_page_count(self) -> int:
        if self._physical_settings:
            return len(self._physical_settings)
        if self._flow:
            return max((line.page + 1 for line in self._flow.lines), default=1)
        if not self._slots:
            return 1
        return max(p for p, _ in self._slots) + 1

    def page_count(self) -> int:
        object_pages = max((obj.get("page", 0) + 1 for obj in visible_objects(self.objects)), default=1)
        return max(object_pages, self.body_page_count(), max(self._footnotes, default=-1) + 1)

    def page_for_position(self, position: int) -> int:
        if self._sections:
            for start, end, page, offset, child in reversed(self._sections):
                if start <= position:
                    return page + child.page_for_position(min(position, end) - start)
        position = layout_position(position, self._inserted_hyphens)
        if self._flow:
            return self._flow.page_for_position(position) + 1
        if self._tail and position >= self._tail[0]:
            return self._tail[1] + self._tail[2].page_for_position(position - self._tail[0])
        block = self.document.findBlock(position)
        if not block.isValid():
            return 1
        layout = self.document.documentLayout()
        top = layout.blockBoundingRect(block).top()
        text_layout = block.layout()
        if text_layout.lineCount():
            line = text_layout.lineForTextPosition(min(position - block.position(), block.length() - 1))
            if line.isValid():
                top += line.y()
        index = int(top / self.body_h + 0.001)
        index = max(0, min(index, len(self._slots) - 1))
        return self._slots[index][0] + 1

    def running_text(self, which: str, physical_index: int) -> str:
        settings = self.page_settings(physical_index)
        template = settings.header if which == "header" else settings.footer
        return (template
                .replace("{page}", str(physical_index + 1))
                .replace("{pages}", str(self.page_count()))
                .replace("{CurrentPage}", str(physical_index + 1))
                .replace("{TotalPages}", str(self.page_count()))
                .replace("{title}", self.title))

    def paint_page(self, painter: QPainter, physical_index: int,
                   scale: float = 1.0) -> None:
        if not 0 <= physical_index < self.page_count():
            raise IndexError("Page index out of range")
        painter.save()
        try:
            painter.scale(scale, scale)
            width, height = self.page_size(physical_index)
            paint_page_decoration(painter, QRectF(0, 0, width, height), self.design)
            paint_objects(painter, self.objects, physical_index, "behind")
            if self._sections:
                for start, end, page, offset, child in self._sections:
                    child_pages = max(child.body_page_count(), max(child._footnotes, default=-1) + 1)
                    if page <= physical_index < page + child_pages:
                        painter.save()
                        try:
                            painter.translate(0, offset if physical_index == page else 0)
                            child._paint_body(painter, physical_index - page)
                            child._paint_footnotes(painter, physical_index - page)
                        finally:
                            painter.restore()
            else:
                self._paint_body(painter, physical_index)
                self._paint_footnotes(painter, physical_index)
            paint_objects(painter, self.objects, physical_index, "front")
            paint_page_decoration(painter, QRectF(0, 0, width, height), self.design, "front")
            self._paint_running(painter, physical_index)
        finally:
            painter.restore()

    def _paint_running(self, painter, physical_index):
        settings = self.page_settings(physical_index)
        width, height = self.page_size(physical_index)
        left, right = mm_to_px(settings.left_mm), mm_to_px(settings.right_mm)
        top, bottom = mm_to_px(settings.top_mm), mm_to_px(settings.bottom_mm)
        body_font = QFont("Calibri")
        body_font.setPixelSize(12)
        painter.setFont(body_font)
        painter.setPen(QColor("#5b6b7f"))
        if settings.header:
            rect = QRectF(left, 0.0, width - left - right, top)
            painter.drawText(rect, Qt.AlignmentFlag.AlignVCenter
                             | Qt.AlignmentFlag.AlignLeft,
                             self.running_text("header", physical_index))
        if settings.footer:
            rect = QRectF(left, height - bottom, width - left - right, bottom)
            painter.drawText(rect, Qt.AlignmentFlag.AlignVCenter
                             | Qt.AlignmentFlag.AlignHCenter,
                             self.running_text("footer", physical_index))
    def _paint_body(self, painter, physical_index, line_numbers=True):
        if self._flow:
            self._flow.draw_page(painter, physical_index)
        elif self._tail and physical_index == self._tail[1]:
            self._tail[2]._paint_body(painter, 0, False)
        for logical, (physical, slot) in enumerate(self._slots if not self._flow and not (
                self._tail and physical_index == self._tail[1]) else []):
            if physical != physical_index:
                continue
            x = self.margin_left + slot * (self.column_w + self.gutter)
            painter.save()
            painter.setClipRect(QRectF(x, self.margin_top,
                                       self.column_w, self.body_h))
            painter.translate(x, self.margin_top)
            painter.translate(0.0, -logical * self.body_h)
            source = QRectF(0.0, logical * self.body_h,
                            self.column_w, self.body_h)
            self._draw_body_slice(painter, source)
            painter.restore()
        if line_numbers and self._numbered_lines:
            number_font = QFont("Calibri")
            number_font.setPixelSize(10)
            painter.setFont(number_font)
            painter.setPen(QColor("#6b7280"))
            distance = mm_to_px(self.settings.line_number_distance_mm)
            for physical, slot, local_y, number in self._numbered_lines:
                if physical != physical_index:
                    continue
                column_x = self.margin_left + slot * (
                    self.column_w + self.gutter)
                right = column_x - distance
                painter.drawText(
                    QRectF(max(0.0, right - 34.0),
                           self.margin_top + local_y - 2.0,
                           34.0, 14.0),
                    Qt.AlignmentFlag.AlignRight
                    | Qt.AlignmentFlag.AlignVCenter,
                    str(number))

    def _draw_body_slice(self, painter, source):
        layout = self.document.documentLayout()
        for frame in self.document.rootFrame().childFrames():
            rect = layout.frameBoundingRect(frame).intersected(source)
            if rect.isEmpty():
                continue
            painter.save()
            painter.setClipRect(rect, Qt.ClipOperation.IntersectClip)
            self.document.drawContents(painter, rect.adjusted(0, 0.01, 0, -0.01))
            painter.restore()
        block = self.document.begin()
        while block.isValid():
            if QTextCursor(block).currentFrame() != self.document.rootFrame():
                block = block.next()
                continue
            rect = layout.blockBoundingRect(block)
            block_format = block.blockFormat()
            if block_format.background().style() != Qt.BrushStyle.NoBrush:
                painter.fillRect(rect.intersected(source), block_format.background())
            text_layout = block.layout()
            for index in range(text_layout.lineCount()):
                line = text_layout.lineAt(index)
                y = rect.top() + line.y()
                if y < source.top() - 0.01 or y >= source.bottom() - 0.01:
                    continue
                line.draw(painter, rect.topLeft())
                if index == 0 and block.textList():
                    text_list = block.textList()
                    marker = text_list.itemText(block)
                    marker = {QTextListFormat.Style.ListDisc: "\u2022",
                              QTextListFormat.Style.ListCircle: "\u25e6",
                              QTextListFormat.Style.ListSquare: "\u25aa"}.get(
                                  text_list.format().style(), marker)
                    painter.save()
                    painter.setFont(block.charFormat().font().resolve(self.document.defaultFont()))
                    painter.setPen(block.charFormat().foreground().color())
                    painter.drawText(QRectF(rect.left() - 40, y, 33, line.height()),
                                     Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
                                     marker)
                    painter.restore()
            block = block.next()
        from .formatting import draw_paragraph_borders

        draw_paragraph_borders(painter, self.document)

    def paint(self, device) -> int:
        try:
            resolution = float(device.resolution())
        except Exception:
            resolution = LAYOUT_DPI
        scale = resolution / LAYOUT_DPI
        indices = list(range(self.page_count()))
        try:
            from PySide6.QtPrintSupport import QPrinter
        except ImportError:
            QPrinter = None
        if QPrinter is not None and isinstance(device, QPrinter):
            if device.printRange() == QPrinter.PrintRange.PageRange:
                ranges = device.pageRanges()
                if not ranges.isEmpty():
                    indices = [i for i in indices if ranges.contains(i + 1)]
                elif device.fromPage() > 0:
                    indices = [i for i in indices
                               if device.fromPage() <= i + 1 <= device.toPage()]
        elif hasattr(device, "pageRanges"):
            ranges = device.pageRanges()
            if not ranges.isEmpty():
                indices = [i for i in indices if ranges.contains(i + 1)]
        elif hasattr(device, "fromPage") and device.fromPage() > 0:
            indices = [i for i in indices
                       if device.fromPage() <= i + 1 <= device.toPage()]
        if not indices:
            raise RuntimeError(
                "The selected page range contains no document pages.")
        if hasattr(device, "setPageLayout"):
            device.setPageLayout(page_layout_for(self.page_settings(indices[0])))
        painter = QPainter(device)
        if not painter.isActive():
            raise RuntimeError("Could not start the print/PDF device.")
        try:
            for emitted, page_index in enumerate(indices):
                if emitted and hasattr(device, "setPageLayout"):
                    device.setPageLayout(page_layout_for(self.page_settings(page_index)))
                if emitted and not device.newPage():
                    raise RuntimeError(
                        "Could not create the next output page.")
                self.paint_page(painter, page_index, scale)
        finally:
            painter.end()
        return len(indices)


def write_pdf(path: Path, document, settings: PageSettings, title: str) -> int:
    path = Path(path)
    publication = Publication(document, settings, title)
    fd, temp_name = tempfile.mkstemp(prefix=path.name + ".", suffix=".pdf",
                                     dir=str(path.parent))
    os.close(fd)
    try:
        writer = QPdfWriter(temp_name)
        writer.setPageLayout(page_layout_for(settings))
        writer.setResolution(LAYOUT_DPI)
        writer.setTitle(title)
        count = publication.paint(writer)
        del writer
        os.replace(temp_name, path)
        return count
    except BaseException:
        try:
            os.unlink(temp_name)
        except OSError:
            pass
        raise
