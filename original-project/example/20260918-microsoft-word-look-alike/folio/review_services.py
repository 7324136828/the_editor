from __future__ import annotations

import copy
import difflib
import json
import math
import re
from dataclasses import dataclass, field
from pathlib import Path

from PySide6.QtCore import QObject, Signal, Qt
from PySide6.QtGui import QColor, QTextCharFormat, QTextCursor, QTextDocumentFragment

from .editor import SafeDocument
from .review import ReviewError, ReviewTracker


def checked_ink(data: object) -> dict:
    if not isinstance(data, dict):
        raise ValueError("Ink drawing must be an object.")
    width, height = data.get("width"), data.get("height")
    if any(type(value) is not int or not 1 <= value <= 4096 for value in (width, height)):
        raise ValueError("Ink dimensions must be whole numbers from 1 to 4096.")
    strokes = data.get("strokes")
    if not isinstance(strokes, list) or len(strokes) > 10000:
        raise ValueError("Ink must contain at most 10,000 strokes.")
    result = {"width": width, "height": height, "strokes": []}
    count = 0
    for stroke in strokes:
        if not isinstance(stroke, dict):
            raise ValueError("Ink stroke must be an object.")
        color, thickness, points = stroke.get("color"), stroke.get("width"), stroke.get("points")
        if not isinstance(color, str) or not re.fullmatch(r"#[0-9a-fA-F]{6}", color):
            raise ValueError("Ink color must be a six-digit hex color.")
        if (isinstance(thickness, bool) or not isinstance(thickness, (int, float))
                or not math.isfinite(thickness) or not 0 < thickness <= 64):
            raise ValueError("Ink stroke width must be greater than zero and at most 64.")
        if not isinstance(points, list) or not points:
            raise ValueError("Ink stroke must have at least one point.")
        count += len(points)
        if count > 200000:
            raise ValueError("Ink drawing exceeds 200,000 points.")
        checked_points = []
        for point in points:
            if (not isinstance(point, (list, tuple)) or len(point) != 2
                    or any(isinstance(value, bool) or not isinstance(value, (int, float))
                           or not math.isfinite(value) for value in point)
                    or not 0 <= point[0] <= width or not 0 <= point[1] <= height):
                raise ValueError("Ink point is outside the drawing.")
            checked_points.append([float(point[0]), float(point[1])])
        result["strokes"].append({"color": color.lower(), "width": float(thickness), "points": checked_points})
    return result


@dataclass(frozen=True)
class AccessibilityIssue:
    start: int
    length: int
    message: str
    rule: str


def contrast_ratio(first: QColor, second: QColor) -> float:
    def luminance(color):
        values = [color.redF(), color.greenF(), color.blueF()]
        values = [value / 12.92 if value <= 0.04045
                  else ((value + 0.055) / 1.055) ** 2.4 for value in values]
        return sum(value * weight for value, weight in zip(values, (0.2126, 0.7152, 0.0722)))

    dark, light = sorted((luminance(first), luminance(second)))
    return (light + 0.05) / (dark + 0.05)


def accessibility_issues(document) -> list[AccessibilityIssue]:
    issues = []
    previous_heading = 0
    seen_tables = set()
    block = document.begin()
    while block.isValid():
        level = block.blockFormat().headingLevel()
        if level:
            if level > previous_heading + 1:
                issues.append(AccessibilityIssue(block.position(), block.length() - 1,
                              f"Heading level jumps from {previous_heading} to {level}.", "headings"))
            previous_heading = level
        cursor = QTextCursor(block)
        table = cursor.currentTable()
        if table is not None and table.firstPosition() not in seen_tables:
            seen_tables.add(table.firstPosition())
            if not table.format().headerRowCount():
                issues.append(AccessibilityIssue(block.position(), block.length() - 1,
                              "Table has no designated header row.", "table_headers"))
        iterator = block.begin()
        while not iterator.atEnd():
            fragment = iterator.fragment()
            if fragment.isValid():
                fmt = fragment.charFormat()
                if fmt.isImageFormat() and not str(fmt.property(fmt.Property.ImageAltText) or "").strip():
                    issues.append(AccessibilityIssue(fragment.position(), fragment.length(),
                                  "Image has no alternative text.", "image_alt"))
                if fmt.isAnchor() and fmt.anchorHref() and fragment.text().strip().casefold() in (
                        "click here", "here", "read more", "link"):
                    issues.append(AccessibilityIssue(fragment.position(), fragment.length(),
                                  "Use descriptive link text instead of a generic label.", "link_text"))
                if not fmt.isImageFormat() and fragment.text().strip():
                    foreground = (fmt.foreground().color() if fmt.foreground().style() != Qt.BrushStyle.NoBrush
                                  else QColor("#000000"))
                    background = (fmt.background().color() if fmt.background().style() != Qt.BrushStyle.NoBrush
                                  else QColor("#ffffff"))
                    size = fmt.fontPointSize() or document.defaultFont().pointSizeF()
                    minimum = 3.0 if size >= 18 or size >= 14 and fmt.fontWeight() >= 700 else 4.5
                    ratio = contrast_ratio(foreground, background)
                    if ratio + 0.001 < minimum:
                        issues.append(AccessibilityIssue(fragment.position(), fragment.length(),
                                      f"Text contrast is {ratio:.1f}:1; check against {minimum:g}:1.", "contrast"))
            iterator += 1
        block = block.next()
    return issues


def review_document(document, state, mode: str) -> SafeDocument:
    if mode not in ("All Markup", "Simple Markup", "No Markup", "Original"):
        raise ValueError("Unknown review display mode")
    result = SafeDocument()
    result.setDefaultFont(document.defaultFont())
    result.setHtml(document.toHtml())
    if mode == "No Markup":
        return result
    if mode == "Original":
        tracker = ReviewTracker(result, copy.deepcopy(state))
        tracker.state.review = {}
        tracker.refresh_protection()
        for revision in reversed(list(tracker.pending_revisions())):
            tracker.reject_revision(revision.id)
        return result
    pending = [revision for revision in state.revisions if revision.status == "pending"]
    for revision in sorted(pending, key=lambda value: value.start, reverse=True):
        cursor = QTextCursor(result)
        start = min(revision.start, result.characterCount() - 1)
        cursor.setPosition(start)
        cursor.setPosition(min(start + revision.length, result.characterCount() - 1),
                           QTextCursor.MoveMode.KeepAnchor)
        fmt = QTextCharFormat()
        fmt.setBackground(QColor("#e8ddff" if revision.kind == "format" else "#dcf5dd"))
        if mode == "All Markup" and revision.kind == "text":
            fmt.setFontUnderline(True)
        if cursor.hasSelection():
            cursor.mergeCharFormat(fmt)
        elif mode == "Simple Markup":
            cursor.movePosition(QTextCursor.MoveOperation.NextCharacter, QTextCursor.MoveMode.KeepAnchor)
            cursor.mergeCharFormat(fmt)
        if mode == "All Markup" and revision.before_text and revision.kind == "text":
            cursor.setPosition(start)
            deleted = QTextCharFormat()
            deleted.setForeground(QColor("#ad2337"))
            deleted.setFontStrikeOut(True)
            deleted.setBackground(QColor("#ffe7eb"))
            cursor.insertText(revision.before_text, deleted)
    return result


@dataclass(frozen=True)
class MergeConflict:
    base_start: int
    base_end: int
    base: str
    current: str
    incoming: str


@dataclass
class ComparisonResult:
    text: str
    conflicts: list[MergeConflict] = field(default_factory=list)


def compare_three_way(base: str, current: str, incoming: str) -> ComparisonResult:
    original = base.splitlines(keepends=True)
    versions = [current.splitlines(keepends=True), incoming.splitlines(keepends=True)]
    edits = []
    for side, version in enumerate(versions):
        for kind, start, end, left, right in difflib.SequenceMatcher(
                None, original, version, autojunk=False).get_opcodes():
            if kind != "equal":
                edits.append((start, end, side, version[left:right]))
    edits.sort(key=lambda item: (item[0], item[1], item[2]))
    groups = []
    for edit in edits:
        if groups:
            previous = groups[-1]
            overlaps = any((edit[0] < value[1] and edit[1] > value[0])
                           or (edit[0] == edit[1] and value[0] <= edit[0] <= value[1])
                           or (value[0] == value[1] and edit[0] <= value[0] <= edit[1])
                           for value in previous)
            if overlaps:
                previous.append(edit)
                continue
        groups.append([edit])
    output = []
    conflicts = []
    position = 0
    for group in groups:
        start = min(item[0] for item in group)
        end = max(item[1] for item in group)
        output.extend(original[position:start])
        alternatives = []
        for side in (0, 1):
            lines = list(original[start:end])
            for left, right, _, replacement in reversed([item for item in group if item[2] == side]):
                lines[left - start:right - start] = replacement
            alternatives.append("".join(lines))
        before = "".join(original[start:end])
        if alternatives[0] == alternatives[1]:
            output.append(alternatives[0])
        elif alternatives[0] == before:
            output.append(alternatives[1])
        elif alternatives[1] == before:
            output.append(alternatives[0])
        else:
            conflicts.append(MergeConflict(start + 1, end, before, *alternatives))
            output.append("<<<<<<< CURRENT\n" + alternatives[0].rstrip("\n")
                          + "\n=======\n" + alternatives[1].rstrip("\n")
                          + "\n>>>>>>> INCOMING\n")
        position = end
    output.extend(original[position:])
    return ComparisonResult("".join(output), conflicts)


def load_glossary(path: str | Path) -> tuple[str, str, dict[str, str]]:
    source = Path(path)
    if source.stat().st_size > 5 * 1024 * 1024:
        raise ValueError("Glossary must be smaller than 5 MiB.")
    data = json.loads(source.read_text(encoding="utf-8-sig"))
    if not isinstance(data, dict) or not isinstance(data.get("entries"), dict):
        raise ValueError('Use a JSON object with "source_language", "target_language", and "entries".')
    entries = data["entries"]
    if (len(entries) > 10000 or any(not isinstance(key, str) or not key.strip()
                                  or not isinstance(value, str) for key, value in entries.items())):
        raise ValueError("Glossary entries must be nonempty text keys and text values (at most 10,000 entries).")
    for language in ("source_language", "target_language"):
        if not isinstance(data.get(language), str) or not data[language].strip():
            raise ValueError(f"Glossary needs {language}.")
    return data["source_language"], data["target_language"], entries


def translate_glossary(text: str, entries: dict[str, str]) -> tuple[str, int]:
    if not entries:
        return text, 0
    terms = sorted(entries, key=len, reverse=True)
    lookup = {key.casefold(): value for key, value in entries.items()}
    pattern = re.compile(r"(?<!\w)(?:" + "|".join(re.escape(key) for key in terms) + r")(?!\w)", re.I)
    return pattern.subn(lambda match: lookup[match.group().casefold()], text)


class SpeechReader(QObject):
    word_range = Signal(int, int)
    message = Signal(str)
    finished = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.engine = None
        self._offset = 0

    def start(self, text: str, offset: int = 0) -> bool:
        if not text.strip():
            self.message.emit("Select or enter some text to read.")
            return False
        if self.engine is None:
            try:
                from PySide6.QtTextToSpeech import QTextToSpeech
                available = QTextToSpeech.availableEngines()
                native = next((name for name in ("sapi", "winrt", "speechd", "darwin", "flite")
                               if name in available), None)
                if native is None:
                    self.message.emit("No supported local speech engine is installed.")
                    return False
                self.engine = QTextToSpeech(native, self)
                self.engine.sayingWord.connect(self._word)
                self.engine.stateChanged.connect(self._state)
                self.engine.errorOccurred.connect(lambda *_: self.message.emit(self.engine.errorString()))
            except (ImportError, RuntimeError) as error:
                self.message.emit(f"Local speech is unavailable: {error}")
                return False
        if not self.engine.availableVoices():
            self.message.emit("The local speech engine has no voice installed.")
            return False
        from PySide6.QtTextToSpeech import QTextToSpeech
        if not self.engine.engineCapabilities() & QTextToSpeech.Capability.WordByWordProgress:
            self.message.emit("This local speech engine cannot report word positions for highlighting.")
        self.stop()
        self._offset = offset
        self.engine.say(text)
        return self.engine.state() != QTextToSpeech.State.Error

    def _word(self, word: str, _identifier: int, start: int, length: int) -> None:
        self.word_range.emit(self._offset + start, length)

    def _state(self, state) -> None:
        from PySide6.QtTextToSpeech import QTextToSpeech
        if state == QTextToSpeech.State.Ready:
            self.finished.emit()
        elif state == QTextToSpeech.State.Error:
            self.message.emit(self.engine.errorString() or "Local speech is unavailable.")
            self.finished.emit()

    def pause(self) -> None:
        if self.engine is not None:
            self.engine.pause()

    def resume(self) -> None:
        if self.engine is not None:
            self.engine.resume()

    def stop(self) -> None:
        if self.engine is not None:
            self.engine.stop()
        self.finished.emit()
