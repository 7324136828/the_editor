from __future__ import annotations

import time
from contextlib import contextmanager

from PySide6.QtCore import QObject, Signal
from PySide6.QtGui import QTextCursor, QTextDocumentFragment, QTextFormat

from .editor import SafeDocument
from .models import Comment, CommentReply, DocumentState, Revision

OBJECT_MARKER = "￼"
STRUCTURAL_MARKERS = ("￼", "﷐", "﷑")
COALESCE_SECONDS = 2.0


def _has_structural(text: str) -> bool:
    return any(marker in text for marker in STRUCTURAL_MARKERS)


class ReviewError(RuntimeError):
    pass


def _utf16_to_py(text: str, units: int) -> int:
    count = 0
    for i, ch in enumerate(text):
        width = 2 if ord(ch) > 0xFFFF else 1
        if count + width > units:
            return i
        count += width
        if count >= units:
            return i + 1
    return len(text)


def _normalize(text: str) -> str:
    return text.replace(" ", "\n").replace(" ", "\n")


def _map_span(start: int, end: int, pos: int, removed: int, added: int):
    if start == end:
        point = start
        p1 = point - min(max(point - pos, 0), removed)
        p2 = p1 + added if pos < p1 else p1
        overlap = removed > 0 and pos < point < pos + removed
        return p2, p2, overlap, False
    overlap = removed > 0 and pos < end and pos + removed > start
    s1 = start - min(max(start - pos, 0), removed)
    e1 = end - min(max(end - pos, 0), removed)
    inside = added > 0 and s1 < pos < e1
    s2 = s1 + added if pos <= s1 else s1
    e2 = e1 + added if pos < e1 else e1
    if e2 < s2:
        e2 = s2
    return s2, e2, overlap, inside


class ReviewTracker(QObject):
    changed = Signal()
    edit_blocked = Signal(str)

    def __init__(self, document, state: DocumentState):
        super().__init__()
        self._doc = document
        self.state = state
        self._suppress = 0
        self._recording = True
        self._shadow_text = document.toPlainText()
        self._shadow_doc = None
        self._last_insert: Revision | None = None
        self._last_time = 0.0
        self.author = "You"
        self._guard_doc = None
        document.folio_review_tracker = self
        if state.track_changes:
            self._shadow_doc = self._clone_doc()
        document.contentsChange.connect(self._on_contents_change)
        self.refresh_protection()

    def _clone_doc(self):
        from .document_features import capture_node_bindings, restore_node_bindings

        clone = SafeDocument()
        clone.setDefaultFont(self._doc.defaultFont())
        clone.setHtml(self._doc.toHtml())
        restore_node_bindings(clone, capture_node_bindings(self._doc))
        return clone

    @contextmanager
    def loading(self):
        self._suppress += 1
        try:
            yield
        finally:
            self._suppress -= 1

    @contextmanager
    def applying(self):
        previous = self._recording
        self._recording = False
        try:
            yield
        finally:
            self._recording = previous

    def reset(self, state: DocumentState) -> None:
        self.state = state
        self._shadow_text = self._doc.toPlainText()
        self._shadow_doc = self._clone_doc() if state.track_changes else None
        self._last_insert = None
        self.refresh_protection()
        self.changed.emit()

    def refresh_protection(self) -> None:
        protected = self.state.review.get("locks") or self.state.review.get("formatting_locked")
        self._guard_doc = self._clone_doc() if protected else None

    def allows_edit(self, start: int, end: int, formatting: bool = False) -> bool:
        if formatting and self.state.review.get("formatting_locked"):
            return False
        for lock in self.state.review.get("locks", []):
            if ((start == end and lock["start"] <= start < lock["end"])
                    or (start < lock["end"] and end > lock["start"])):
                return False
        return True

    def _restore_protected(self) -> None:
        previous = self._guard_doc
        if previous is None:
            return
        with self.loading():
            if self._doc.isUndoAvailable():
                self._doc.undo()
            if self._doc.toHtml() != previous.toHtml():
                from .document_features import capture_node_bindings, restore_node_bindings

                self._doc.setHtml(previous.toHtml())
                restore_node_bindings(self._doc, capture_node_bindings(previous))
            self._doc.clearUndoRedoStacks(self._doc.Stacks.RedoStack)
        self.edit_blocked.emit("This range is protected, or formatting is restricted.")

    def set_tracking(self, enabled: bool) -> None:
        self.state.track_changes = bool(enabled)
        if enabled:
            if self._shadow_doc is None:
                self._shadow_doc = self._clone_doc()
        else:
            self._shadow_doc = None
            self._last_insert = None
        self.changed.emit()

    def _range_text(self, pos: int, length: int) -> str:
        cursor = QTextCursor(self._doc)
        cursor.setPosition(pos)
        cursor.setPosition(pos + length, QTextCursor.MoveMode.KeepAnchor)
        return _normalize(cursor.selectedText())

    def _range_fragment(self, pos: int, length: int) -> QTextDocumentFragment:
        cursor = QTextCursor(self._doc)
        cursor.setPosition(pos)
        cursor.setPosition(pos + length, QTextCursor.MoveMode.KeepAnchor)
        return QTextDocumentFragment(cursor)

    def _shadow_range_html(self, pos: int, length: int) -> str:
        if self._shadow_doc is None or length <= 0:
            return ""
        return range_html(self._shadow_doc, pos, length)

    def _range_html(self, pos: int, length: int) -> str:
        return range_html(self._doc, pos, length)

    def _shadow_replace(self, pos, removed, added):
        cursor = QTextCursor(self._shadow_doc)
        cursor.setPosition(pos)
        cursor.setPosition(pos + removed, QTextCursor.MoveMode.KeepAnchor)
        cursor.insertFragment(self._range_fragment(pos, added))

    def _on_contents_change(self, pos: int, removed: int, added: int) -> None:
        if self._suppress:
            return
        shadow = self._shadow_text
        shadow_length = len(shadow.encode("utf-16-le")) // 2
        removed = min(removed, max(0, shadow_length - pos))
        added = min(added, max(0, self._doc.characterCount() - 1 - pos))
        py_pos = _utf16_to_py(shadow, pos)
        py_end = _utf16_to_py(shadow, pos + removed)
        old_text = shadow[py_pos:py_end]
        new_text = self._range_text(pos, added)
        format_only = removed == added and old_text == new_text
        if (self._guard_doc is not None and (removed or added)
                and not self.allows_edit(pos, pos + removed, format_only)):
            self._restore_protected()
            return
        if removed == added and old_text == new_text:
            if self._shadow_doc is not None and removed:
                before_html = self._shadow_range_html(pos, removed)
                after_html = self._range_html(pos, added)
                if (self.state.track_changes and self._recording
                        and self.state.review.get("track_formatting")
                        and before_html != after_html
                        and not _has_structural(new_text)):
                    for revision in self.pending_revisions():
                        if (revision.kind == "format" and revision.start < pos + removed
                                and revision.start + revision.length > pos):
                            revision.conflicted = True
                    self.state.revisions.append(Revision(
                        start=pos, length=added, before_text=old_text,
                        after_text=new_text, before_html=before_html,
                        after_html=after_html, author=self.author, kind="format"))
                    self._last_insert = None
                    self.changed.emit()
                self._shadow_replace(pos, removed, added)
            self.refresh_protection()
            return
        self._transform_all(pos, removed, added)
        if self.state.track_changes and self._recording:
            if not _has_structural(old_text) and not _has_structural(new_text):
                self._record(pos, removed, added, old_text, new_text)
            else:
                self._last_insert = None
        else:
            self._last_insert = None
        self._shadow_text = shadow[:py_pos] + new_text + shadow[py_end:]
        if self._shadow_doc is not None:
            cursor = QTextCursor(self._shadow_doc)
            cursor.setPosition(pos)
            if removed:
                cursor.setPosition(pos + removed, QTextCursor.MoveMode.KeepAnchor)
                cursor.removeSelectedText()
                cursor = QTextCursor(self._shadow_doc)
                cursor.setPosition(pos)
            if added:
                fragment = self._range_fragment(pos, added)
                cursor.insertFragment(fragment)
        self.refresh_protection()
        self.changed.emit()

    def _transform_all(self, pos: int, removed: int, added: int) -> None:
        for revision in self.state.revisions:
            if revision.status != "pending":
                continue
            start, end = revision.start, revision.start + revision.length
            s2, e2, overlap, inside = _map_span(start, end, pos, removed, added)
            if s2 != start or e2 != end:
                revision.start = s2
                revision.length = e2 - s2
            if overlap or inside:
                revision.conflicted = True
        for comment in self.state.comments:
            start, end = comment.start, comment.end
            s2, e2, overlap, _inside = _map_span(start, end, pos, removed, added)
            if s2 != start or e2 != end:
                comment.start = s2
                comment.end = e2
            if overlap:
                comment.orphaned = True
        for lock in self.state.review.get("locks", []):
            lock["start"], lock["end"], _, _ = _map_span(
                lock["start"], lock["end"], pos, removed, added)
        if self.state.toc is not None:
            toc = self.state.toc
            start, end = toc.start, toc.start + toc.length
            if pos + removed <= start and not (removed == 0 and pos == start):
                toc.start += added - removed
            elif pos <= end:
                self.state.toc = None

    def _record(self, pos: int, removed: int, added: int,
                old_text: str, new_text: str) -> None:
        now = time.monotonic()
        if removed == 0 and added > 0:
            last = self._last_insert
            if (last is not None and last.status == "pending"
                    and not last.conflicted and not last.before_text
                    and pos == last.start + last.length
                    and now - self._last_time < COALESCE_SECONDS
                    and "\n" not in new_text and "\n" not in last.after_text):
                last.length += added
                last.after_text += new_text
                last.after_html = self._range_html(last.start, last.length)
                self._last_time = now
                return
            revision = Revision(
                start=pos, length=added, before_text="", after_text=new_text,
                before_html="", after_html=self._range_html(pos, added),
                author=self.author,
            )
            self.state.revisions.append(revision)
            self._last_insert = revision
            self._last_time = now
            return
        self._last_insert = None
        if removed > 0 and added == 0:
            before_html = self._shadow_range_html(pos, removed)
            revision = Revision(
                start=pos, length=0, before_text=old_text, after_text="",
                before_html=before_html, after_html="",
                author=self.author,
            )
            self.state.revisions.append(revision)
            return
        if removed > 0 and added > 0:
            before_html = self._shadow_range_html(pos, removed)
            revision = Revision(
                start=pos, length=added, before_text=old_text,
                after_text=new_text, before_html=before_html,
                after_html=self._range_html(pos, added), author=self.author,
            )
            self.state.revisions.append(revision)

    def _find(self, revision_id: str) -> Revision | None:
        for revision in self.state.revisions:
            if revision.id == revision_id:
                return revision
        return None

    def pending_revisions(self) -> list[Revision]:
        return [r for r in self.state.revisions if r.status == "pending"]

    def unresolved_comments(self) -> list[Comment]:
        return [c for c in self.state.comments if not c.resolved]

    def add_comment(self, start: int, end: int, text: str,
                    author: str = "You") -> Comment:
        if end < start:
            start, end = end, start
        quote = self._range_text(start, end - start)
        comment = Comment(start=start, end=end, quote=quote, text=text, author=author)
        self.state.comments.append(comment)
        self.changed.emit()
        return comment

    def resolve_comment(self, comment_id: str, resolved: bool = True) -> None:
        for comment in self.state.comments:
            if comment.id == comment_id:
                comment.resolved = resolved
                self.changed.emit()
                return

    def reply_to_comment(self, comment_id: str, text: str,
                         author: str | None = None) -> CommentReply:
        if not text.strip():
            raise ReviewError("A reply cannot be empty.")
        comment = next((item for item in self.state.comments if item.id == comment_id), None)
        if comment is None:
            raise ReviewError("This comment no longer exists.")
        reply = CommentReply(text=text.strip(), author=author or self.author)
        comment.replies.append(reply)
        self.changed.emit()
        return reply

    def delete_comment(self, comment_id: str) -> None:
        self.state.comments[:] = [item for item in self.state.comments if item.id != comment_id]
        self.changed.emit()

    def chronological_comments(self, author: str | None = None,
                               include_resolved: bool = True) -> list[Comment]:
        return sorted((item for item in self.state.comments
                       if (include_resolved or not item.resolved)
                       and (not author or item.author == author
                            or any(reply.author == author for reply in item.replies))),
                      key=lambda item: (item.created, item.id))

    def _check_rejectable(self, revision: Revision) -> None:
        if revision.conflicted:
            raise ReviewError(
                "This change overlaps newer edits; review it manually."
            )
        current = self._range_text(revision.start, revision.length)
        if current != revision.after_text:
            revision.conflicted = True
            raise ReviewError(
                "This range has changed; review it manually."
            )
        if (revision.kind == "format" and formatting_signature(
                self._range_html(revision.start, revision.length), self._doc.defaultFont())
                != formatting_signature(revision.after_html, self._doc.defaultFont())):
            revision.conflicted = True
            raise ReviewError("Formatting in this range has changed; review it manually.")

    def reject_revision(self, revision_id: str) -> None:
        revision = self._find(revision_id)
        if revision is None or revision.status != "pending":
            raise ReviewError("This change is no longer pending.")
        if not self.allows_edit(revision.start, revision.start + revision.length,
                                revision.kind == "format"):
            raise ReviewError("Remove this range's editing restriction before rejecting its change.")
        try:
            self._check_rejectable(revision)
        except ReviewError:
            self.changed.emit()
            raise
        self.state.revisions.remove(revision)
        cursor = QTextCursor(self._doc)
        cursor.setPosition(revision.start)
        cursor.setPosition(revision.start + revision.length,
                           QTextCursor.MoveMode.KeepAnchor)
        with self.applying():
            cursor.beginEditBlock()
            if revision.kind == "format":
                restore_formatting(self._doc, revision.start, revision.length,
                                   revision.before_html)
            else:
                if revision.length:
                    cursor.removeSelectedText()
                if revision.before_html:
                    fragment = QTextDocumentFragment.fromHtml(revision.before_html)
                    cursor.insertFragment(fragment)
            cursor.endEditBlock()
        self._last_insert = None
        self.changed.emit()

    def accept_revision(self, revision_id: str) -> None:
        revision = self._find(revision_id)
        if revision is None or revision.status != "pending":
            raise ReviewError("This change is no longer pending.")
        revision.status = "accepted"
        self._last_insert = None
        self.changed.emit()

    def accept_all(self) -> int:
        count = 0
        for revision in self.state.revisions:
            if revision.status == "pending":
                revision.status = "accepted"
                count += 1
        self._last_insert = None
        if count:
            self.changed.emit()
        return count

    def reject_all(self) -> int:
        pending = self.pending_revisions()
        for revision in pending:
            self._check_rejectable(revision)
            if not self.allows_edit(revision.start, revision.start + revision.length,
                                    revision.kind == "format"):
                raise ReviewError("Remove editing restrictions before rejecting these changes.")
        count = 0
        for _, revision in sorted(enumerate(pending),
                                  key=lambda item: (item[1].start, item[0]), reverse=True):
            self.reject_revision(revision.id)
            count += 1
        return count


def restore_formatting(document, start: int, length: int, source: str) -> None:
    original = SafeDocument()
    original.setDefaultFont(document.defaultFont())
    original.setHtml(source)
    block = original.begin()
    bound = start + length
    lists = {}
    while block.isValid() and start + block.position() <= bound:
        target = QTextCursor(document)
        target.setPosition(min(start + block.position(), bound))
        source_list = block.textList()
        if source_list is None:
            if target.currentList() is not None:
                target.currentList().remove(target.block())
        elif source_list.objectIndex() in lists:
            lists[source_list.objectIndex()].add(target.block())
        else:
            lists[source_list.objectIndex()] = target.createList(source_list.format())
        if block_signature(target.blockFormat()) != block_signature(block.blockFormat()):
            target.setBlockFormat(block.blockFormat())
        iterator = block.begin()
        while not iterator.atEnd():
            fragment = iterator.fragment()
            if fragment.isValid():
                left = start + fragment.position()
                right = min(bound, left + fragment.length())
                if left < right:
                    target.setPosition(left)
                    target.setPosition(right, QTextCursor.MoveMode.KeepAnchor)
                    target.setCharFormat(fragment.charFormat())
            iterator += 1
        block = block.next()


def block_signature(fmt):
    return (fmt.alignment(), fmt.headingLevel(), fmt.indent(), fmt.textIndent(),
            fmt.topMargin(), fmt.bottomMargin(), fmt.leftMargin(), fmt.rightMargin(),
            fmt.lineHeight(), fmt.lineHeightType(), fmt.pageBreakPolicy(),
            fmt.nonBreakableLines(), fmt.background().style(), fmt.background().color().rgba(),
            tuple((tab.position, tab.type, tab.delimiter) for tab in fmt.tabPositions()),
            tuple((key, repr(value)) for key, value in sorted(fmt.properties().items())
                  if key >= QTextFormat.Property.UserProperty))


def formatting_signature(source, default_font):
    document = SafeDocument()
    document.setDefaultFont(default_font)
    document.setHtml(source)
    result = []
    block = document.begin()
    while block.isValid():
        runs = []
        iterator = block.begin()
        while not iterator.atEnd():
            fragment = iterator.fragment()
            if fragment.isValid():
                fmt = fragment.charFormat()
                signature = (fmt.font().resolve(default_font).toString(),
                             fmt.foreground().style(), fmt.foreground().color().rgba(),
                             fmt.background().style(), fmt.background().color().rgba(),
                             fmt.verticalAlignment(), fmt.anchorHref(), tuple(fmt.anchorNames()))
                if runs and runs[-1][1] == signature:
                    runs[-1] = (runs[-1][0] + fragment.text(), signature)
                else:
                    runs.append((fragment.text(), signature))
            iterator += 1
        listing = block.textList()
        list_signature = (listing.format().style(), listing.format().indent(),
                          listing.format().numberPrefix(), listing.format().numberSuffix()) if listing else None
        result.append((block_signature(block.blockFormat()), list_signature, tuple(runs)))
        block = block.next()
    return tuple(result)


def range_html(document, start, length):
    cursor = QTextCursor(document)
    cursor.setPosition(start)
    cursor.setPosition(start + length, QTextCursor.MoveMode.KeepAnchor)
    fragment = SafeDocument()
    fragment.setDefaultFont(document.defaultFont())
    QTextCursor(fragment).insertFragment(QTextDocumentFragment(cursor))
    return fragment.toHtml()
