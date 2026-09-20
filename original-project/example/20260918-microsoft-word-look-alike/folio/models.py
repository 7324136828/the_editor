from __future__ import annotations

import math
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


PAPER_SIZES_MM: dict[str, tuple[float, float]] = {
    "A4": (210.0, 297.0),
    "Letter": (215.9, 279.4),
    "Legal": (215.9, 355.6),
    "Envelope #10": (104.775, 241.3),
    "Envelope DL": (110.0, 220.0),
    "Envelope C5": (162.0, 229.0),
}

MIN_COLUMN_WIDTH_MM = 35.0
MIN_CONTENT_HEIGHT_MM = 40.0
SCHEMA_VERSION = 1


class ValidationError(ValueError):
    pass


def _finite_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


@dataclass
class PageSettings:
    paper: str = "A4"
    landscape: bool = False
    top_mm: float = 25.4
    right_mm: float = 25.4
    bottom_mm: float = 25.4
    left_mm: float = 25.4
    columns: int = 1
    gutter_mm: float = 10.0
    header: str = ""
    footer: str = "Page {page} of {pages}"
    line_numbering: str = "none"
    line_number_start: int = 1
    line_number_count_by: int = 1
    line_number_distance_mm: float = 5.0
    hyphenation: str = "none"
    hyphenate_caps: bool = False
    consecutive_hyphen_limit: int = 0
    balance_columns: bool = True

    def size_mm(self) -> tuple[float, float]:
        width, height = PAPER_SIZES_MM[self.paper]
        if self.landscape:
            return height, width
        return width, height

    def content_width_mm(self) -> float:
        width, _ = self.size_mm()
        return width - self.left_mm - self.right_mm

    def column_width_mm(self) -> float:
        usable = self.content_width_mm() - (self.columns - 1) * self.gutter_mm
        return usable / self.columns

    def content_height_mm(self) -> float:
        _, height = self.size_mm()
        return height - self.top_mm - self.bottom_mm

    def validate(self) -> None:
        if self.paper not in PAPER_SIZES_MM:
            raise ValidationError(f"Unknown paper size: {self.paper!r}")
        if not isinstance(self.columns, int) or isinstance(self.columns, bool) or not 1 <= self.columns <= 3:
            raise ValidationError("Columns must be a whole number between 1 and 3")
        for name in ("top_mm", "right_mm", "bottom_mm", "left_mm", "gutter_mm"):
            value = getattr(self, name)
            if not _finite_number(value) or value < 0:
                raise ValidationError(f"{name} must be a finite, non-negative number")
        if self.content_width_mm() <= 0 or self.content_height_mm() <= 0:
            raise ValidationError("Margins consume the entire page content area")
        if self.column_width_mm() < MIN_COLUMN_WIDTH_MM:
            raise ValidationError(
                f"Each column must be at least {MIN_COLUMN_WIDTH_MM:.0f} mm wide "
                f"(current: {self.column_width_mm():.1f} mm)"
            )
        if self.content_height_mm() < MIN_CONTENT_HEIGHT_MM:
            raise ValidationError(
                f"Content height must be at least {MIN_CONTENT_HEIGHT_MM:.0f} mm"
            )
        if self.line_numbering not in (
                "none", "continuous", "restart_page", "restart_section"):
            raise ValidationError("Unknown line-numbering mode")
        if (not isinstance(self.line_number_start, int)
                or isinstance(self.line_number_start, bool)
                or self.line_number_start < 1):
            raise ValidationError("Line numbering must start at 1 or later")
        if (not isinstance(self.line_number_count_by, int)
                or isinstance(self.line_number_count_by, bool)
                or self.line_number_count_by < 1):
            raise ValidationError("Line-number interval must be at least 1")
        if (not _finite_number(self.line_number_distance_mm)
                or self.line_number_distance_mm < 0):
            raise ValidationError(
                "Line-number distance must be a finite, non-negative number")
        if self.hyphenation not in ("none", "automatic"):
            raise ValidationError("Unknown hyphenation mode")
        if not isinstance(self.balance_columns, bool):
            raise ValidationError("balance_columns must be a boolean")
        if not isinstance(self.hyphenate_caps, bool):
            raise ValidationError("hyphenate_caps must be a boolean")
        if (not isinstance(self.consecutive_hyphen_limit, int)
                or isinstance(self.consecutive_hyphen_limit, bool)
                or self.consecutive_hyphen_limit < 0):
            raise ValidationError(
                "Consecutive-hyphen limit must be zero or a positive integer")

    def to_dict(self) -> dict:
        return {
            "paper": self.paper,
            "landscape": bool(self.landscape),
            "top_mm": float(self.top_mm),
            "right_mm": float(self.right_mm),
            "bottom_mm": float(self.bottom_mm),
            "left_mm": float(self.left_mm),
            "columns": int(self.columns),
            "gutter_mm": float(self.gutter_mm),
            "header": str(self.header),
            "footer": str(self.footer),
            "line_numbering": str(self.line_numbering),
            "line_number_start": int(self.line_number_start),
            "line_number_count_by": int(self.line_number_count_by),
            "line_number_distance_mm": float(self.line_number_distance_mm),
            "hyphenation": str(self.hyphenation),
            "hyphenate_caps": bool(self.hyphenate_caps),
            "consecutive_hyphen_limit": int(self.consecutive_hyphen_limit),
            "balance_columns": self.balance_columns,
        }

    @classmethod
    def from_dict(cls, data: object) -> PageSettings:
        if not isinstance(data, dict):
            raise ValidationError("Page settings must be an object")
        allowed = {
            "paper": str, "landscape": bool, "top_mm": float, "right_mm": float,
            "bottom_mm": float, "left_mm": float, "columns": int,
            "gutter_mm": float, "header": str, "footer": str,
            "line_numbering": str, "line_number_start": int,
            "line_number_count_by": int, "line_number_distance_mm": float,
            "hyphenation": str, "hyphenate_caps": bool,
            "consecutive_hyphen_limit": int,
            "balance_columns": bool,
        }
        values: dict = {}
        for key, kind in allowed.items():
            raw = data.get(key)
            if raw is None:
                continue
            if kind is bool:
                if not isinstance(raw, bool):
                    raise ValidationError(f"page.{key} must be a boolean")
                values[key] = raw
            elif kind is int:
                if not isinstance(raw, int) or isinstance(raw, bool):
                    raise ValidationError(f"page.{key} must be an integer")
                values[key] = raw
            elif kind is float:
                if not _finite_number(raw):
                    raise ValidationError(f"page.{key} must be a finite number")
                values[key] = float(raw)
            else:
                if not isinstance(raw, str):
                    raise ValidationError(f"page.{key} must be a string")
                values[key] = raw
        settings = cls(**values)
        settings.validate()
        return settings


@dataclass
class CommentReply:
    text: str
    author: str = "You"
    created: str = field(default_factory=utc_now)
    id: str = field(default_factory=lambda: str(uuid.uuid4()))

    def to_dict(self) -> dict:
        return {"text": self.text, "author": self.author,
                "created": self.created, "id": self.id}

    @classmethod
    def from_dict(cls, data: object) -> CommentReply:
        if not isinstance(data, dict) or not isinstance(data.get("text"), str):
            raise ValidationError("Comment reply must contain text")
        return cls(text=data["text"], author=_opt_str(data, "author") or "You",
                   created=_opt_str(data, "created") or utc_now(),
                   id=_opt_str(data, "id") or str(uuid.uuid4()))


@dataclass
class Comment:
    start: int
    end: int
    quote: str
    text: str
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    author: str = "You"
    created: str = field(default_factory=utc_now)
    resolved: bool = False
    orphaned: bool = False
    replies: list[CommentReply] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "start": self.start,
            "end": self.end,
            "quote": self.quote,
            "text": self.text,
            "author": self.author,
            "created": self.created,
            "resolved": self.resolved,
            "orphaned": self.orphaned,
            "replies": [reply.to_dict() for reply in self.replies],
        }

    @classmethod
    def from_dict(cls, data: object) -> Comment:
        if not isinstance(data, dict):
            raise ValidationError("Comment must be an object")
        try:
            start = data["start"]
            end = data["end"]
            quote = data["quote"]
            text = data["text"]
        except KeyError as exc:
            raise ValidationError(f"Comment missing field {exc}") from exc
        if not _is_position(start) or not _is_position(end) or end < start:
            raise ValidationError("Comment range invalid")
        if not isinstance(quote, str) or not isinstance(text, str):
            raise ValidationError("Comment text must be a string")
        replies = data.get("replies", [])
        if not isinstance(replies, list):
            raise ValidationError("Comment replies must be a list")
        return cls(
            start=start,
            end=end,
            quote=quote,
            text=text,
            id=_opt_str(data, "id") or str(uuid.uuid4()),
            author=_opt_str(data, "author") or "You",
            created=_opt_str(data, "created") or utc_now(),
            resolved=_opt_bool(data, "resolved"),
            orphaned=_opt_bool(data, "orphaned"),
            replies=[CommentReply.from_dict(reply) for reply in replies],
        )


@dataclass
class Revision:
    start: int
    length: int
    before_text: str
    after_text: str
    before_html: str
    after_html: str
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    author: str = "You"
    created: str = field(default_factory=utc_now)
    status: str = "pending"
    conflicted: bool = False
    kind: str = "text"

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "start": self.start,
            "length": self.length,
            "before_text": self.before_text,
            "after_text": self.after_text,
            "before_html": self.before_html,
            "after_html": self.after_html,
            "author": self.author,
            "created": self.created,
            "status": self.status,
            "conflicted": self.conflicted,
            "kind": self.kind,
        }

    @classmethod
    def from_dict(cls, data: object) -> Revision:
        if not isinstance(data, dict):
            raise ValidationError("Revision must be an object")
        try:
            start = data["start"]
            length = data["length"]
            before_text = data["before_text"]
            after_text = data["after_text"]
            before_html = data["before_html"]
            after_html = data["after_html"]
        except KeyError as exc:
            raise ValidationError(f"Revision missing field {exc}") from exc
        if not _is_position(start) or not _is_position(length):
            raise ValidationError("Revision range invalid")
        for name, value in (("before_text", before_text), ("after_text", after_text),
                            ("before_html", before_html), ("after_html", after_html)):
            if not isinstance(value, str):
                raise ValidationError(f"Revision {name} must be a string")
        status = data.get("status", "pending")
        if status not in ("pending", "accepted", "rejected"):
            raise ValidationError("Revision status invalid")
        kind = data.get("kind", "text")
        if kind not in ("text", "format"):
            raise ValidationError("Revision kind invalid")
        return cls(
            start=start,
            length=length,
            before_text=before_text,
            after_text=after_text,
            before_html=before_html,
            after_html=after_html,
            id=_opt_str(data, "id") or str(uuid.uuid4()),
            author=_opt_str(data, "author") or "You",
            created=_opt_str(data, "created") or utc_now(),
            status=status,
            conflicted=_opt_bool(data, "conflicted"),
            kind=kind,
        )


@dataclass
class TocRegion:
    start: int
    length: int

    def to_dict(self) -> dict:
        return {"start": self.start, "length": self.length}

    @classmethod
    def from_dict(cls, data: object) -> TocRegion | None:
        if data is None:
            return None
        if not isinstance(data, dict):
            raise ValidationError("TOC region must be an object")
        start = data.get("start")
        length = data.get("length")
        if not _is_position(start) or not _is_position(length):
            raise ValidationError("TOC region invalid")
        return cls(start=start, length=length)


@dataclass
class DocumentState:
    title: str = "Untitled"
    html: str = ""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    page: PageSettings = field(default_factory=PageSettings)
    comments: list[Comment] = field(default_factory=list)
    revisions: list[Revision] = field(default_factory=list)
    toc: TocRegion | None = None
    track_changes: bool = False
    schema_version: int = SCHEMA_VERSION
    design: dict = field(default_factory=dict)
    features: dict = field(default_factory=lambda: {"nodes": {}, "objects": [], "sections": []})
    review: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "schema_version": self.schema_version,
            "id": self.id,
            "title": self.title,
            "html": self.html,
            "page": self.page.to_dict(),
            "comments": [c.to_dict() for c in self.comments],
            "revisions": [r.to_dict() for r in self.revisions],
            "toc": self.toc.to_dict() if self.toc else None,
            "track_changes": self.track_changes,
            "design": self.design,
            "features": self.features,
            "review": self.review,
        }

    @classmethod
    def from_dict(cls, data: object) -> DocumentState:
        if not isinstance(data, dict):
            raise ValidationError("Document state must be an object")
        if data.get("schema_version") != SCHEMA_VERSION:
            raise ValidationError("Unsupported state schema version")
        html = data.get("html", "")
        if not isinstance(html, str):
            raise ValidationError("html must be a string")
        state = cls(
            id=_opt_str(data, "id") or str(uuid.uuid4()),
            title=_opt_str(data, "title") or "Untitled",
            html=html,
            page=PageSettings.from_dict(data.get("page") or {}),
            track_changes=_opt_bool(data, "track_changes"),
            schema_version=SCHEMA_VERSION,
        )
        from .document_features import checked_features
        state.features = checked_features(data.get("features"))
        state.review = checked_review_settings(data.get("review", {}))
        design = data.get("design", {})
        if not isinstance(design, dict):
            raise ValidationError("Design settings must be an object")
        if design:
            from .design import DesignSettings
            state.design = DesignSettings.from_dict(design).to_dict()
        comments_raw = data.get("comments") or []
        revisions_raw = data.get("revisions") or []
        if not isinstance(comments_raw, list) or not isinstance(revisions_raw, list):
            raise ValidationError("comments/revisions must be lists")
        state.comments = [Comment.from_dict(c) for c in comments_raw]
        state.revisions = [Revision.from_dict(r) for r in revisions_raw]
        state.toc = TocRegion.from_dict(data.get("toc"))
        bound = 1 << 26
        for comment in state.comments:
            if comment.end > bound:
                raise ValidationError("Comment range exceeds document bounds")
        for revision in state.revisions:
            if revision.start + revision.length > bound:
                raise ValidationError("Revision range exceeds document bounds")
        if state.toc and state.toc.start + state.toc.length > bound:
            raise ValidationError("TOC region exceeds document bounds")
        return state


def sanitize_ranges(state: DocumentState, char_count: int) -> None:
    bound = max(char_count, 0)
    for comment in state.comments:
        if comment.start > bound or comment.end > bound:
            comment.orphaned = True
            comment.start = min(comment.start, bound)
            comment.end = min(comment.end, bound)
            if comment.end < comment.start:
                comment.end = comment.start
    for revision in state.revisions:
        if revision.start > bound or revision.start + revision.length > bound:
            revision.conflicted = True
            revision.start = min(revision.start, bound)
            revision.length = min(revision.length, bound - revision.start)
    if state.toc is not None and state.toc.start + state.toc.length > bound:
        state.toc = None
    for lock in state.review.get("locks", []):
        lock["start"] = min(lock["start"], bound)
        lock["end"] = min(lock["end"], bound)


def checked_review_settings(data: object) -> dict:
    if not isinstance(data, dict):
        raise ValidationError("Review settings must be an object")
    result = {name: _opt_bool(data, name)
              for name in ("track_formatting", "formatting_locked") if name in data}
    locks = data.get("locks", [])
    if not isinstance(locks, list):
        raise ValidationError("Editing restrictions must be a list")
    if "locks" in data:
        result["locks"] = []
    for item in locks:
        if not isinstance(item, dict):
            raise ValidationError("Editing restriction must be an object")
        start, end = item.get("start"), item.get("end")
        if (not _is_position(start) or not _is_position(end)
                or end < start or end > (1 << 26)):
            raise ValidationError("Editing restriction range invalid")
        result["locks"].append({"start": start, "end": end,
                                "name": _opt_str(item, "name") or "Protected section",
                                "id": _opt_str(item, "id") or str(uuid.uuid4())})
    return result


def _opt_str(data: dict, key: str) -> str | None:
    value = data.get(key)
    return value if isinstance(value, str) else None


def _opt_bool(data: dict, key: str) -> bool:
    value = data.get(key, False)
    if not isinstance(value, bool):
        raise ValidationError(f"{key} must be a boolean")
    return value


def _is_position(value: object) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= 0
