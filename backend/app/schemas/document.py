"""Document schemas and validation rules matching Code Office Studio."""

from __future__ import annotations

import re
from typing import Any, Literal
from pydantic import BaseModel, Field, model_validator


ID_PATTERN = re.compile(r"^[a-zA-Z0-9][a-zA-Z0-9_-]{0,127}$")
COORD_PATTERN = re.compile(r"^([A-Z]{1,3})([1-9][0-9]{0,6})$")


class ValidationError(ValueError):
    """Raised when document data fails validation."""
    def __init__(self, message: str, status_code: int = 400):
        super().__init__(message)
        self.status_code = status_code


def validate_document_model(doc_type: str, data: dict[str, Any], doc_id: str) -> None:
    """Validate document structure matching original-project server/index.mjs."""
    if not isinstance(data, dict):
        raise ValidationError("A document model is required.")

    if data.get("id") != doc_id or not isinstance(data.get("title"), str):
        raise ValidationError("Document ID must match the URL, and title must be text.")

    if doc_type == "word":
        for field in ("author", "createdAt", "modifiedAt"):
            if not isinstance(data.get(field), str):
                raise ValidationError("Invalid Word document metadata.")

        headings = data.get("headings")
        if not isinstance(headings, list):
            raise ValidationError("Invalid Word outline.")
        for h in headings:
            if not isinstance(h, dict) or not isinstance(h.get("id"), str) or not isinstance(h.get("text"), str):
                raise ValidationError("Invalid Word outline.")
            if h.get("level") not in (1, 2, 3, 4) or not isinstance(h.get("pageIndex"), (int, float)):
                raise ValidationError("Invalid Word outline.")

        sections = data.get("sections")
        if not isinstance(sections, list):
            raise ValidationError("Invalid Word section.")
        for s in sections:
            if not isinstance(s, dict) or not isinstance(s.get("id"), str) or not isinstance(s.get("title"), str):
                raise ValidationError("Invalid Word section.")
            if s.get("orientation") not in ("portrait", "landscape") or not isinstance(s.get("pageNumber"), (int, float)):
                raise ValidationError("Invalid Word section.")

        comments = data.get("comments")
        if not isinstance(comments, list):
            raise ValidationError("Invalid Word comment.")
        for c in comments:
            if not isinstance(c, dict):
                raise ValidationError("Invalid Word comment.")
            for cf in ("id", "author", "avatarColor", "timestamp", "selectedText", "comment"):
                if not isinstance(c.get(cf), str):
                    raise ValidationError("Invalid Word comment.")
            if not isinstance(c.get("resolved"), bool):
                raise ValidationError("Invalid Word comment.")

        typography = data.get("typography")
        stats = data.get("stats")
        if not isinstance(typography, dict) or not isinstance(stats, dict):
            raise ValidationError("Word typography and statistics are required.")
        if not isinstance(typography.get("fontFamily"), str) or not isinstance(typography.get("textColor"), str):
            raise ValidationError("Word typography and statistics are required.")
        if not isinstance(typography.get("fontSize"), (int, float)) or typography.get("fontSize", 0) <= 0:
            raise ValidationError("Word typography and statistics are required.")
        if not isinstance(typography.get("lineHeight"), (int, float)) or typography.get("lineHeight", 0) <= 0:
            raise ValidationError("Word typography and statistics are required.")
        if typography.get("marginSize") not in ("normal", "narrow", "wide"):
            raise ValidationError("Word typography and statistics are required.")

        for sf in ("words", "characters", "charactersNoSpaces", "paragraphs", "pages", "readingTimeMinutes"):
            if not isinstance(stats.get(sf), (int, float)):
                raise ValidationError("Word typography and statistics are required.")
        if not isinstance(stats.get("readabilityScore"), str):
            raise ValidationError("Word typography and statistics are required.")

        paragraphs = data.get("paragraphs")
        if not isinstance(paragraphs, list):
            raise ValidationError("Invalid paragraphs.")
        p_ids = set()
        for p in paragraphs:
            if not isinstance(p, dict) or not isinstance(p.get("id"), str):
                raise ValidationError("Invalid paragraph.")
            if p["id"] in p_ids:
                raise ValidationError("Paragraph IDs must be unique.")
            p_ids.add(p["id"])

            p_type = p.get("type")
            if p_type not in ("heading-1", "heading-2", "heading-3", "body", "quote", "callout", "table", "smartart"):
                raise ValidationError("Invalid paragraph.")
            if "text" in p and not isinstance(p["text"], str):
                raise ValidationError("Paragraph text must be text.")
            if "runs" in p:
                if not isinstance(p["runs"], list) or any(not isinstance(r, dict) or not isinstance(r.get("text"), str) for r in p["runs"]):
                    raise ValidationError("Invalid text runs.")
            if "tableData" in p:
                if not isinstance(p["tableData"], list) or any(
                    not isinstance(row, list) or any(not isinstance(cell, str) for cell in row)
                    for row in p["tableData"]
                ):
                    raise ValidationError("Invalid table.")
            if (p_type == "smartart") != ("smartArt" in p):
                raise ValidationError("Invalid SmartArt diagram.")
            if p_type == "smartart":
                sa = p["smartArt"]
                if not isinstance(sa, dict) or not isinstance(sa.get("title"), str):
                    raise ValidationError("Invalid SmartArt diagram.")
                if sa.get("layout") not in ("process", "cycle", "hierarchy", "pyramid"):
                    raise ValidationError("Invalid SmartArt diagram.")
                if not isinstance(sa.get("accentColor"), str) or not re.match(r"^#[a-fA-F0-9]{6}$", sa["accentColor"]):
                    raise ValidationError("Invalid SmartArt diagram.")
                items = sa.get("items")
                if not isinstance(items, list) or not (1 <= len(items) <= 12):
                    raise ValidationError("Invalid SmartArt diagram.")
                sa_ids = set()
                for item in items:
                    if not isinstance(item, dict) or not isinstance(item.get("id"), str) or not isinstance(item.get("text"), str) or len(item["text"]) > 500:
                        raise ValidationError("Invalid SmartArt diagram.")
                    if item["id"] in sa_ids:
                        raise ValidationError("Invalid SmartArt diagram.")
                    sa_ids.add(item["id"])

    elif doc_type == "excel":
        if not isinstance(data.get("activeSheetId"), str):
            raise ValidationError("An active worksheet is required.")
        sheets = data.get("sheets")
        if not isinstance(sheets, list) or len(sheets) == 0 or len(sheets) > 50:
            raise ValidationError("Invalid worksheet or grid exceeds 1,000 rows, 100 columns, or 10,000 cells.")
        sheet_ids = set()
        for sheet in sheets:
            if not isinstance(sheet, dict) or not isinstance(sheet.get("id"), str) or not isinstance(sheet.get("name"), str):
                raise ValidationError("Invalid worksheet.")
            sheet_ids.add(sheet["id"])
            row_count = sheet.get("rowCount")
            col_count = sheet.get("colCount")
            cells = sheet.get("cells")
            columns = sheet.get("columns")
            if (
                not isinstance(row_count, int) or not isinstance(col_count, int)
                or row_count < 1 or row_count > 1000 or col_count < 1 or col_count > 100
                or (row_count * col_count) > 10000
                or not isinstance(cells, dict) or len(cells) > 10000
                or not isinstance(columns, list) or len(columns) != col_count
            ):
                raise ValidationError("Invalid worksheet or grid exceeds 1,000 rows, 100 columns, or 10,000 cells.")

            for col in columns:
                if not isinstance(col, dict) or not isinstance(col.get("key"), str) or not re.match(r"^[A-Z]{1,3}$", col["key"]) or not isinstance(col.get("width"), (int, float)):
                    raise ValidationError("Invalid worksheet columns.")

            for coord, cell in cells.items():
                m = COORD_PATTERN.match(coord)
                if not m or not isinstance(cell, dict):
                    raise ValidationError("Invalid worksheet cell.")
                val = cell.get("value")
                if val is not None and not isinstance(val, (str, int, float, bool)):
                    raise ValidationError("Invalid worksheet cell.")
                if "formula" in cell and not isinstance(cell["formula"], str):
                    raise ValidationError("Invalid worksheet cell.")
                letters, row_str = m.groups()
                col_idx = 0
                for char in letters:
                    col_idx = col_idx * 26 + (ord(char) - 64)
                if col_idx > col_count or int(row_str) > row_count:
                    raise ValidationError("A worksheet cell is outside the visible grid.")

        if len(sheet_ids) != len(sheets) or data["activeSheetId"] not in sheet_ids:
            raise ValidationError("An active worksheet and unique worksheet IDs are required.")

    elif doc_type == "powerpoint":
        if not isinstance(data.get("activeSlideId"), str):
            raise ValidationError("An active slide is required.")
        if not isinstance(data.get("themeName"), str) or not isinstance(data.get("accentColor"), str):
            raise ValidationError("Invalid presentation theme.")
        slides = data.get("slides")
        if not isinstance(slides, list) or len(slides) == 0 or len(slides) > 200:
            raise ValidationError("An active slide and unique slide IDs are required.")
        slide_ids = set()
        for slide in slides:
            if not isinstance(slide, dict) or not isinstance(slide.get("id"), str):
                raise ValidationError("Invalid slide.")
            slide_ids.add(slide["id"])
            for sf in ("title", "background", "notes"):
                if not isinstance(slide.get(sf), str):
                    raise ValidationError("Invalid slide.")
            if not isinstance(slide.get("slideNumber"), (int, float)):
                raise ValidationError("Invalid slide.")
            if slide.get("transition") not in ("none", "fade", "slide", "zoom"):
                raise ValidationError("Invalid slide.")
            if slide.get("layout") not in ("title", "content", "two-column", "dashboard", "blank"):
                raise ValidationError("Invalid slide.")

            objects = slide.get("objects")
            if not isinstance(objects, list) or len(objects) > 1000:
                raise ValidationError("Invalid slide.")
            obj_ids = set()
            for obj in objects:
                if not isinstance(obj, dict) or not isinstance(obj.get("id"), str):
                    raise ValidationError("Invalid slide object.")
                if obj["id"] in obj_ids:
                    raise ValidationError("Slide object IDs must be unique.")
                obj_ids.add(obj["id"])
                for f in ("fill", "color"):
                    if not isinstance(obj.get(f), str):
                        raise ValidationError("Invalid slide object.")
                for num_f in ("x", "y", "width", "height", "fontSize", "zIndex"):
                    if not isinstance(obj.get(num_f), (int, float)):
                        raise ValidationError("Invalid slide object.")
                if obj.get("kind") not in ("text", "shape", "metric", "image", "chart"):
                    raise ValidationError("Invalid slide object.")

        if len(slide_ids) != len(slides) or data["activeSlideId"] not in slide_ids:
            raise ValidationError("An active slide and unique slide IDs are required.")

    elif doc_type == "code":
        if not isinstance(data.get("content"), str) or data.get("language") not in ("typescript", "javascript", "python", "markdown", "json"):
            raise ValidationError("Invalid text document.")
    else:
        raise ValidationError("Unsupported document type.")


class DocumentPayload(BaseModel):
    name: str = Field(..., min_length=1, max_length=240)
    revision: int | None = None
    document: dict[str, Any]

    @model_validator(mode="after")
    def validate_name(self) -> DocumentPayload:
        if any(c in self.name for c in ("\x00", "\x1f", "\\", "/")):
            raise ValueError("Filename contains invalid characters.")
        return self


class DocumentRecord(BaseModel):
    id: str
    name: str
    revision: int
    updatedAt: str
    document: dict[str, Any]


class DocumentEnvelope(BaseModel):
    schemaVersion: int = 1
    head: DocumentRecord
    versions: list[DocumentRecord] = []

