"""AI Document Chat Assistant Service.

Provides context-aware analysis and conversational intelligence for
Word, Excel, PowerPoint, and Code documents. Supports both offline local
analysis engine and optional Google Gemini API generation when configured.
"""

from __future__ import annotations

import os
import re
from typing import Any
import httpx

from ..schemas.chat import ChatMessage, ChatResponse


class ChatService:
    """Intelligent chat assistant for documents."""

    def __init__(self, gemini_api_key: str | None = None) -> None:
        self.api_key = gemini_api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")

    async def generate_response(
        self,
        messages: list[ChatMessage],
        query: str | None,
        document: dict[str, Any] | None,
    ) -> ChatResponse:
        """Generate a response using Gemini if available, or the local engine."""
        latest_user_query = query
        if not latest_user_query:
            for m in reversed(messages):
                if m.role == "user":
                    latest_user_query = m.content
                    break
        if not latest_user_query:
            latest_user_query = "Hello"

        # Attempt Gemini API if key is present
        if self.api_key:
            try:
                gemini_reply = await self._call_gemini(latest_user_query, messages, document)
                if gemini_reply:
                    suggestions = self._build_suggestions(document, latest_user_query)
                    return ChatResponse(message=gemini_reply, suggestions=suggestions)
            except Exception as e:
                # Log error and fall back gracefully to local engine
                print(f"[ChatService] Gemini API call skipped or failed ({e}); falling back to local analyzer.")

        # Local contextual intelligence fallback
        return self._local_analysis(latest_user_query, document)

    async def _call_gemini(
        self,
        query: str,
        messages: list[ChatMessage],
        document: dict[str, Any] | None,
    ) -> str | None:
        """Call Google Gemini REST API."""
        doc_context = self._format_doc_summary(document)
        system_instruction = (
            "You are Office Studio Copilot, an expert document assistant embedded in a Word, "
            "Excel, and PowerPoint workspace. Assist the user with analyzing, reviewing, summarizing, "
            "and improving the active document. Provide clear, structured, professional answers with "
            "Markdown formatting where helpful.\n\n"
            f"ACTIVE DOCUMENT CONTEXT:\n{doc_context}"
        )

        contents = []
        # Add recent conversation history (up to last 6 turns)
        for m in messages[-6:]:
            role = "user" if m.role == "user" else "model"
            contents.append({"role": role, "parts": [{"text": m.content}]})

        if not contents or contents[-1].get("role") != "user":
            contents.append({"role": "user", "parts": [{"text": query}]})

        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={self.api_key}"
        payload = {
            "contents": contents,
            "systemInstruction": {"parts": [{"text": system_instruction}]},
            "generationConfig": {
                "temperature": 0.3,
                "maxOutputTokens": 1024,
            },
        }

        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.post(url, json=payload)
            if resp.status_code == 200:
                data = resp.json()
                candidates = data.get("candidates", [])
                if candidates and "content" in candidates[0]:
                    parts = candidates[0]["content"].get("parts", [])
                    if parts and "text" in parts[0]:
                        return parts[0]["text"]
        return None

    def _format_doc_summary(self, document: dict[str, Any] | None) -> str:
        if not document:
            return "No document is currently active."
        doc_type = document.get("type", "unknown")
        data = document.get("data", {})
        title = data.get("title", "Untitled")

        if doc_type == "word":
            headings = [h.get("text", "") for h in data.get("headings", [])]
            stats = data.get("stats", {})
            sample_text = " ".join([p.get("text", "") for p in data.get("paragraphs", [])[:8] if p.get("text")])
            return (
                f"Type: Word Document\nTitle: {title}\n"
                f"Headings: {', '.join(headings) if headings else 'None'}\n"
                f"Stats: {stats.get('words', 0)} words, {stats.get('paragraphs', 0)} paragraphs, readability: {stats.get('readabilityScore', 'N/A')}\n"
                f"Sample Content: {sample_text[:500]}"
            )
        elif doc_type == "excel":
            sheets = data.get("sheets", [])
            sheet_names = [s.get("name", "") for s in sheets]
            active_id = data.get("activeSheetId")
            active_sheet = next((s for s in sheets if s.get("id") == active_id), sheets[0] if sheets else None)
            sample_cells = []
            if active_sheet:
                for k, v in list(active_sheet.get("cells", {}).items())[:10]:
                    formula = f" ({v.get('formula')})" if v.get("formula") else ""
                    sample_cells.append(f"{k}: {v.get('value')}{formula}")
            return (
                f"Type: Excel Workbook\nTitle: {title}\n"
                f"Sheets: {', '.join(sheet_names)}\n"
                f"Active Sheet: {active_sheet.get('name') if active_sheet else 'N/A'}\n"
                f"Sample Cells: {', '.join(sample_cells)}"
            )
        elif doc_type == "powerpoint":
            slides = data.get("slides", [])
            slide_titles = [f"Slide {s.get('slideNumber')}: {s.get('title', '')}" for s in slides]
            return (
                f"Type: PowerPoint Presentation\nTitle: {title}\n"
                f"Total Slides: {len(slides)}\n"
                f"Slide Outline:\n" + "\n".join(slide_titles)
            )
        elif doc_type == "code":
            lang = data.get("language", "plaintext")
            content = data.get("content", "")
            return f"Type: Code Document\nTitle: {title}\nLanguage: {lang}\nContent Preview:\n{content[:400]}"
        return f"Document: {title} ({doc_type})"

    def _local_analysis(self, query: str, document: dict[str, Any] | None) -> ChatResponse:
        """Local heuristic & semantic response generator for offline execution."""
        q_lower = query.lower()

        if not document:
            return ChatResponse(
                message=(
                    "👋 **Hello! I'm your Code Office Studio Assistant.**\n\n"
                    "You can open or create a **Word document**, **Excel workbook**, or **PowerPoint presentation**, "
                    "and I will help you analyze data, inspect formulas, summarize sections, review presentation slides, "
                    "or draft new content.\n\n"
                    "How can I help you today?"
                ),
                suggestions=[
                    "Create a new document",
                    "How do I upload files?",
                    "What formats are supported?",
                ],
            )

        doc_type = document.get("type", "")
        data = document.get("data", {})
        title = data.get("title", "Untitled")

        # Word Document Processing
        if doc_type == "word":
            headings = data.get("headings", [])
            stats = data.get("stats", {})
            paragraphs = data.get("paragraphs", [])
            comments = data.get("comments", [])

            if any(k in q_lower for k in ("summar", "overview", "tl;dr", "explain", "about")):
                hd_list = "\n".join([f"- **Level {h.get('level')}**: {h.get('text')}" for h in headings]) if headings else "- No headings defined yet"
                sample_body = [p.get("text", "") for p in paragraphs if p.get("type") == "body" and p.get("text")]
                excerpt = sample_body[0] if sample_body else "No body text available."

                msg = (
                    f"### 📄 Document Summary: *{title}*\n\n"
                    f"**Structure & Outline:**\n{hd_list}\n\n"
                    f"**Key Metrics:**\n"
                    f"- **Words**: {stats.get('words', 0):,} ({stats.get('pages', 1)} estimated pages)\n"
                    f"- **Readability**: {stats.get('readabilityScore', 'Standard')}\n"
                    f"- **Reading Time**: ~{stats.get('readingTimeMinutes', 1)} min\n\n"
                    f"**First Paragraph Excerpt:**\n> {excerpt[:280]}..."
                )
                return ChatResponse(
                    message=msg,
                    suggestions=["Analyze readability", "Check document comments", "Suggest next heading", "Format recommendations"],
                )

            elif any(k in q_lower for k in ("readab", "stat", "metric", "word count", "pages")):
                msg = (
                    f"### 📊 Readability & Content Statistics for *{title}*\n\n"
                    f"- **Total Words**: `{stats.get('words', 0):,}`\n"
                    f"- **Total Characters**: `{stats.get('characters', 0):,}` (excluding spaces: `{stats.get('charactersNoSpaces', 0):,}`)\n"
                    f"- **Paragraphs Count**: `{stats.get('paragraphs', 0)}`\n"
                    f"- **Estimated Pages**: `{stats.get('pages', 1)}`\n"
                    f"- **Estimated Reading Time**: `{stats.get('readingTimeMinutes', 1)} minute(s)`\n"
                    f"- **Readability Level**: **{stats.get('readabilityScore', 'Good')}**\n\n"
                    f"💡 *Tip: Headings and shorter paragraphs improve the readability score.*"
                )
                return ChatResponse(
                    message=msg,
                    suggestions=["Summarize document", "Review comments", "Add executive summary"],
                )

            elif any(k in q_lower for k in ("comment", "review", "feedback")):
                if comments:
                    cm_list = "\n".join([
                        f"- **{c.get('author', 'Reviewer')}** on *\"{c.get('selectedText', '')}\"*: {c.get('comment')} "
                        f"({'✅ Resolved' if c.get('resolved') else '⏳ Pending'})"
                        for c in comments
                    ])
                    msg = f"### 💬 Document Comments ({len(comments)} total):\n\n{cm_list}"
                else:
                    msg = f"### 💬 Comments\n\nThere are currently no review comments in *{title}*. You can highlight text in the editor and click **Add Comment** to annotate."
                return ChatResponse(
                    message=msg,
                    suggestions=["Summarize document", "Analyze readability", "Check typography"],
                )

            # Default Word Response
            return ChatResponse(
                message=(
                    f"I'm reviewing your Word document **\"{title}\"** ({stats.get('words', 0)} words).\n\n"
                    f"You can ask me to:\n"
                    f"1. **Summarize** the document sections\n"
                    f"2. **Analyze readability** and statistics\n"
                    f"3. **Review comments** and action items\n"
                    f"4. **Suggest additions** or improve styling"
                ),
                suggestions=["Summarize document", "Analyze readability", "Check comments"],
            )

        # Excel Spreadsheet Processing
        elif doc_type == "excel":
            sheets = data.get("sheets", [])
            active_id = data.get("activeSheetId")
            active_sheet = next((s for s in sheets if s.get("id") == active_id), sheets[0] if sheets else None)
            sheet_name = active_sheet.get("name", "Sheet1") if active_sheet else "Sheet1"
            cells = active_sheet.get("cells", {}) if active_sheet else {}

            formula_cells = {k: v for k, v in cells.items() if v.get("formula")}
            numeric_vals = [v.get("value") for v in cells.values() if isinstance(v.get("value"), (int, float))]

            if any(k in q_lower for k in ("formula", "calc", "sum", "average", "math", "function")):
                if formula_cells:
                    f_list = "\n".join([f"- **Cell {k}**: `{v.get('formula')}` ➔ evaluated value: `{v.get('value')}`" for k, v in list(formula_cells.items())[:12]])
                    msg = (
                        f"### 🧮 Formulas in Worksheet *\"{sheet_name}\"*\n\n"
                        f"Found **{len(formula_cells)}** calculated formula cell(s):\n{f_list}\n\n"
                        f"**Supported Functions**: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `ABS`, `ROUND`, basic arithmetic `+ - * / %`, and cross-sheet references (e.g. `Sheet2!A1`)."
                    )
                else:
                    msg = (
                        f"### 🧮 Formulas in *\"{sheet_name}\"*\n\n"
                        f"No formula cells found in this worksheet yet. To add a formula, type `=` in any cell, e.g.:\n"
                        f"- `=SUM(A1:A10)`\n"
                        f"- `=AVERAGE(B2:B20)`\n"
                        f"- `=ROUND(C1 * 1.15, 2)`"
                    )
                return ChatResponse(
                    message=msg,
                    suggestions=["Analyze numeric trends", "List all worksheets", "Suggest formula"],
                )

            elif any(k in q_lower for k in ("trend", "sum", "total", "stat", "data", "average", "metric")):
                if numeric_vals:
                    total = sum(numeric_vals)
                    avg = total / len(numeric_vals)
                    min_val = min(numeric_vals)
                    max_val = max(numeric_vals)
                    msg = (
                        f"### 📈 Data Metrics for *\"{sheet_name}\"*\n\n"
                        f"- **Numeric Data Points**: `{len(numeric_vals)}`\n"
                        f"- **Sum Total**: `{total:,.2f}`\n"
                        f"- **Average**: `{avg:,.2f}`\n"
                        f"- **Minimum**: `{min_val:,.2f}`\n"
                        f"- **Maximum**: `{max_val:,.2f}`\n"
                        f"- **Total Sheets in Workbook**: `{len(sheets)}`"
                    )
                else:
                    msg = f"Worksheet *\"{sheet_name}\"* does not contain numeric data cells yet."
                return ChatResponse(
                    message=msg,
                    suggestions=["Explain formulas", "How to sort data", "Export to XLSX"],
                )

            # Default Excel Response
            return ChatResponse(
                message=(
                    f"I'm inspecting your workbook **\"{title}\"** (Sheet: *{sheet_name}*, {len(cells)} populated cells).\n\n"
                    f"Ask me to:\n"
                    f"- **Explain formulas** currently in the sheet\n"
                    f"- **Analyze numbers**, sums, and averages\n"
                    f"- **Suggest spreadsheet formulas** for your data\n"
                    f"- **Guide formatting** and sorting"
                ),
                suggestions=["Explain formulas", "Analyze numbers", "Workbook structure"],
            )

        # PowerPoint Presentation Processing
        elif doc_type == "powerpoint":
            slides = data.get("slides", [])
            active_id = data.get("activeSlideId")
            active_slide = next((s for s in slides if s.get("id") == active_id), slides[0] if slides else None)
            slide_num = active_slide.get("slideNumber", 1) if active_slide else 1

            if any(k in q_lower for k in ("outline", "slide", "overview", "summar", "presentation")):
                s_list = "\n".join([
                    f"- **Slide {s.get('slideNumber')}** ({s.get('layout', 'content')}): *{s.get('title', 'Untitled')}* — {len(s.get('objects', []))} visual object(s)"
                    for s in slides
                ])
                msg = (
                    f"### 📽️ Presentation Outline: *\"{title}\"*\n\n"
                    f"**Theme**: `{data.get('themeName', 'Standard')}` (Accent: `{data.get('accentColor', '#0078d4')}`)\n"
                    f"**Total Slides**: `{len(slides)}`\n\n"
                    f"{s_list}\n\n"
                    f"💡 *Press F5 in the editor to enter full-screen Presentation Mode.*"
                )
                return ChatResponse(
                    message=msg,
                    suggestions=["Review speaker notes", "Active slide details", "Suggest slide layout"],
                )

            elif any(k in q_lower for k in ("note", "speak", "script")):
                notes = active_slide.get("notes", "") if active_slide else ""
                if notes:
                    msg = f"### 🎙️ Speaker Notes for Slide {slide_num} (*{active_slide.get('title')}*):\n\n> {notes}"
                else:
                    msg = (
                        f"### 🎙️ Speaker Notes for Slide {slide_num}\n\n"
                        f"No speaker notes recorded for this slide yet. You can click the **Notes** panel at the bottom of the slide canvas to add presenter reminders."
                    )
                return ChatResponse(
                    message=msg,
                    suggestions=["Presentation outline", "Slide layout tips", "Present slides (F5)"],
                )

            # Default PowerPoint Response
            return ChatResponse(
                message=(
                    f"I'm looking at your presentation **\"{title}\"** ({len(slides)} slides, currently on Slide {slide_num}).\n\n"
                    f"I can help you with:\n"
                    f"- **Presentation outlines** and slide structure\n"
                    f"- **Speaker notes** and presentation scripts\n"
                    f"- **Design recommendations** and transitions"
                ),
                suggestions=["Presentation outline", "Speaker notes", "Slide details"],
            )

        # Code / Text File Processing
        elif doc_type == "code":
            lang = data.get("language", "plaintext")
            content = data.get("content", "")
            lines = content.splitlines()
            return ChatResponse(
                message=(
                    f"### 💻 Code File: *\"{title}\"*\n\n"
                    f"- **Language**: `{lang}`\n"
                    f"- **Lines of Code**: `{len(lines)}`\n"
                    f"- **Characters**: `{len(content):,}`\n\n"
                    f"You can edit code with syntax highlighting, search/replace across files, and inspect JSON or symbols."
                ),
                suggestions=["Analyze symbols", "Code formatting tips", "Download file"],
            )

        # Generic Response
        return ChatResponse(
            message=f"I have loaded **\"{title}\"**. How can I assist you with this document?",
            suggestions=["Summarize document", "Export document", "Show help"],
        )

    def _build_suggestions(self, document: dict[str, Any] | None, query: str) -> list[str]:
        if not document:
            return ["Create new document", "Upload document", "View help"]
        doc_type = document.get("type", "")
        if doc_type == "word":
            return ["Summarize document", "Analyze readability", "Check comments", "Add heading"]
        elif doc_type == "excel":
            return ["Explain formulas", "Analyze numbers", "Add SUM formula", "List sheets"]
        elif doc_type == "powerpoint":
            return ["Presentation outline", "Speaker notes", "Slide layout suggestions"]
        elif doc_type == "code":
            return ["Analyze code structure", "Find functions", "Format file"]
        return ["Summarize", "Explain structure", "Export options"]

