from __future__ import annotations

import json
import platform
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

from PySide6.QtGui import QTextCursor, QTextDocument
from PySide6.QtWidgets import (QComboBox, QDialog, QDialogButtonBox, QFileDialog,
                              QFormLayout, QHBoxLayout, QLabel, QLineEdit,
                              QListWidget, QMessageBox, QPlainTextEdit,
                              QPushButton, QTextBrowser, QVBoxLayout)

from . import __version__
from .docx_io import atomic_write
from .editor import STYLE_ORDER


HELP_TOPICS = [
    ("Start a document", "File > New opens the template gallery. Choose Blank, then type. "
     "Save with Ctrl+S to a .docx file on your computer. No account is required."),
    ("Styles and navigation", "Apply Heading 1, 2, or 3 with Ctrl+1, Ctrl+2, or Ctrl+3. "
     "Ctrl+0 restores Normal. View > Navigation shows headings, pages and objects. "
     "References > TOC builds contents from those headings."),
    ("Sources and references", "References provides notes, a local source library, citations, "
     "bibliographies, captions, lists of figures, cross-references, indexes and authorities. "
     "Generated content updates as you edit. Review citation details against your required style guide."),
    ("Mail merge", "Mailings opens local CSV or XLSX recipient data. Map columns, choose active "
     "recipients, insert fields or address and greeting blocks, and preview before exporting. "
     "Conditional rules select text or skip recipients. Email output is saved as local drafts; it is not sent."),
    ("Review and accessibility", "Review offers spelling and grammar heuristics, word counts, "
     "a small English thesaurus, accessibility checks, comments and tracked edits. "
     "Accessibility findings assist human review; they do not certify WCAG conformance. "
     "Local author names identify comments; Folio has no account system."),
    ("Speech and translation", "Read Aloud uses installed system voices when supported. "
     "Translation uses an explicitly imported local glossary. It replaces known phrases; "
     "it is not machine translation and does not send document text anywhere."),
    ("Document views", "View offers Print Layout, Web Layout, Read Mode, Outline Mode and Draft Mode. "
     "Web and Draft views edit the same active document. Read Mode lays out pages horizontally. "
     "New View and Split share the document and undo history. Sync scrolling follows the relative "
     "vertical position. Closing a view leaves the document open."),
    ("JavaScript macros", "View > Macros runs manually invoked JavaScript using a small document API. "
     "folio.text and folio.selection read the starting text. folio.insertText(text), "
     "folio.appendText(text), folio.replaceAll(find, replacement), and folio.applyStyle(name) "
     "queue changes. They are applied together after successful evaluation. "
     "Macros have no file, network, or application automation API. VBA is not supported."),
    ("Recovery and history", "Unsaved changes are recovered from local SQLite storage. "
     "File > History lists document versions. Explicit saves mark recovery clean. "
     "Use Save As for an independent file; secondary views share the active document."),
    ("Printing and layout", "Layout sets paper, margins, columns and sections. Print Preview and "
     "PDF use the publishing layout, including columns and page fields. The main editing canvas "
     "uses a single page-width text column."),
    ("Support and feedback", "Help > Feedback creates a structured local bug report, feature request "
     "or support request. Export it and send it through your own support channel. "
     "No support service, telemetry, contact API or release feed is configured."),
    ("Keyboard shortcuts", "Ctrl+S Save · Ctrl+O Open · Ctrl+N New · Ctrl+P Print · "
     "Ctrl+F Find · Ctrl+H Replace · Ctrl+Enter Page break · Ctrl+Z Undo · Ctrl+Y Redo · "
     "Esc Exit focus mode."),
]

TUTORIALS = {
    "Write and save": ["File > New > Blank starts a new document.",
                       "Type a title and choose Home > Title.",
                       "Add a paragraph and use Home > Normal.",
                       "Press Ctrl+S and choose a .docx path."],
    "Build an academic document": ["Mark chapter titles as Heading 1 and subsections as Heading 2.",
                                  "Use References to add sources and insert citations.",
                                  "Insert notes and numbered captions beside the relevant text.",
                                  "Insert a bibliography and table of contents; inspect Print Preview."],
    "Create a mailing": ["Prepare a CSV or XLSX with one recipient per row and unique column headings.",
                         "Choose Mailings and load recipients; filter or exclude records as needed.",
                         "Insert merge fields, an address block and a greeting line.",
                         "Preview multiple recipients, then export a batch or compound document."],
    "Review a draft": ["Set your local author name in Review.",
                       "Turn on Track Changes before editing, and enable Track Formatting if needed.",
                       "Select text and add a comment; use threads to reply or resolve.",
                       "Review pending changes, then accept or reject them and save."],
}

RELEASE_NOTES = """Folio — bundled release notes

References: local sources and citation styles, notes, captions, cross-references and generated lists.
Mailings: recipient management, semantic fields, rules, envelope and label layouts, local output batches.
Review: comment threads, format tracking, review projections, accessibility checks and local speech tools.
View: live document views, split panes, page fitting, JavaScript document macros and local properties.
Help: searchable documentation, guided tutorials and structured feedback export.

These notes ship with this installation. Folio does not fetch a release feed.
Folio is a Windows application; no mobile client or mobile onboarding link is published.
"""


def search_help(query):
    terms = query.casefold().split()
    return [(title, body) for title, body in HELP_TOPICS
            if all(term in (title + " " + body).casefold() for term in terms)]


def feedback_record(kind, title, detail, contact=""):
    if kind not in ("Bug report", "Feature request", "Support request"):
        raise ValueError("Choose a feedback category.")
    if not title.strip() or not detail.strip():
        raise ValueError("Enter a title and description.")
    if len(title) > 500 or len(detail) > 100000 or len(contact) > 500:
        raise ValueError("Feedback exceeds the supported length.")
    return {"schema": 1, "category": kind, "title": title.strip(),
            "description": detail.strip(), "contact": contact.strip(),
            "created_at": datetime.now(timezone.utc).isoformat(),
            "folio_version": __version__, "platform": platform.system(),
            "delivery": "local export"}


def document_properties(owner):
    values = {"Title": owner._title, "Folio version": __version__,
              "Document ID": owner.state.id, "Storage": "Local files and SQLite",
              "Path": str(owner._path) if owner._path else "Unsaved",
              "Pending changes": str(sum(r.status == "pending" for r in owner.state.revisions)),
              "Comments": str(len(owner.state.comments))}
    if owner._path and owner._path.exists():
        info = owner._path.stat()
        values["File size"] = f"{info.st_size:,} bytes"
        values["Modified (UTC)"] = datetime.fromtimestamp(info.st_mtime, timezone.utc).isoformat()
        for folder in owner._path.resolve().parents:
            metadata = folder / ".git"
            if not metadata.is_dir():
                continue
            head = metadata / "HEAD"
            if head.is_file() and head.stat().st_size < 4096:
                value = head.read_text(encoding="utf-8").strip()
                values["Repository"] = str(folder)
                values["Git HEAD"] = value
            break
    return values


def plan_macro(source, text, selection="", timeout=3.0):
    if not isinstance(source, str) or len(source) > 100000:
        raise ValueError("Macro source must be at most 100,000 characters.")
    program = """(function() {
        const operations = [];
        function add(kind, args) {
            if (operations.length >= 1000) throw new Error('Too many operations');
            if (args.some(value => typeof value !== 'string' || value.length > 1000000))
                throw new Error('Arguments must be strings of at most 1,000,000 characters');
            operations.push([kind, ...args]);
        }
        const folio = Object.freeze({
            text: TEXT_VALUE, selection: SELECTION_VALUE,
            insertText: text => add('insert', [text]),
            appendText: text => add('append', [text]),
            replaceAll: (find, replacement) => add('replace', [find, replacement]),
            applyStyle: name => add('style', [name])
        });
        (function(folio) { USER_SOURCE }).call(undefined, folio);
        return JSON.stringify(operations);
    })()"""
    program = program.replace("USER_SOURCE", source, 1)
    program = program.replace("SELECTION_VALUE", json.dumps(selection), 1)
    program = program.replace("TEXT_VALUE", json.dumps(text), 1)
    try:
        result = subprocess.run([sys.executable, "-m", "folio.macro_worker"],
                                input=program, text=True, encoding="utf-8",
                                capture_output=True, timeout=timeout,
                                cwd=Path(__file__).resolve().parents[1],
                                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except subprocess.TimeoutExpired as error:
        raise ValueError("Macro exceeded its execution time limit.") from error
    except OSError as error:
        raise ValueError(f"Could not start the JavaScript engine: {error}") from error
    if result.returncode:
        raise ValueError("The JavaScript engine failed to run.")
    try:
        payload = json.loads(result.stdout)
        if payload.get("error"):
            raise ValueError(payload["error"])
        operations = json.loads(payload["result"])
    except (KeyError, json.JSONDecodeError) as error:
        raise ValueError("The macro did not produce valid operations.") from error
    if not isinstance(operations, list) or len(operations) > 1000:
        raise ValueError("Invalid macro operations.")
    total = 0
    for operation in operations:
        if (not isinstance(operation, list) or not operation
                or operation[0] not in ("insert", "append", "replace", "style")
                or len(operation) != (3 if operation[0] == "replace" else 2)
                or any(not isinstance(value, str) for value in operation)):
            raise ValueError("Invalid macro operation.")
        if operation[0] == "style" and operation[1] not in STYLE_ORDER:
            raise ValueError("Unknown paragraph style: " + operation[1])
        if operation[0] == "replace" and not operation[1]:
            raise ValueError("Replace requires nonempty search text.")
        total += sum(len(value) for value in operation)
    if total > 2_000_000:
        raise ValueError("Macro output exceeds the size limit.")
    return operations


def apply_macro(owner, operations):
    if not owner._can_edit() or owner.editor.isReadOnly():
        raise ValueError("The document is not editable.")
    document = owner.editor.document()
    if not owner.tracker.allows_edit(0, document.characterCount() - 1,
                                    any(operation[0] == "style" for operation in operations)):
        raise ValueError("Remove editing restrictions before running a document macro.")
    projected = document.toPlainText()
    for operation in operations:
        kind, *args = operation
        if kind == "replace":
            projected = projected.replace(*args)
        elif kind in ("insert", "append"):
            projected += args[0]
        if len(projected) > 2_000_000:
            raise ValueError("Macro would exceed the document size limit.")
    edit = owner.editor.textCursor()
    edit.beginEditBlock()
    try:
        for kind, *args in operations:
            if kind == "style":
                owner.editor.apply_style(args[0])
            elif kind == "insert":
                owner.editor.insertPlainText(args[0])
            elif kind == "append":
                saved = owner.editor.textCursor()
                cursor = QTextCursor(document)
                cursor.movePosition(QTextCursor.MoveOperation.End)
                owner.editor.setTextCursor(cursor)
                owner.editor.insertPlainText(args[0])
                owner.editor.setTextCursor(saved)
            else:
                matches = []
                cursor = QTextCursor(document)
                while True:
                    cursor = document.find(args[0], cursor, QTextDocument.FindFlag.FindCaseSensitively)
                    if cursor.isNull():
                        break
                    matches.append(QTextCursor(cursor))
                for cursor in reversed(matches):
                    cursor.insertText(args[1])
    finally:
        edit.endEditBlock()


class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Folio Help")
        self.resize(850, 600)
        layout = QVBoxLayout(self)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search documentation")
        layout.addWidget(self.search)
        row = QHBoxLayout()
        self.topics = QListWidget()
        self.body = QTextBrowser()
        row.addWidget(self.topics, 1)
        row.addWidget(self.body, 2)
        layout.addLayout(row)
        self.results = []
        self.search.textChanged.connect(self.find)
        self.topics.currentRowChanged.connect(self.show_topic)
        self.find("")

    def find(self, query):
        self.results = search_help(query)
        self.topics.clear()
        self.topics.addItems([title for title, _body in self.results])
        if self.results:
            self.topics.setCurrentRow(0)
        else:
            self.body.setPlainText("No topics match. Try a shorter search.")

    def show_topic(self, row):
        if 0 <= row < len(self.results):
            title, body = self.results[row]
            self.body.setPlainText(title + "\n\n" + body)


class TutorialDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Folio Tutorials")
        self.resize(650, 350)
        layout = QVBoxLayout(self)
        self.course = QComboBox()
        self.course.addItems(TUTORIALS)
        layout.addWidget(self.course)
        self.body = QLabel()
        self.body.setWordWrap(True)
        layout.addWidget(self.body, 1)
        row = QHBoxLayout()
        self.previous = QPushButton("Previous")
        self.next = QPushButton("Next")
        row.addWidget(self.previous)
        row.addWidget(self.next)
        layout.addLayout(row)
        self.index = 0
        self.previous.clicked.connect(lambda: self.move(-1))
        self.next.clicked.connect(lambda: self.move(1))
        self.course.currentTextChanged.connect(self.reset)
        self.render()

    def reset(self):
        self.index = 0
        self.render()

    def move(self, delta):
        self.index += delta
        self.render()

    def render(self):
        steps = TUTORIALS[self.course.currentText()]
        self.body.setText(f"Step {self.index + 1} of {len(steps)}\n\n{steps[self.index]}")
        self.previous.setEnabled(self.index > 0)
        self.next.setEnabled(self.index < len(steps) - 1)


class FeedbackDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Export Feedback")
        self.resize(650, 500)
        layout = QVBoxLayout(self)
        note = QLabel("Create a local report to send through your support channel. "
                      "No document content is collected automatically.")
        note.setWordWrap(True)
        layout.addWidget(note)
        form = QFormLayout()
        self.kind = QComboBox()
        self.kind.addItems(["Bug report", "Feature request", "Support request"])
        self.title = QLineEdit()
        self.contact = QLineEdit()
        self.detail = QPlainTextEdit()
        form.addRow("Category", self.kind)
        form.addRow("Title", self.title)
        form.addRow("Contact (optional)", self.contact)
        form.addRow("Description / steps", self.detail)
        layout.addLayout(form)
        export = QPushButton("Export report…")
        export.clicked.connect(self.export)
        layout.addWidget(export)

    def export(self):
        try:
            data = feedback_record(self.kind.currentText(), self.title.text(),
                                   self.detail.toPlainText(), self.contact.text())
            filename, _filter = QFileDialog.getSaveFileName(self, "Export feedback", "folio-feedback.json", "JSON (*.json)")
            if filename:
                payload = json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8")
                atomic_write(Path(filename), payload)
                self.accept()
        except (ValueError, OSError) as error:
            QMessageBox.warning(self, "Feedback", str(error))


class MacroDialog(QDialog):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner
        self.setWindowTitle("JavaScript Document Macro")
        self.resize(800, 550)
        layout = QVBoxLayout(self)
        label = QLabel("folio.text · folio.selection · folio.insertText(text) · folio.appendText(text)\n"
                       "folio.replaceAll(find, replacement) · folio.applyStyle(name)")
        layout.addWidget(label)
        self.source = QPlainTextEdit('folio.replaceAll("old phrase", "new phrase");')
        layout.addWidget(self.source)
        self.result = QLabel("Changes apply as one undo step after successful evaluation.")
        self.result.setWordWrap(True)
        layout.addWidget(self.result)
        row = QHBoxLayout()
        for text, slot in (("Load .js…", self.load), ("Save .js…", self.save), ("Run", self.run)):
            button = QPushButton(text)
            button.clicked.connect(slot)
            row.addWidget(button)
        layout.addLayout(row)

    def load(self):
        path, _filter = QFileDialog.getOpenFileName(self, "Load macro", "", "JavaScript (*.js)")
        if path:
            try:
                if Path(path).stat().st_size > 400000:
                    raise ValueError("Macro file exceeds the size limit.")
                self.source.setPlainText(Path(path).read_text(encoding="utf-8-sig"))
            except (OSError, ValueError) as error:
                self.result.setText(str(error))

    def save(self):
        path, _filter = QFileDialog.getSaveFileName(self, "Save macro", "macro.js", "JavaScript (*.js)")
        if path:
            try:
                atomic_write(Path(path), self.source.toPlainText().encode("utf-8"))
            except OSError as error:
                self.result.setText(str(error))

    def run(self):
        try:
            operations = plan_macro(self.source.toPlainText(), self.owner.editor.toPlainText(),
                                    self.owner.editor.textCursor().selectedText())
            apply_macro(self.owner, operations)
            self.result.setText(f"Applied {len(operations)} operations. Ctrl+Z undoes this macro.")
        except (ValueError, RuntimeError) as error:
            self.result.setText(str(error))
