# Folio

Folio is a native Windows word processor written in Python 3.12 with PySide6.
It is local-first: documents are real `.docx` files on disk, recovery and
version history live in a local SQLite store, and no account or network
connection is ever required.

## Bootstrap

Requires Python 3.12 (the pinned interpreter is
`C:\Users\zachn\AppData\Local\Python\pythoncore-3.12-64\python.exe`).

```bat
py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

Pinned dependencies (`requirements.txt`):

- PySide6==6.8.3
- python-docx==1.2.0
- pyspellchecker==0.8.3
- openpyxl==3.1.5

Development extras (`requirements-dev.txt`): pytest==8.3.5.

## Run

```bat
run.bat            :: bootstraps .venv if missing, then launches Folio
.venv\Scripts\python.exe main.py
.venv\Scripts\python.exe main.py path\to\file.docx
```

## Test

Always run tests offscreen and never against the real user store:

```bat
set QT_QPA_PLATFORM=offscreen
set FOLIO_SCREENSHOT=1
.venv\Scripts\python.exe -m pytest -q
```

- `QT_QPA_PLATFORM=offscreen` is required; do not launch an interactive GUI
  in tests.
- Tests use `tmp_path`-backed `LocalStore` instances (see `tests/conftest.py`)
  so the real user data directory is never touched.
- `FOLIO_SCREENSHOT=1` saves window captures to `artifacts/` and enables the
  font-glyph assertions in `test_screenshot`.
- On Windows the Qt `offscreen` plugin discovers **no fonts**; the test
  `qapp` fixture calls `folio.app.configure_application`, which registers
  the installed Segoe UI / Calibri faces from `%WINDIR%\Fonts` via
  `QFontDatabase.addApplicationFont`. Without that, every label renders as
  missing-glyph boxes. Native runs need no registration — the `windows`
  plugin finds all system fonts.
- A periodic native smoke (`QT_QPA_PLATFORM=windows`, temp `FOLIO_DATA_DIR`,
  `configure_application`, show → `grab()` → close) confirms real font
  rendering; it must close the window and leave no GUI running.

## Module roles

| File | Role |
| --- | --- |
| `main.py` / `run.bat` | Entry points. |
| `folio/app.py` | `configure_application(app)` (org/name, Fusion, Segoe UI 10, icon, stylesheet, offscreen font registration) + `main()`. |
| `folio/window.py` | `FolioWindow` main window: ribbon, ruler, docks, save/lifecycle, TOC, proofing orchestration. |
| `folio/editor.py` | `RichEditor` (`QTextEdit`) and `SafeDocument` (`QTextDocument` with data-URI image loading only). Styles, lists, tables, images, links, merge fields. |
| `folio/models.py` | `DocumentState`, `PageSettings`, `Comment`, `Revision`, `TocRegion` plus validation and `sanitize_ranges()`. |
| `folio/review.py` | `ReviewTracker`: tracked insert/delete revisions, rebasing, conflict detection, comments, TOC range transforms. Uses UTF-16 positions because Qt cursors count UTF-16 code units. |
| `folio/proofing.py` | Offline spelling (`pyspellchecker`) plus basic grammar/style heuristics. All issue offsets are UTF-16. |
| `folio/docx_io.py` | `write_docx`/`read_docx`: export via python-docx, import of external DOCX, atomic writes, embedded metadata. |
| `folio/publishing.py` | `Publication` layout/painting (columns, headers/footers, page ranges) and `write_pdf` via `QPdfWriter`. Geometry uses 96-DPI logical pixels derived from page millimetres. |
| `folio/mailmerge.py` | `{{Field}}` discovery, CSV (UTF-8 BOM) / XLSX sources, `merge_document`, `merge_to_directory`. |
| `folio/templates.py` | Built-in document templates. |
| `folio/dialogs.py` | All modal dialogs (page setup, find/replace, merge, history, recovery, about, etc.). |
| `folio/icons.py` | SVG-derived toolbar icons and the app mark. |
| `folio/storage.py` | `LocalStore`: SQLite recovery entries, version history, settings. |

## Embedded Word metadata

`write_docx` embeds `customXml/folio-state.json` inside the `.docx` ZIP. It
contains the `DocumentState` (comments, revisions, TOC region, page
settings, live HTML) plus a SHA-256 digest of every `word/` member. On
`read_docx`:

- If the digest matches and the schema validates, the embedded state is
  reused for round-trip fidelity. The digest is not a security signature.
- If the file was changed outside Folio (digest mismatch) or the metadata
  is missing/invalid, the visible Word body is imported instead and a
  compatibility warning is returned.
- Oversized archives or embedded metadata are rejected with a readable
  error rather than imported.

`write_docx` deep-copies the state, replaces `html` with the document's
live `toHtml()`, validates `state.page`, and refuses to write when the
encoded metadata exceeds `MAX_METADATA_BYTES` (16 MiB) — the error is
raised before the atomic replace so an existing file stays intact. The
caller's state is never mutated and `track_changes` is preserved;
`mailmerge.merge_to_directory` is the only path that forces
`track_changes=False` on its per-recipient copies.

## Local data

`LocalStore` uses `QStandardPaths.AppLocalDataLocation` for its root
(`FOLIO_DATA_DIR` env var overrides). Recovery rows are written on an
autosave debounce and only entries marked `dirty` are offered for
restoration; explicit saves and accepted discard/close paths mark them
clean.

## Supported vs. limited

Supported: styled text, lists, tables (incl. merged cells), images, links,
page setup/columns, headers/footers with page fields, TOC generation,
comments, tracked text insertions/deletions with accept/reject, mail
merge to `.docx`, PDF/print publishing, local recovery + version history.

Honest limits (also stated in About): proofing is heuristic only; tracked
changes cover text edits and are Folio metadata, not Word-native
revisions; complex external DOCX features (sections, fields, floating
graphics) import with best effort and may simplify; the editing canvas
shows a single page-width column — multi-column layout is rendered in
the publishing preview/PDF only; embedded Folio metadata is capped at
16 MiB and saves that exceed it are refused rather than truncated.

## Conventions

- No new dependencies without approval; no network features.
- Cursor/document positions are UTF-16 offsets throughout.
- Keep edits minimal and consistent with existing style; no incidental
  churn. Do not add source comments.
- Tests must not label skips as passes and must not hide unimplemented
  features behind passing tests.
