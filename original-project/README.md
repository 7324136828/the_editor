# Code Office Studio

An editable document workbench with a VS Code layout: activity bar, adaptive explorer, file tabs, editor, inspector, bottom panels and command palette. Word documents, spreadsheets, presentations and source files share one local workspace.

## Run

Requires **Node.js 24+** and npm.

```sh
npm install
npm run dev
```

Open **http://localhost:5174**. This command starts both the Vite frontend and the local storage/MCP server on **127.0.0.1:3001**. On Windows PowerShell with script execution restrictions, use `npm.cmd` instead of `npm`.

```sh
npm run build       # Production frontend
npm start           # Serve production frontend + API at http://127.0.0.1:3001
npm run typecheck
npm test            # Persistence, formulas, native formats, plugins, MCP and state regressions
npm run test:e2e     # Isolated browser workflow tests; Microsoft Edge must be installed
npm run build:lib   # Embeddable React library + type declarations + dist/style.css
```

Browser tests use separate ports 5185/3102 and a temporary data directory, so they do not modify your normal saved workspace. Screenshots and the HTML report are written under `.verification/`.

## Using the workbench

- **New** creates a Word document, workbook, presentation, or text/code file. **Open**, the explorer's **Load File**, and drag-and-drop import files. Multiple files of the same type keep separate identities.
- **Save / Ctrl+S** persists the current model to the backend. **Ctrl+Shift+S** saves all dirty workspace files. The indicator distinguishes unsaved, saving, saved and failed states. Changes made during a save remain dirty until persisted.
- **Download** generates an editable DOCX, XLSX or PPTX, or a text/code file. Downloads are separate from workspace saves. Imports and exports support the model described below, rather than preserving every original Office part.
- **File > Save a copy** creates a separate identity. **Rename** changes the workspace filename. Closing a dirty tab offers Save, Discard or Cancel. Closed saved files remain in the explorer.
- **History** in the activity bar shows the last 50 successful saved revisions. Restoring loads a revision into the editor; saving it appends a new revision.
- **Settings** enables optional autosave after 1.5 seconds of inactivity and changes the color theme. Browser recovery drafts survive reload independently of server saves. A full or unavailable browser cache is reported.
- **Extensions** enables/disables plugins, applies real layout presets, inserts editable SmartArt in Word documents, exports document context and tests the local MCP server. Compatible plugin commands also appear in the command palette.
- **Resize panels** by dragging the divider beside either sidebar or above the bottom panel. Focus a divider and use arrow keys for keyboard resizing; double-click it to restore its default size. The workbench remembers panel sizes in this browser and fits them to the available window space.
- **View > Customize views** shows, hides and reorders the sidebar and inspector views separately for each application. Add a named view for personal notes or live document context. Hiding a view keeps its notes; deleting a custom view asks before removing its notes. View choices and notes are stored locally in this browser, separately from document saves. **Restore default views** resets the view arrangement while retaining hidden custom notes. **View > Reset panel sizes** restores default widths and height.

### Editors

| Editor | Implemented behavior |
| --- | --- |
| Word | Edit headings, paragraphs and table cells; paragraph styles, bold/italic and alignment; fonts, colors, spacing, margins and orientation; add/remove blocks, table rows/columns and comments; live outline/statistics |
| Spreadsheet | Edit cells/formulas immediately, keyboard navigation, tabular paste, formats, rows/columns/sheets, value-only sorting with a retained header; dependency recalculation and error/cycle detection |
| Presentation | Add, duplicate, delete and reorder slides; edit text, shapes and metrics; drag objects, edit geometry/colors/layer order, reflow layouts, edit notes and present slides |
| Text/code | Edit and save source, Tab indentation, line numbers, derived symbols and JSON diagnostics; source is never executed by opening it |

Spreadsheet formulas support **SUM, AVERAGE, MIN, MAX, COUNT, ABS, ROUND**, arithmetic, parentheses, percentages, cell/range references and cross-sheet references. Unsupported functions/syntax return an error instead of a fabricated result. Sort is explicitly unavailable on sheets containing formulas until reference rebasing is implemented.

**Edit > Find / Replace in files** or **Ctrl+Shift+F** opens search in the bottom panel. Search covers all workspace files and navigates to paragraphs, tables, sheets/cells, slides and code lines. Case and whole-word options are shared by search and replacement. Replace All previews the operation in a confirmation dialog, uses literal replacement text, preserves unrelated Word run styles and spreadsheet string types, and enters each document's undo history. Regex search uses a restricted pattern subset: no groups, alternatives, backreferences or repetition braces, and at most one quantifier. Patterns with * or + must begin with ^, to bound backtracking.

The **Context** inspector reports actual content, structure and diagnostics. It is deterministic local analysis; no AI model or cloud service is connected.

### Keyboard shortcuts

| Shortcut | Action |
| --- | --- |
| Ctrl+N / Ctrl+O | New / open |
| Ctrl+S / Ctrl+Shift+S | Save / save all |
| Ctrl+W | Close active editor with dirty-state handling |
| Ctrl+Z / Ctrl+Y | Undo / redo document edits |
| Ctrl+Shift+P / F1 | Command palette |
| Ctrl+Shift+F | Find and replace |
| Ctrl+Shift+X | Extensions |
| Ctrl+B | Toggle sidebar; bold when editing Word text |
| Ctrl+J / Ctrl+Alt+B | Toggle bottom panel / inspector |
| Ctrl+P | Browser print / PDF dialog |
| F5 / Escape | Present / leave slideshow |

## Storage and API

The server stores JSON document models and their revision history under **`.office-data/`**, ignored by Git. Back up this directory to retain the workspace. Imported source files on your computer are never overwritten. Server writes use temporary files plus atomic rename; a failed write leaves the prior saved version intact. A revision mismatch returns HTTP 409 and preserves the browser edits. Save a copy to retain a conflicting branch, or use Reload saved after explicitly discarding the local branch.

| Request | Behavior |
| --- | --- |
| GET `/api/health` | Storage availability |
| GET `/api/documents` | Saved document records |
| PUT `/api/documents/:id` | Save `{name, document, revision}`; use `null` for a new file |
| GET `/api/documents/:id/history` | Saved revision metadata |
| GET `/api/documents/:id/history/:revision` | Read a saved version |
| POST `/mcp` | Read-only MCP discovery and document/context resources |

`OFFICE_PORT`, `OFFICE_DATA_DIR` and `OFFICE_ALLOWED_ORIGINS` configure the server. The default development proxy targets port 3001; change `vite.config.ts` if changing that port. Use `npm run dev:client` only when the server is started separately with `npm run server`.

The backend binds to loopback, validates paths, request size, models, Host and Origin, and never accepts a filesystem path from a document ID. It is intended for a single-user local workspace. It does not provide network authentication or multiuser authorization.

## Plugins and MCP

See **[plugin development and MCP usage](docs/plugins.md)** for the typed host, working JavaScript layout plugin, client configuration and Python stdio bridge:

- [`examples/plugins/editorial-layout.js`](examples/plugins/editorial-layout.js): register a reviewed layout module with the real host.
- [`examples/plugins/mcp_stdio.py`](examples/plugins/mcp_stdio.py): bridge an MCP stdio client to the local HTTP endpoint (Python 3 required).
- The built-in MCP tools list saved files, read a saved model and return saved text/outline/statistics. Unsaved browser content is accessible only to the local editor context plugin.

The host validates returned document identity and declared read/write access and isolates mutations using cloned models. Trusted application JavaScript is not sandboxed. There is no arbitrary script upload or remote extension installation path.

## Embedding

```tsx
import { VSCodeOfficeStudio, defaultPluginHost } from 'vscode-office-studio';
import 'vscode-office-studio/style.css';

export function Workspace() {
  return <VSCodeOfficeStudio initialTheme="dark" defaultActiveFile="word" />;
}
```

Build the library before installing this folder as a package. The host application must route `/api` and `/mcp` to the backend for persistent saves and MCP; without it, editing, recovery and downloads remain available. Plugin APIs, parsers, serializers and individual editors are exported.

## Format and scope boundaries

This is a usable core workbench, not full Office parity. The [coverage register](docs/requirements-coverage.md) accounts for all **915 definitions** in the three supplied requirement files, with explicit partial/deferred status. Regenerate it with `npm run coverage:requirements`.

Supported imports include DOCX, XLSX, XLS, CSV, PPTX and declared text/code extensions. DOC/PPT legacy binary files and unsupported formats fail with an actionable error. Limits include 20 MB input files and bounded Office packages; workbooks are limited to 50 sheets, 1,000 rows, 100 columns and 10,000 visible cells per sheet. These limits keep the current nonvirtualized editor responsive.

DOCX supports basic runs, paragraphs/headings, tables, comments and global layout. Word formatting controls act on a paragraph; editing its text can flatten mixed run formatting. Page counts are estimates and browser print is not Word pagination. PPTX supports editable text/shapes, geometry, fills and notes; gradients may reduce to a solid color and advanced media/masters/animations are omitted. XLSX exports values, formulas, dimensions and number formats; rich cell fonts/fills are retained in workspace saves but are not exported by the current writer. CSV/XLS imports download as XLSX. Browser printing prints the current editor view, not a full Office print engine.

Notebooks/OneNote, cloud collaboration, tracked changes, advanced page layout, charts/pivots, full formula compatibility, media/ink, macros, enterprise protection, generative AI and LSP integration are documented as deferred. The native export packages are checked by automated round-trip tests and have also been opened successfully by LibreOffice.

Dependencies use the maintained SheetJS distribution documented in the [official installation guide](https://docs.sheetjs.com/docs/getting-started/installation/nodejs/) and the updated Vite toolchain. The Python MCP bridge is provided as an example; its runtime was not verified on this machine because the installed Python interpreter was inaccessible.
