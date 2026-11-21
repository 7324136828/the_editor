# Office Editor (Chrome / Edge extension)

View and edit Word, Excel, PowerPoint and LaTeX documents in Chrome and Edge
without Office. Manifest V3. Most formats work entirely inside the browser. An
optional local Rust service (`office-backend`) adds `.doc`, `.rtf` and `.odt`
support, gives the most faithful `.docx`/`.xlsx`/`.pptx` output and, when a
TeX engine is installed, compiles LaTeX to PDF.

## Features

- **Word editor**: `.docx`, `.md` and `.txt` open and save without the backend.
  `.doc`, `.rtf` and `.odt` need the backend. You can edit headings,
  bold/italic/underline/strike, super/subscript, colours, highlight, fonts,
  sizes, links, alignment, nested lists, tables, images and page breaks.
  Save as `.docx`, `.md`, `.html`, `.tex` or `.txt`, plus `.rtf` with the backend.
  Tools: find & replace (Ctrl+F / Ctrl+H), navigation pane, word count, zoom,
  symbols, date & time, change case, clear formatting, format painter,
  templates (letter, résumé, report, meeting notes, memo), page setup (paper
  size, orientation, margins), headers and footers with page numbers, and
  document properties. All of these are saved in `.docx` by both writers.
- **Excel editor**: `.xlsx`, `.xlsm`, `.xls`, `.ods`, `.csv` and `.tsv`. A
  multi-sheet grid with a formula bar and name box. It has a built-in formula
  engine that recalculates as you type (`lib/formula.js`), plus number and
  date formats, cell styles, merged cells, undo/redo, and adding, renaming
  and deleting sheets. Save as `.xlsx`, `.xls`, `.ods`, `.csv` or `.tsv`.
  Tools: sort (quick and multi-level), AutoFilter drop-downs, freeze panes,
  find & replace, fill handle with series (numbers, dates, weekdays, months),
  formula autocomplete with argument hints, and charts (column, bar, line,
  pie, scatter, area) that are saved as real Excel charts.
- **PowerPoint editor**: `.pptx` opens and saves without the backend. Slide
  thumbnails (drag to reorder), layouts, themes, text boxes, shapes, images,
  tables, speaker notes, move/resize/rotate with snapping, arrange and align,
  undo/redo, transitions and a full-screen slide show (F5). Save as `.pptx`,
  a self-contained `.html` slide show, or a Markdown/text outline.
- **LaTeX editor**: `.tex` with syntax highlighting, line numbers,
  autocomplete, auto-closed environments, find & replace, and a live preview
  (math as MathML) next to the source. Also an outline, a problems list with
  line numbers, snippets, a symbol palette and templates (article, report,
  letter, beamer, homework, CV). Save as `.tex`, `.docx`, `.html`, `.md` or
  `.txt`; *Compile PDF* and Save As `.pdf` need the backend plus a TeX engine
  (TeX Live, MiKTeX or Tectonic). Without one, *Print* gives a PDF of the preview.
- **New documents**: the popup's *New document*, *New spreadsheet*, *New
  presentation* and *New LaTeX document* buttons, or *New* in the editor.
- **Save in place**: files opened with *Open file…* in the editor are saved
  back to the same file (File System Access API). Files opened from the web
  are saved as downloads. *Save As* lets you pick a format and location.
- **Recent files**: the popup lists the last 10 documents you opened or
  saved, one entry per file name, always the latest saved version. Each entry
  has a remove button, and *Clear all* empties the list.
- **Web integration** (optional, see below): links and downloads of Office
  documents open in the editor, the *Open in Office Editor* context-menu item
  appears on document links, and Office files dropped onto web pages can open
  in the editor (opt-in).

## Opening files

| How | Works on |
|---|---|
| Popup → **Open file…** (opens the editor tab and its file dialog) | all platforms |
| Drag files onto an editor tab | all platforms |
| Paste a file into an editor tab | all platforms |
| Popup → **Open from URL…** (http/https only) | all platforms |
| Click a link to a document on a website (after enabling web integration) | all platforms |
| Right-click a document link → **Open in Office Editor** | all platforms |
| Drag a local file into the browser window (`file://`) | needs *Allow access to file URLs* on the extension details page |
| OS "Open with" / double-click (`file_handlers` in the manifest) | **ChromeOS only**: Chrome and Edge ignore extension file handlers on Windows, macOS and Linux |

## Install

1. Open `chrome://extensions` (or `edge://extensions`) and turn on **Developer mode**.
2. Click **Load unpacked** and select this `office-chrome-extension/` folder.
3. Pin the extension and click its icon. To open document links from
   websites in the editor, click **Enable** in the popup (or use
   *Settings → Websites → Allow*).
4. Optional: start the backend (next section).

Packaging (any OS, Node 20+, no npm install needed):

```text
node tools/build.mjs          dist/extension.zip   (Chrome Web Store upload / sideload)
node tools/build.mjs --crx    + dist/extension.crx (signed CRX3 packed extension)
node tools/build.mjs --verify dist/extension.crx   (check the signature, print the extension ID)
```

`--crx` signs with `--key FILE`, the `CRX_PRIVATE_KEY` environment variable,
or `key.pem` in this folder; if none exists it creates `key.pem` (git-ignored).
Keep that key: the extension ID is derived from it. On Windows,
`.\build.ps1` (or `build.bat`) still writes `dist\extension.zip`. The
GitHub Actions workflow `.github/workflows/chrome-extension.yml` builds both
files on pushes to `main` and `chrome-extension`, on pull requests, on `v*`
tags (which also publish a GitHub release) and on manual runs.

## Backend (optional)

The extension talks to `office-backend` on `http://127.0.0.1:47831`. It only
accepts requests from the extension (other web pages get a 403), listens on
loopback only, and never stores files.

Start it from the repository's `src/` folder:

```text
Windows      double-click src\run-backend.bat
             (or: powershell -ExecutionPolicy Bypass -File src\run-backend.ps1 [-Port 47900])
macOS/Linux  ./src/run-backend.sh [--port 47900]
any system   cd src && cargo run --release -p office-backend
```

The scripts build the backend the first time (Rust from <https://rustup.rs> is
required) and run it. Keep the window open while you work. To use another
port, pass `--port N` / `-Port N` or set `OFFICE_BACKEND_ADDR=127.0.0.1:N`,
then enter the same address in **Settings → Local backend** and click
**Test connection**. The test shows the version and which formats the
backend can open and save.

## Settings

All settings save automatically and apply immediately, including in tabs
that are already open.

| Setting | Default | What it does |
|---|---|---|
| Use the local backend / address | on, `http://127.0.0.1:47831` | Shows a warning if the address is not on this computer, because documents are uploaded to it. |
| Website access | not granted | The optional `http://*/*`, `https://*/*` permission. It is needed to read the document files you click on websites. Without it the extension cannot see any web page. |
| Open Office documents from links and downloads | on (takes effect once website access is allowed) | Applies to `.docx .doc .rtf .odt .xlsx .xls .xlsm .ods .pptx`. |
| Also open text files | off | Applies to `.txt .md .csv .tsv .tex`. Off by default because links such as GitHub's `README.md` are web pages. |
| Open Office files dropped onto web pages | off | Upload areas, text boxes and editors on the page always keep their drops. |

How link and download interception behaves:

- Only a **plain left click** is intercepted. Ctrl/Cmd/Shift-click, middle
  click, Alt-click (download) and right-click → *Save link as…* keep the
  browser's normal behaviour, and their downloads are left alone.
- The page sees every click first. If a site handles a click itself (for
  example a single-page app), the extension does nothing.
- The extension fetches the file once. If the response is not really a
  document (a sign-in page, an HTML file viewer, an error), the browser
  navigates to the link as usual.
- Downloads are held before their file name is decided, so they cannot
  finish while the extension fetches the file. The hold is released, and the
  download continues normally, whenever the editor cannot take the file:
  fetch failed, not a document, larger than 100 MB, or too many at once.
  Downloads started by the extension itself (Save, Download original) are
  never intercepted.

## Format support

| Format | Without backend | With backend |
|---|---|---|
| `.docx` | open + edit + save (built-in reader/writer, `lib/local-docx.js`) | same, plus higher-fidelity parse/render |
| `.md`, `.txt` | open + edit + save | same |
| `.doc` | – (does not open; the editor says the backend is needed) | open + edit (text, headings, lists, tables; reduced fidelity, with a notice); save as `.docx`/`.rtf`/`.md`/`.html`/`.txt` |
| `.rtf`, `.odt` | – (does not open; the editor says the backend is needed) | open + edit; save as `.docx`/`.rtf`/`.md`/`.html`/`.txt` |
| `.xlsx`, `.xlsm`, `.xls`, `.ods` | open + edit + save (SheetJS; number formats, styles, merges, widths, hidden sheets kept) | same through the backend; Save As offers `.xlsx`/`.csv`/`.tsv` from the backend plus `.xls`/`.ods` from SheetJS |
| `.csv`, `.tsv` | open + edit + save (encoding and delimiter detection) | same |
| `.pptx` | open + edit + save (`lib/local-pptx.js`) | same through the backend (`pptx.rs`) |
| `.tex` | open + edit + save; convert to `.docx`/`.html`/`.md`/`.txt` | same, plus `.pdf` when a TeX engine is installed |
| Formulas | evaluated in the browser (`lib/formula.js`) | same; cached values written to `.xlsx` |

Not supported: tracked changes, comments, macros (`.xlsm` macros are not
kept), pivot tables, embedded-object editing, legacy `.ppt` files,
PowerPoint charts/SmartArt/animations (shown as placeholders), and images,
`.bib` files and TikZ in the LaTeX preview (placeholders).

The LaTeX engine is found on `PATH`, in the usual MiKTeX / TeX Live folders,
or through `OFFICE_BACKEND_LATEX=<path to engine>`. Compiling runs without
shell escape in a temporary folder with a 60-second limit.

## Troubleshooting

- **Clicking a .docx link still downloads it.** Allow website access (popup
  → *Enable*, or *Settings → Websites → Allow*) and check that *Open Office
  documents from links and downloads* is on. Modifier clicks and links that
  open in a new tab are downloaded normally by design. Use right-click →
  *Open in Office Editor* to open them anyway.
- **A link opened a web page instead of the editor.** The server returned
  HTML, usually a sign-in page. Sign in and try again, or download the file
  and open it with *Open file…*.
- **"Backend: offline" in the popup.** Start `src/run-backend.bat` (or one of
  the other scripts above) and use *Test connection* in the settings. "Port
  already in use" means the backend is already running, or another program
  uses the port; start it with `--port` and update the address.
- **`.doc` / `.rtf` / `.odt` will not open.** These formats need the backend.
- **Links on an already open page are not intercepted after an update.** The
  extension re-injects itself into open tabs. If a page still does not
  react, reload it. Until then the browser behaves normally.
- **Dragging a local file into Chrome downloads it instead of opening it.**
  Turn on *Allow access to file URLs* on the extension's details page, or
  drag the file onto an editor tab.

## Tests

```powershell
cd office-chrome-extension
npm test                  # node --test: lib, viewer, word, excel, formula, slides, latex, plumbing, crx packaging
node tests/browser/run.mjs  # Word editor DOM <-> model tests in headless Edge/Chrome
node tests/e2e/run.mjs      # the real extension UI with real mouse/keyboard input
node tests/e2e/word-features.mjs   # find/replace, templates, page setup, headers/footers...
node tests/e2e/excel-features.mjs  # sort, filter, freeze, fill handle, charts...
node tests/e2e/slides.mjs          # PowerPoint editor and slide show
node tests/e2e/latex.mjs           # LaTeX editor and preview
cd ..\src
cargo test --offline      # backend unit and round-trip tests
```

`tests/plumbing.test.js` covers the link and download decision logic, the
content script (run in a VM with a fake DOM), and the OPFS handoff/recents
stores (against an in-memory OPFS).

The two browser runners start their own headless Edge or Chrome (own
profile and debugging port; `BROWSER_PATH` picks the browser) and close only
that instance. `tests/e2e/run.mjs` loads this folder as an unpacked
extension and runs with the backend turned off; set
`E2E_BACKEND_URL=http://127.0.0.1:47831` to test against a running backend.

## Layout

```text
manifest.json          MV3 manifest (optional host access, file_handlers for ChromeOS)
background/            service worker: downloads, link opens, context menu, content-script registration
content/policy.js      pure integration rules shared by every context (also loaded by tests)
content/content.js     link-click / drop handling on web pages (registered only after opt-in)
offscreen/             offscreen document: fetches documents and writes OPFS handoffs
popup/                 toolbar popup: new/open, open from URL, recents, backend status
options/               settings page
viewer/                editor (viewer.html/js, excel.js grid, slides.js, latex.js)
lib/                   shared modules: model-word, model-excel, model-slides, formula, latex,
                       local-docx, local-xlsx, local-pptx, templates-word, formats,
                       handoff (OPFS + recents), backend client
tests/                 node:test suites
vendor/                SheetJS, docx-preview, JSZip (vendored, unmodified)
```

### Permissions

| Permission | Why |
|---|---|
| `storage` | Settings. |
| `downloads` | Opening downloaded documents in the editor; saving files. |
| `offscreen` | A hidden page that fetches documents and stages them in browser storage for the editor tab. |
| `scripting` | Registers the link/drop script only after you allow website access. |
| `contextMenus` | The *Open in Office Editor* menu item. |
| host `127.0.0.1` / `localhost` / `[::1]` | Talking to the local backend. |
| host `file:///*` | Opening local files you drag into the browser (only with *Allow access to file URLs*). |
| optional `http://*/*`, `https://*/*` | Requested only when you enable web integration. |
