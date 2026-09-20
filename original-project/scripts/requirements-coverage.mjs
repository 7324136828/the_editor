import { readFile, mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const catalogues = [
  "requirements1.txt",
  "requirements2.txt",
  "requirement3.txt",
];
const definitions = [];
const moduleNames = new Map();

for (const filename of catalogues) {
  const lines = (
    await readFile(path.join(root, "requirements", filename), "utf8")
  )
    .replace(/\r\n?/g, "\n")
    .split("\n");
  for (let line = 0; line < lines.length; line++) {
    const value = lines[line].trim();
    if (/^[A-Z]+-[A-Z]+$/.test(value) && /^\d+$/.test(lines[line + 2]?.trim()))
      moduleNames.set(value, lines[line + 1].trim());
    let match;
    if (filename === "requirements1.txt") {
      match = value.match(/^(REQ-[A-Z0-9-]+)\s*\|\s*(.*?)\s*\[([BOLE])\]$/);
      if (match)
        definitions.push({
          id: match[1],
          title: match[2],
          profile: match[3],
          requirement: lines[line + 1]?.trim() ?? "",
          acceptance:
            lines[line + 2]?.trim().replace(/^Acceptance:\s*/, "") ?? "",
          source: `requirements/${filename}`,
          sourceLine: line + 1,
        });
    } else if (filename === "requirements2.txt") {
      match = value.match(/^(REQ-[A-Z0-9-]+)$/);
      if (match)
        definitions.push({
          id: match[1],
          title: lines[line + 1]?.trim() ?? "",
          profile: null,
          requirement: lines[line + 2]?.trim() ?? "",
          acceptance: null,
          source: `requirements/${filename}`,
          sourceLine: line + 1,
        });
    } else {
      match = value.match(/^\*\s*(REQ-[A-Z0-9-]+)\s*\(([^)]+)\):\s*(.*)$/);
      if (match)
        definitions.push({
          id: match[1],
          title: match[2],
          profile: null,
          requirement: match[3],
          acceptance: null,
          source: `requirements/${filename}`,
          sourceLine: line + 1,
        });
    }
  }
}

const partial = new Map();
function add(source, prefix, numbers, evidence, available, remaining) {
  for (const number of numbers.split(" "))
    partial.set(`${source}:REQ-${prefix}-${number}`, {
      evidence,
      available,
      remaining,
    });
}
const source = "requirements/requirements1.txt";
const workspace = [
  "src/App.tsx",
  "src/hooks/useWorkspace.ts",
  "src/services/api.ts",
  "server/index.mjs",
];
const word = [
  "src/components/editor/WordEditor.tsx",
  "src/plugins/builtins.ts",
];
const grid = [
  "src/components/editor/ExcelEditor.tsx",
  "src/services/formulas.ts",
];
const slides = [
  "src/components/editor/PptEditor.tsx",
  "src/components/editor/PresentationModal.tsx",
];
const search = [
  "src/services/search.ts",
  "src/components/bottomPanel/FindInFilesView.tsx",
  "src/App.tsx",
];
const extensions = [
  "src/plugins/host.ts",
  "src/plugins/builtins.ts",
  "src/components/plugins/ExtensionsPanel.tsx",
  "server/mcp.mjs",
];

add(
  source,
  "COM-LIFE",
  "01 02 04 05 06 07 08 10 12",
  workspace,
  "Blank documents, upload, workspace reopening, explicit/save-all and optional autosave, copied/renamed files, browser recovery, dirty-close handling, basic metadata, and saved revision history are implemented in a local workspace.",
  "No native filesystem picker handles, Office template catalogue, cloud repositories, comprehensive property schema, delta recovery journal, or cross-device recovery. These compound acceptance criteria have not been certified.",
);
add(
  source,
  "COM-IO",
  "01 02 03 07",
  [
    "src/parsers/fileLoader.ts",
    "src/parsers/docxParser.ts",
    "src/parsers/xlsxParser.ts",
    "src/parsers/pptxParser.ts",
    "src/services/exportDocument.ts",
  ],
  "Basic DOCX, XLSX, PPTX import/export and editable text/code files are available.",
  "The application uses a limited internal model and may omit unsupported parts, media, styles, masters, and complex layout. No full fidelity, complete Office compatibility, web/image export, or legacy format write guarantee.",
);
add(
  source,
  "COM-EDIT",
  "01 06",
  [
    "src/components/editor/WordEditor.tsx",
    "src/components/editor/CodeEditor.tsx",
    "src/components/editor/ExcelEditor.tsx",
  ],
  "Native text fields support selection and clipboard operations; spreadsheet paste accepts tabular text.",
  "No cross-document rich clipboard router, semantic object selection engine, or clipboard history.",
);
add(
  source,
  "COM-EDIT",
  "05",
  workspace,
  "Document history supports undo and redo.",
  "History is local and bounded; repeat-command semantics and transactional edits across documents are not implemented.",
);
add(
  source,
  "COM-EDIT",
  "07 08 09 10",
  search,
  "Find and replace with case/whole-word/regular-expression options and navigation to document locations are implemented.",
  "No full Office wildcard language, formatting-aware find, arbitrary structural navigation, or complete match-by-match selection semantics.",
);
add(
  source,
  "COM-EDIT",
  "11",
  [
    "src/components/layout/CommandPalette.tsx",
    "src/App.tsx",
    "src/plugins/host.ts",
  ],
  "The command palette searches application and enabled plugin commands.",
  "Contextual help is limited to labels, docs, and notifications.",
);
add(
  source,
  "COM-EDIT",
  "13",
  slides,
  "Slides can be duplicated and slide objects can be moved.",
  "No general cross-application object drag-and-drop or rich clipboard duplication.",
);
add(
  source,
  "COM-PREF",
  "01 02 03 05 06 07",
  [
    "src/App.tsx",
    "src/components/layout/ActivityBar.tsx",
    "src/components/layout/StatusBar.tsx",
    "README.md",
    "docs/plugins.md",
  ],
  "Theme, panel visibility, autosave, command shortcuts, dynamic document icons, save status, docs, and user-visible diagnostics are available.",
  "No exhaustive preferences, shortcut customization, training/support service, feedback submission, or formal accessibility certification.",
);
add(
  source,
  "OBJ-TYPE",
  "01 03 07",
  [...word, ...grid, ...slides],
  "Font controls, paragraph/cell bold and italic, and selected-object text color exist in the supported editors.",
  "Typography coverage varies by editor; arbitrary selected-run formatting, underline, and full mixed-run fidelity are not implemented.",
);
add(
  source,
  "OBJ-PARA",
  "01 02 11",
  word,
  "Paragraph alignment, document line spacing, and heading/body/quote/callout styles are editable.",
  "No comprehensive paragraph spacing, cascading style inheritance, list engine, or pagination control.",
);
add(
  source,
  "OBJ-SHAPE",
  "01 02 03 06 07 11 12",
  slides,
  "Presentation text, shapes and metric objects have editable content, fill, size, position, rotation, and layer values.",
  "No connectors, decorative text engine, precise selection pane parity, shape effects, flips, or aspect-ratio lock engine.",
);
add(
  source,
  "DOC-PAGE",
  "01 02",
  word,
  "Three margin presets and portrait/landscape page orientation are implemented.",
  "No arbitrary page sizes, mirror/custom margin engine, mixed-section pagination, or physical page layout conformance.",
);
add(
  source,
  "DOC-TABLE",
  "01 05 06",
  word,
  "Editable tables can be inserted, extended with rows/columns, and removed as blocks.",
  "No freehand drawing, individual cell/row/column deletion, advanced table geometry, or complete table layout engine.",
);
add(
  source,
  "DOC-REVIEW",
  "09",
  word,
  "Paragraph-context comments can be added, resolved, and reopened.",
  "No reply threads, mentions, arbitrary text-range anchors, or full comment CRUD/interchange fidelity.",
);
add(
  source,
  "DOC-VIEW",
  "01 03 04 06 09",
  [...word, "src/components/sidebar/WordOutlineView.tsx"],
  "A document page canvas, heading navigation, accordion outline, and zoom are available.",
  "No draft/web view modes, complete outline editing, true multi-page layout, or print-pagination verification.",
);
add(
  source,
  "GRID-EDIT",
  "01 02 08 10 13",
  grid,
  "Cells and the formula bar are editable; entries can be committed/cancelled; values can be cleared and tabular text pasted; formula text is searchable.",
  "No complete multiline interaction/clipboard formats, clear-format modes, annotations, or advanced fill engine.",
);
add(
  source,
  "GRID-FMT",
  "01 02 04",
  [...grid, "src/components/rightView/PropertyInspector.tsx"],
  "Basic number/currency/percent formatting and cell alignment are represented in the model and inspector.",
  "No full locale-format library, accounting/date/custom formatting, or full vertical alignment and typography parity.",
);
add(
  source,
  "GRID-SHEET",
  "01",
  grid,
  "Rows, columns, and worksheets can be appended.",
  "Insertion at arbitrary boundaries, shifting references, structured tables, and advanced sheet operations are not implemented.",
);
add(
  source,
  "GRID-FUNC",
  "03 04 09 10 16",
  [...grid, "README.md"],
  "Arithmetic, cell/range/cross-sheet references, SUM, AVERAGE, MIN, MAX, COUNT, ABS, and ROUND are implemented without evaluating JavaScript.",
  "No copy/fill reference translation, external workbook links, full math/statistics library, locale aliases, or all function-family semantics. Unsupported expressions report errors.",
);
add(
  source,
  "GRID-CALC",
  "01 02 03 04 12 13",
  grid,
  "Supported formulas recalculate with dependency resolution and cycle/error reporting.",
  "No incremental dependency graph, manual/iterative calculation modes, formula step debugger, complete audit UI, or full spreadsheet numerical conformance.",
);
add(
  source,
  "GRID-VIEW",
  "03 05 17",
  [...grid, "src/plugins/context.ts"],
  "The grid shows formula entry, row/column headings, sticky headers, and statistics/diagnostics.",
  "No user-defined freeze/split pane engine, full display-option matrix, or large-grid virtualization guarantee.",
);
add(
  source,
  "DECK-SLIDE",
  "01 04 10",
  slides,
  "Create, duplicate, delete and reorder slides; edit object text.",
  "No hierarchical list-level editor or complete slide lifecycle/layout inheritance acceptance.",
);
add(
  source,
  "DECK-DESIGN",
  "05",
  [...slides, "src/components/rightView/PropertyInspector.tsx"],
  "Slide background color is editable.",
  "No full gradient/image/texture fill engine or theme background inheritance.",
);
add(
  source,
  "DECK-SHOW",
  "01 09 12",
  slides,
  "Start a slideshow, navigate slides with keyboard/buttons, display notes, and exit.",
  "No full presenter overview, blanking/pause controls, multi-monitor presenter view, audience features, or complete delivery engine.",
);
add(
  source,
  "DECK-VIEW",
  "01 02 05 09",
  slides,
  "An editable slide canvas, filmstrip, notes pane, slide ordering, and zoom are available.",
  "No separate full-screen slide sorter, complete outline/comments view, notes-page layout, or color-display mode library.",
);
add(
  source,
  "EXT-AUTO",
  "10 11 12",
  extensions,
  "Trusted source-controlled plugins can register commands, declare document permissions, enable/disable, and inspect document context; read-only MCP serves saved documents.",
  "No remote install marketplace, arbitrary plugin execution, OS sandbox, macro engine, or debugger. JavaScript permissions are host contracts, not a security sandbox.",
);
add(
  source,
  "ENG-CORE",
  "01 02 03 04 05 06 09 10",
  [
    ...workspace,
    ...extensions,
    "tests/backend.test.mjs",
    "tests/plugins-host.test.mjs",
  ],
  "Typed models/commands, undo, atomic saved records with revision checks, local origin validation, keyboard controls, browser draft recovery, and error reporting exist.",
  "No distributed transactions, multiuser authorization, formal accessibility audit, guaranteed offline durability, or complete suite-wide acceptance verification.",
);
add(
  source,
  "ENG-TEST",
  "01 02 03",
  [
    "scripts/requirements-coverage.mjs",
    "docs/requirements-coverage.json",
    "docs/requirements-coverage.md",
    "README.md",
  ],
  "Every supplied requirement definition is indexed with an explicit status; supported scope is documented.",
  "The registry is an implementation assessment, not acceptance evidence for every clause or a release certification dossier.",
);
add(
  source,
  "ENG-TEST",
  "04 06 07 08",
  [
    "tests/backend.test.mjs",
    "tests/export.test.mjs",
    "tests/plugins-host.test.mjs",
    "tests/plugins-mcp.test.mjs",
  ],
  "Automated checks cover selected exported package structures, persistence conflicts/restart, plugin permissions, and protocol behavior.",
  "No comprehensive Office conformance corpus, exhaustive calculation oracle, full browser/assistive-technology matrix, or full requirement-by-requirement acceptance suite.",
);

const supplemental = "requirements/requirements2.txt";
add(
  supplemental,
  "DOC-FILE",
  "01 02",
  workspace,
  "Basic DOCX/text interchange, atomic workspace save, dirty state, and browser draft recovery.",
  "No legacy DOC/DOTX fidelity, filesystem locks, or timed differential recovery journal.",
);
add(
  supplemental,
  "DOC-TYPO",
  "02 03 05 06",
  [...word, ...search],
  "Browser-rendered typography, paragraph alignment/spacing, basic block styles, and text search.",
  "No custom shaping engine, OpenType controls, style inheritance tree, or full formatting-aware search semantics.",
);
add(
  supplemental,
  "DOC-REV",
  "03",
  word,
  "Comments record selected paragraph context and resolution.",
  "No anchored text-span topology, reply threads, or coauthor mentions.",
);
add(
  "requirements/requirement3.txt",
  "W-REV",
  "02",
  word,
  "Paragraph-context comments and resolution.",
  "No arbitrary run anchors, native threaded comment import/export, reply graph, or mentions.",
);
add(
  "requirements/requirement3.txt",
  "W-VIEW",
  "01 03 04",
  [...word, "src/components/sidebar/WordOutlineView.tsx"],
  "Page canvas, readable layout preset, zoom, outline projection, and heading navigation.",
  "No draft/web rendering modes, viewport virtualization, speech/syllabification engine, or drag-and-drop heading tree reorganization.",
);
add(
  "requirements/requirement3.txt",
  "X-FILE",
  "01 02",
  workspace,
  "Basic XLSX interchange, atomic saved state, browser recovery cache.",
  "No XLSM/XLSB preservation, multi-gigabyte/multithreaded serialization, memory-mapped locks, or delta calculation journals.",
);
add(
  "requirements/requirement3.txt",
  "X-VIEW",
  "01",
  grid,
  "Sticky worksheet headers support basic grid navigation.",
  "No four-quadrant freeze pane engine or complete virtualization system.",
);

const records = definitions.map((definition) => {
  const details = partial.get(`${definition.source}:${definition.id}`);
  return {
    key: `${definition.source}#${definition.id}`,
    ...definition,
    module: definition.id.replace(/^REQ-/, "").replace(/-\d+$/, ""),
    status: details ? "partial" : "not-implemented",
    acceptanceVerified: false,
    evidence: details?.evidence ?? [],
    available:
      details?.available ??
      "No delivered implementation of this requirement is claimed.",
    remaining:
      details?.remaining ??
      "Deferred: the requirement and its acceptance criteria remain unimplemented/unverified in this build.",
  };
});
const primary = records.filter((record) => record.source === source);
if (primary.length !== 753)
  throw new Error(
    `Expected 753 primary catalogue definitions; parsed ${primary.length}. Review the parser before updating coverage.`,
  );
for (const [key] of partial)
  if (!records.some((record) => `${record.source}:${record.id}` === key))
    throw new Error(
      `Coverage rule does not identify an actual definition: ${key}`,
    );
const count = (list) => ({
  definitions: list.length,
  partial: list.filter((item) => item.status === "partial").length,
  notImplemented: list.filter((item) => item.status === "not-implemented")
    .length,
  acceptanceVerified: 0,
});
const manifest = {
  schemaVersion: 1,
  assessedAt: "2026-09-19",
  interpretation:
    "Conservative source-code coverage assessment. Partial is not a passed acceptance test. All compound catalogue requirements remain unverified in full; unlisted capability portions are deferred. Definitions use source + id because identifiers can have different meanings in different catalogues.",
  summary: count(records),
  sources: catalogues.map((filename) => ({
    source: `requirements/${filename}`,
    ...count(
      records.filter((record) => record.source === `requirements/${filename}`),
    ),
  })),
  requirements: records,
};
const modules = [...new Set(primary.map((record) => record.module))].map(
  (module) => ({
    module,
    name: moduleNames.get(module) ?? module,
    ...count(primary.filter((record) => record.module === module)),
  }),
);
const overview = `# Requirements coverage

This build is a usable local document workbench with a VS Code layout. It is not a full implementation of the supplied Office-suite specifications. Those specifications include ${primary.length} primary requirements plus ${records.length - primary.length} supplemental definitions covering features such as notebooks, enterprise rights management, advanced typography, analytics, and media engines.

The [machine-readable coverage register](./requirements-coverage.json) indexes **every requirement definition in all three supplied files**, retaining its title, requirement text, source, profile, acceptance checkpoint where supplied, and a conservative implementation status. Duplicate IDs in different files remain distinct by source. References to IDs in explanatory prose are not counted as new definitions.

**Partial** means a limited related feature exists; it does not mean the full compound requirement or its acceptance test passes. **Not implemented** means deferred with no delivery claim. No catalogue acceptance criterion is marked fully verified. This deliberately distinguishes delivered code from complete Office feature parity.

## Delivered scope

| Area | Available | Main boundary |
| --- | --- | --- |
| Workbench | Dynamic file-type activity icon, explorer, tabs, command palette, inspector, search, status, dark/light theme | Local single-user app; no full VS Code extension compatibility |
| Persistence | New/open/edit/save/save-all/copy/rename, optional autosave, recovery cache, undo/redo, revision conflict reporting and saved history | Backend stores internal JSON models; browser drafts depend on browser storage; no cloud coauthoring |
| Word | Paragraph text/styles, basic bold/italic/alignment, font controls, tables, comments, page layout presets | Paragraph-level formatting; no Word pagination, arbitrary run selection, tracked changes, or advanced layout fidelity |
| Spreadsheet | Cell/formula entry, tabular paste, sheet creation, row/column extension, number formats, supported formula recalculation | Limited evaluator; no full Excel function library, pivot tables, charts, query engine, or very-large-grid guarantees |
| Presentation | Slide creation/duplication/reorder/deletion, editable text/shapes, geometry, notes, slideshow | No master/layout inheritance, media playback, animation engine, or PowerPoint parity |
| Interchange | Basic DOCX/XLSX/PPTX import/export and text/code files | Supported model only; complex styles, media, unsupported parts and native layout may be omitted |
| Extensions | Typed trusted plugin host, layout plugin, current-document context/checks, JS example | No arbitrary uploaded script execution or third-party extension marketplace; permissions are not a JS sandbox |
| MCP | Read-only local JSON-RPC HTTP server with tools/resources plus Python stdio bridge | Saved documents only; no AI model, LSP client, remote MCP runner, or write tools |

Supported spreadsheet functions are **SUM, AVERAGE, MIN, MAX, COUNT, ABS, and ROUND**, plus supported arithmetic, ranges, and cell/cross-sheet references. Unsupported syntax/functions return explicit errors. This is a declared subset, not Excel formula compatibility.

## Deferred capability families

Notebook/OneNote content models; advanced Word pagination, references, mail merge, forms and revision comparison; Office-native drawing, image editing, diagrams and chart engines; full language/proofing services; pivot/query/data-model/forecasting systems; encrypted packages and enterprise rights management; multiuser synchronization; audio/video/ink; presentation masters and animations; executable macros; AI generation; educational services; full native Office round-trip fidelity; native DirectWrite/Direct3D engines; and certification against the complete requirement catalogue remain outside this build.

## Catalogue totals

| Source | Definitions | Partial | Not implemented | Fully accepted |
| --- | ---: | ---: | ---: | ---: |
${manifest.sources.map((item) => `| ${item.source} | ${item.definitions} | ${item.partial} | ${item.notImplemented} | 0 |`).join("\n")}
| **Total** | **${records.length}** | **${manifest.summary.partial}** | **${manifest.summary.notImplemented}** | **0** |

## Primary catalogue modules

| Module | Capability | Requirements | Partial | Not implemented |
| --- | --- | ---: | ---: | ---: |
${modules.map((item) => `| ${item.module} | ${item.name} | ${item.definitions} | ${item.partial} | ${item.notImplemented} |`).join("\n")}

## Verification and maintenance

The test suite exercises specific persistence, export, formula, search, plugin, and MCP behavior. Passing those tests does not satisfy the complete catalogue. Browser interaction checks verify the implemented workflows; desktop Office layout/fidelity and accessibility conformance are not certified.

Regenerate both coverage files with <code>node scripts/requirements-coverage.mjs</code>. The script parses the supplied definitions and applies explicit reviewed partial-coverage rules; all other definitions default to not implemented. It fails if the primary 753-definition catalogue changes or a coverage rule refers to a nonexistent ID. Update the rules after inspecting new implementation and its verification evidence.
`;
await mkdir(path.join(root, "docs"), { recursive: true });
await writeFile(
  path.join(root, "docs/requirements-coverage.json"),
  `${JSON.stringify(manifest, null, 2)}\n`,
);
await writeFile(path.join(root, "docs/requirements-coverage.md"), overview);
console.log(JSON.stringify(manifest.sources, null, 2));
