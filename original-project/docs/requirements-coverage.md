# Requirements coverage

This build is a usable local document workbench with a VS Code layout. It is not a full implementation of the supplied Office-suite specifications. Those specifications include 753 primary requirements plus 162 supplemental definitions covering features such as notebooks, enterprise rights management, advanced typography, analytics, and media engines.

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
| requirements/requirements1.txt | 753 | 104 | 649 | 0 |
| requirements/requirements2.txt | 126 | 7 | 119 | 0 |
| requirements/requirement3.txt | 36 | 7 | 29 | 0 |
| **Total** | **915** | **118** | **797** | **0** |

## Primary catalogue modules

| Module | Capability | Requirements | Partial | Not implemented |
| --- | --- | ---: | ---: | ---: |
| COM-LIFE | Content lifecycle and persistence | 15 | 9 | 6 |
| COM-IO | Interchange, export, and printing | 14 | 4 | 10 |
| COM-EDIT | Clipboard, editing, and command discovery | 13 | 9 | 4 |
| COM-SHARE | Sharing, collaboration, and governance | 12 | 0 | 12 |
| COM-LANG | Proofing, language, and accessible reading | 10 | 0 | 10 |
| COM-PREF | Preferences, assistance, and workspace controls | 8 | 6 | 2 |
| OBJ-TYPE | Text runs and typography | 14 | 3 | 11 |
| OBJ-PARA | Paragraphs, lists, and semantic styles | 13 | 3 | 10 |
| OBJ-IMG | Pictures, illustrations, and image editing | 14 | 0 | 14 |
| OBJ-SHAPE | Shapes, text containers, and arrangement | 14 | 7 | 7 |
| OBJ-DIAG | Structured diagrams and mathematical notation | 10 | 0 | 10 |
| OBJ-CHART | Charts and data-driven visual objects | 14 | 0 | 14 |
| OBJ-INK | Digital ink and pen interaction | 10 | 0 | 10 |
| DOC-PAGE | Page geometry, sections, and text flow | 12 | 2 | 10 |
| DOC-THEME | Document appearance and page decoration | 10 | 0 | 10 |
| DOC-INSERT | Document structure and reusable content | 13 | 0 | 13 |
| DOC-HF | Headers, footers, and page numbering | 10 | 0 | 10 |
| DOC-TABLE | Document tables and cell layout | 16 | 3 | 13 |
| DOC-REF | References, notes, and generated lists | 15 | 0 | 15 |
| DOC-MERGE | Personalized documents, labels, and envelopes | 13 | 0 | 13 |
| DOC-REVIEW | Revisions, comparison, and editorial review | 13 | 1 | 12 |
| DOC-VIEW | Document views, outline, and windows | 12 | 5 | 7 |
| DOC-FORM | Structured forms and document controls | 10 | 0 | 10 |
| GRID-EDIT | Cell entry, filling, and selection | 16 | 5 | 11 |
| GRID-FMT | Cell formatting and conditional rules | 16 | 3 | 13 |
| GRID-SHEET | Worksheet structure and structured tables | 15 | 1 | 14 |
| GRID-FUNC | Formula authoring and function library | 16 | 5 | 11 |
| GRID-CALC | Calculation, names, and formula auditing | 14 | 6 | 8 |
| GRID-SORT | Sorting, filtering, and record selection | 8 | 0 | 8 |
| GRID-CLEAN | Data preparation, validation, and outlining | 12 | 0 | 12 |
| GRID-QUERY | External data and transformation workflows | 18 | 0 | 18 |
| GRID-MODEL | Relational modeling and advanced analytics | 10 | 0 | 10 |
| GRID-PIVOT | Pivot summaries and multidimensional reports | 18 | 0 | 18 |
| GRID-FILTERUI | Visual filter controls and micro-charts | 12 | 0 | 12 |
| GRID-WHATIF | Forecasting, scenarios, and optimization | 8 | 0 | 8 |
| GRID-PRINT | Workbook themes and print layout | 11 | 0 | 11 |
| GRID-VIEW | Workbook views, review, and protection | 17 | 3 | 14 |
| GRID-AUDIT | Workbook comparison and integrity tools | 4 | 0 | 4 |
| DECK-SLIDE | Slide lifecycle, layouts, and sections | 14 | 3 | 11 |
| DECK-DESIGN | Slide design, dimensions, and backgrounds | 9 | 1 | 8 |
| DECK-MASTER | Slide, handout, and notes templates | 10 | 0 | 10 |
| DECK-INSERT | Presentation objects and interactive navigation | 12 | 0 | 12 |
| DECK-MEDIA | Audio, video, and camera objects | 14 | 0 | 14 |
| DECK-TRANS | Scene transitions | 8 | 0 | 8 |
| DECK-ANIM | Object animation and sequencing | 15 | 0 | 15 |
| DECK-SHOW | Presentation execution and presenter controls | 15 | 3 | 12 |
| DECK-RECORD | Recording, editing, and media publication | 11 | 0 | 11 |
| DECK-VIEW | Presentation review, navigation views, and output | 13 | 4 | 9 |
| NOTE-STRUCT | Notebooks, section groups, and storage | 12 | 0 | 12 |
| NOTE-PAGE | Pages, freeform content, and navigation | 14 | 0 | 14 |
| NOTE-TAG | Tags, tasks, and integrated note actions | 11 | 0 | 11 |
| NOTE-INSERT | Attachments, printouts, tables, and captured material | 16 | 0 | 16 |
| NOTE-AUDIO | Recording, playback, and searchable media | 9 | 0 | 9 |
| NOTE-MATH | Handwriting and mathematical note tools | 8 | 0 | 8 |
| NOTE-HISTORY | Synchronization, versions, and author history | 13 | 0 | 13 |
| NOTE-VIEW | Page display, search, and note protection | 13 | 0 | 13 |
| EXT-AUTO | Macros, scripts, and extensibility | 12 | 3 | 9 |
| EXT-CONTROL | Controls, structured mappings, and compatibility adapters | 8 | 0 | 8 |
| EXT-ASSIST | Assisted authoring and connected computation | 10 | 0 | 10 |
| EXT-EDU | Educational and specialized workspace extensions | 6 | 0 | 6 |
| ENG-CORE | Cross-cutting engineering requirements | 10 | 8 | 2 |
| ENG-TEST | Coverage, acceptance, and release management | 10 | 7 | 3 |

## Verification and maintenance

The test suite exercises specific persistence, export, formula, search, plugin, and MCP behavior. Passing those tests does not satisfy the complete catalogue. Browser interaction checks verify the implemented workflows; desktop Office layout/fidelity and accessibility conformance are not certified.

Regenerate both coverage files with <code>node scripts/requirements-coverage.mjs</code>. The script parses the supplied definitions and applies explicit reviewed partial-coverage rules; all other definitions default to not implemented. It fails if the primary 753-definition catalogue changes or a coverage rule refers to a nonexistent ID. Update the rules after inspecting new implementation and its verification evidence.
