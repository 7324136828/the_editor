# Validation record

Validated on Windows with Node.js 24.18 and installed Microsoft Edge.

| Check | Result |
| --- | --- |
| TypeScript typecheck | Passed |
| `npm test` | 56 passing tests |
| Isolated Edge end-to-end suite | 10 passing workflows, no runtime page errors |
| Final spreadsheet contrast rerun | Passed |
| Production frontend build | Passed |
| React library build | Passed, including CommonJS exports, declarations and style.css |
| Dependency installation/audit | No reported vulnerabilities after toolchain/SheetJS updates |
| Independent Office compatibility | DOCX, XLSX and PPTX samples opened and converted by installed LibreOffice |

The browser suite covers actual typing and caret behavior, paragraph formatting/table/comment editing, save before blur, server-backed reopen after clearing browser recovery, separate identities for same-type documents, dirty-close cancel, formula dependency updates, editable presentation objects/notes and slideshow navigation, plugin enable/disable and geometry changes, context/MCP discovery, case/whole-word search replacement, failed saves and retained drafts, stale-revision conflict with an independent saved copy, autosave/reload, undo/redo, rename and discard.

Unit and integration tests exercise disk-save restart durability, concurrent revision protection, input/origin/path validation, bounded history, static serving, native Office package structure and round-trip imports, formula errors/cycles and cross-sheet dependencies, plugin permissions and isolation, MCP JSON-RPC errors and saved-state reads, safe regex restrictions, typed replacements and mixed text styles, validated browser recovery, and refresh/save races.

Browser fixtures use fresh temporary storage and separate ports; normal `.office-data` was not used by those tests. Generated evidence remains under `.verification/` (ignored by Git): four browser screenshots, Playwright traces/reports, and native Office/PDF samples. The last report can reflect a targeted rerun; run `npm run test:e2e` to regenerate the complete report.

This record verifies the implemented workbench subset. It does not certify all compound acceptance criteria in the [requirements register](requirements-coverage.md), native Word pagination, advanced Office feature fidelity, multiuser/cloud behavior, or the Python bridge runtime. The existing Python interpreter could not be launched on this machine; the JavaScript HTTP MCP implementation is tested.
