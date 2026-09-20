# References, Mailings, Review, View, and Help

These ribbons run with the existing dependencies and local storage. No network services or new packages were added. The table distinguishes implemented workflows from portions of the requested specification that remain unavailable.

| Requirement | Available behavior | Limits |
| --- | --- | --- |
| REF-01 | Heading-based TOC refreshes after edits and before output, with page numbers from the publishing engine. | Refresh is debounced. Protected generated text stays unchanged. |
| REF-02 | Independently numbered footnotes/endnotes, editable notes pane, navigation in both directions, page-footer footnotes in preview/PDF, native DOCX footnotes. | Endnotes are document-terminal, not section-terminal. Footnotes reserve 25% of the body height and long notes continue onto later pages. Continuous sections around footnotes start on a new page. |
| REF-03 | Document sources, a global library in the local store, inline citations and bibliographies; APA, MLA, Chicago and IEEE choices. | Common book/article/website patterns, not a complete CSL or publisher-specific citation engine. |
| REF-04 | Figure, Table and Equation captions renumber in document order; generated lists update their text and page numbers. | Caption labels use these three built-in sequences. |
| REF-05 | Hyperlinked heading, bookmark, caption and note references display current target text or page number. | Ctrl-click follows internal links. Missing targets are shown explicitly. |
| REF-06 | Selection-based index and legal-authority tags produce alphabetized page lists and categorized authorities. | Entries are manually tagged; legal text is not classified automatically. |
| MLG-01 | #10, DL and C5 envelopes; standard label-sheet presets, replicated preview and fixed sheet output. | Labels use one label as the editing template and export as embedded sheet images. |
| MLG-02 | CSV/XLSX ingestion, stable sorting, filters and per-recipient inclusion. | No contact APIs. Recipient data is selected locally for the session; mappings persist with the document. |
| MLG-03 | Dataset fields, column mapping, Address Block, Greeting Line and Update Labels. | Blocks use the supplied local column mapping. |
| MLG-04 | Literal IF/ELSE and SKIPIF rules, record previews, individual/compound DOCX, PDF, print batches and unsent `.eml` drafts. | No email delivery. Templates containing structured feature nodes, floating objects or sections require individual DOCX output; other output routes reject them before writing to avoid losing metadata. |
| REV-01 | Existing proofing/thesaurus, live counts from publication layout, checks for heading order, image descriptions, table headers, link labels and contrast. | Heuristic assistance, not WCAG certification or accessible-PDF tagging. |
| REV-02 | Native local speech with word highlighting when the installed voice supports it; selection/document glossary substitution. | No machine-translation endpoint. Glossary substitution is explicitly labeled; audible speech quality is not covered by automated tests. |
| REV-03 | Anchored threads, local author labels, dated replies, resolve/reopen/delete, author filtering and chronological navigation. | No user accounts, presence or remote collaboration. |
| REV-04 | Text deltas, opt-in formatting deltas, individual/bulk accept/reject, All/Simple/No Markup and Original projections. | Projections open as read-only snapshots. Structural changes are not tracked, and revisions remain Folio metadata rather than Word-native revisions. Overlapping/conflicted changes are not silently reconstructed. |
| REV-05 | Three-way text comparison from a common base, conflict-marked result export, range/heading-section locks, formatting restrictions, ink drawings with retained editable strokes. | Comparison is line-based text. Ink exports as pictures while editable strokes stay in Folio metadata. Restrictions are removable local editing safeguards, not encryption or identity-based access control. |
| VIEW-01 | Paginated publication view, editable continuous Web/Draft views, horizontal Read Mode and hierarchical Outline Mode. | Print Layout view navigates back to the main editor for editing. The main editor remains a single page-width canvas. |
| VIEW-02 | Ruler visibility, grid overlay and interactive heading/page/object navigation. | The grid is a viewing aid, not a snapping system. |
| VIEW-03 | 10–500% zoom, page-width fit, single-page fit and side-by-side multipage fit. | Main canvas and secondary views have independent zoom controls. |
| VIEW-04 | Multiple views of the active document, a docked split view, shared edit/undo history and optional proportional scrolling. | Closing or replacing the active document closes its dependent views. New Document Window still creates an independent document. |
| VIEW-05 | Manually run JavaScript document macros with bounded execution, local document properties and Git HEAD display for saved files in repositories. | No VBA, operating-system automation, external repository service or automatic document macros. |
| HLP-01 | Searchable bundled help and step-by-step tutorials. | Documentation is installed locally. |
| HLP-02 | Structured bug, feature and support reports exported as JSON. | No session routing or ticket dispatch; users deliver exports through their own support channel. |
| HLP-03 | Bundled release notes and local onboarding tutorials. | No release feed, mobile client, mobile deep links or QR-code onboarding. |

## JavaScript macros

Open **View > Macros**. `folio.text` and `folio.selection` expose the starting text. The following methods queue changes that apply together after successful evaluation:

```javascript
folio.replaceAll("old phrase", "new phrase");
folio.appendText("\nPrepared locally in Folio.");
folio.insertText("Text at the current selection");
folio.applyStyle("Heading 1");
```

Replacement is literal and case-sensitive. Evaluation runs in a separate, hidden process with a time limit and no host file/network API. Evaluation errors apply no changes; successful changes share one undo step. Macros do not run on document open. Remove document editing restrictions before running a macro.

## Local translation glossary

**Review > Glossary translation** accepts UTF-8 JSON:

```json
{
  "source_language": "English",
  "target_language": "Spanish",
  "entries": {"hello": "hola", "thank you": "gracias"}
}
```

The result replaces matching supplied phrases and leaves unmatched text unchanged. This is a terminology aid, not machine translation.

## Verification

The added `test_references`, `test_advanced_mailmerge`, `test_advanced_review` and `test_view_tools` suites exercise persistence, Unicode positions, generated content, publication, recipient selection, rules, protected edits, review projections, shared view editing, macro failure/timeout behavior and local help. They run offscreen with temporary stores, alongside the existing regression suite. Native smoke scripts close their windows and use temporary data directories.
