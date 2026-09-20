# Recording feature guide

See [References, Mailings, Review, View, and Help](FEATURE_STATUS.md) for the current ribbon additions, requirement coverage, and explicit limits.

Folio runs locally with its existing dependencies. The following commands are available in the ribbon.

| Area | Commands and behavior |
| --- | --- |
| File | **New → Blank** opens an empty document through the existing save/discard workflow. |
| Insert: media | **Online Video** accepts HTTPS YouTube, Vimeo, SlideShare, and TED addresses and inserts a hyperlink. **Stock Media** includes a small built-in Icons, Stickers, Illustrations, and Cartoon People collection. The Images tab imports an image from disk. |
| Insert: text | **WordArt** inserts editable styled text. **Symbol** provides common, Greek, and mathematical characters. **Equation** parses supported TeX or presentation MathML into a syntax tree and renders it; use **Edit / Sign / Extract** to edit the retained source. |
| Insert: pagination | **Page Numbers → Top of Page / Bottom of Page** adds running page fields while retaining existing header/footer text. **Remove Page Numbers** removes those fields. Preview and DOCX/PDF output display them. |
| Insert: signatures | **Signature Fields** creates an interactive field. **Edit / Sign / Extract** accepts a typed name with explicit consent and records the name, timestamp, and text digest locally. |
| Design | **Document Theme** binds semantic styles to theme fonts, accent palettes, and style sets. Direct formatting is preserved. Global paragraph spacing is independent of style definitions. **Page Background** configures page color, borders, and image/text watermarks in front or behind. |
| Layout | **Size** offers A4, Letter, and Legal. **Page Setup** controls margins, columns, balancing, running fields, and line numbers. **Breaks** includes page, column, and Next/Continuous/Even/Odd section breaks; setup changes apply to the cursor's section. **Selection Pane** includes a floating-object hierarchy and the existing inline-image/table list. |
| Review | **Word Count** shows body pages, words, characters with/without spaces, paragraphs, and rendered lines. **Thesaurus** looks up a small built-in English glossary and replaces the captured word. **Tracking**, **Reviewing Pane**, and **Show Markup** control tracked text and optional formatting edits and annotation visibility. Hiding markup does not accept or delete it. |
| Review: compare | **Compare** selects an original DOCX and displays its text deletions/insertions relative to the current document in a separate read-only dialog. Both documents remain unchanged. |
| View: reading | **Immersive Reader** offers an English syllable guide, local Read Aloud, line focus and reading colors. **Read Mode** displays published pages horizontally. **Web Layout** and **Draft Mode** provide continuous editing views sharing the active document and undo history. |
| View: workspace | **Zoom…** controls percentage/page-width zoom or opens one-, two-, and multi-page previews. Previews use the publishing renderer without a printer driver. **Switch Windows** lists visible Folio windows and can create a new blank window. **Gridlines** overlays a viewing grid. Existing **Focus** mode exits with Esc. |
| Help | **Search Help** finds local documentation. **Tutorials**, **Feedback**, and **Release Notes** provide guided learning, structured local report export, and bundled release information. |

## Boundaries

- Video links do not embed or play online video; no online media catalog is queried.
- Typed signatures are local consent records, not verified digital signatures. Folio does not authenticate signers, supply certificates, prevent later edits, or contact signing services.
- WordArt uses editable character formatting, not Word drawing objects. Equations export as images, not native Word OMML. The TeX/MathML presentation subsets support fractions, roots, scripts, sums, integrals, and common symbols; unsupported syntax is rejected explicitly.
- SVG is safely rasterized; external resources and active content are rejected. GLTF/GLB support static triangle meshes and solid materials; textures, skins, animation, and compressed geometry are rejected. The retained mesh renders a static poster, not an interactive 3D viewport.
- Embedded document/file objects are passive package attachments, limited to 8 MiB each. **Edit / Sign / Extract** saves the original bytes to a chosen path. There is no OLE execution or application automation.
- Floating text and watermarks export to Word as anchored pictures. Their editable source and hierarchy remain in Folio metadata. DOCX Tight wrapping uses a rectangular polygon; Folio preview/PDF follows image alpha contours.
- Transcription provides local state/events and transcript insertion, not a built-in microphone recognizer. Comment dispatch is a local signal for integration, not a network collaboration service. Viewing is a UI mode, not a security boundary or file access restriction.
- Automatic preview/PDF hyphenation uses a small bundled English dictionary; unknown words remain unchanged. DOCX also carries Word's automatic-hyphenation settings. Spelling suggestions and the reading syllable guide are separate services.
- Comparison covers text, not formatting, images, or a merged document with Word-native revision markup.
- The thesaurus is a small English glossary. Syllable guides are heuristic English reading aids and do not change the document's spelling or hyphenation.
- Read Aloud requires an installed local speech engine and voice. Its controls report unavailability; closing the reader stops speech. Automated tests cover control behavior, not audible output quality.
- Read Mode and publishing views are read-only; Web and Draft views edit the active document with shared undo. Publishing previews show authoritative section geometry, balanced columns, and floating-object text flow; the main editing canvas remains a single page-width editing surface. Very tall table rows can split in wrapped publishing output.
- Column balancing currently applies to ordinary text sections without wrapping floats or explicit page/column breaks. Explicit breaks retain their requested positions.

## Engineering requirements

| Requirements | Implementation and entry points |
| --- | --- |
| HOME-01–04 | Clipboard, **Format Painter → Sample/Apply**, typography and text effects, sub/superscript, stepped sizes, **Aa** case commands. |
| HOME-05–09 | Paragraph alignment/spacing/indents, dynamic multilevel lists, **Borders & Shading**, semantic styles/navigation, literal Find/Replace. |
| HOME-10 | Header **Editing / Reviewing / Viewing**, proofing, **Transcription**, and `window.hooks` signals (`transcriptReady`, `commentDispatched`, `modeChanged`, `dictationStateChanged`). Reviewing enables tracked text edits; Viewing disables editing commands. |
| INS-01–03 | Cover/blank pages and breaks; table insertion plus **Dimensions**; PNG/JPG/SVG, **3D Model**, retained editable SmartArt steps, charts. |
| INS-04–05 | Links, **Bookmark**, **Cross-reference** (text/page), auto-updating date/time, running `{CurrentPage}` / `{TotalPages}` tokens and legacy aliases. Fields refresh before saving and periodically. |
| INS-06–08 | Floating text/images, drop caps, WordArt, passive attachments, structured equation editor, typed signature fields. |
| DES-01–04 | Semantic theme/style bindings, shape shading tokens, independent spacing policies, page backgrounds/borders/watermarks. Native DOCX theme/style definitions are exported. |
| LAY-01–04 | Section-isolated geometry, column balancing, all four section starts, line-number gutters, dictionary hyphenation. |
| LAY-05–07 | Square/Tight/Behind/In Front flow; z-order; bounding-box editing; align left/center/distribute; grouping, child visibility, and hierarchy inspection. |

Structured sources, object hierarchy, themes, and section settings persist in the existing validated Folio metadata inside DOCX. The total 16 MiB metadata cap still applies. Text Undo does not undo document-level design/arrange settings; use saved local version history for those settings.

Regression coverage includes the `test_feature_*`, `test_document_features`, `test_design`, `test_formatting`, `test_floating`, `test_structured_media`, and `test_hyphenation` suites alongside existing editor, DOCX, publishing, review, and application tests. Tests use offscreen Qt and temporary local stores. UI captures are written to `artifacts/` when `FOLIO_SCREENSHOT=1`.
