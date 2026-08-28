# PDF OCR, Comments, and Conversions Design

**Status:** split decision — comments GO; OCR conditional GO; conversions NO-GO for general release

**Scope:** WisWork PC PDF surface and local export only; no Taskpane or Relay changes

**Reference:** `genspark/main` PDF annotation save path, platform OCR helpers, PDFium extraction, and PDF-to-DOCX/PPTX/XLSX rebuild pipeline. This document does not authorize copying that runtime.

## Decision

Treat these as three products, not one merge:

| Capability           | Decision                      | Initial boundary                                                                                                    |
| -------------------- | ----------------------------- | ------------------------------------------------------------------------------------------------------------------- |
| PDF comments/markup  | **GO**                        | Notes, replies, highlight, underline, strikeout; atomic save with guarded identity                                  |
| OCR                  | **Conditional GO**            | Local, page-scoped assistive text layer on macOS/Windows; no automatic document rewrite                             |
| PDF → DOCX/PPTX/XLSX | **NO-GO** for general release | Research preview only after corpus and size gates; image-preserving fallback is not marketed as editable conversion |

Comments add bounded, reversible document semantics using existing PDF primitives. OCR can improve search and accessibility without mutating the source, but platform variance requires an explicit capability contract. General editable conversion is a much larger layout-reconstruction system—upstream adds roughly 49,000 changed lines across the PDF app and converter—and cannot be safely treated as a feature transplant.

## Authority and transaction model

- The PDF bytes opened from disk, their file identity, and base digest are authoritative committed state.
- Renderer annotations, OCR results, and edits are drafts only. They never become authoritative because they appear on canvas.
- Main process is the sole write authority. It validates a request against the base digest, applies operations in memory, writes a sibling temporary file, reopens and semantically verifies it, fsyncs where supported, and atomically replaces the target.
- A successful comments save returns a new digest and stable semantic inventory. Failure retains the original file and the draft, with a safe retry/export-copy path.
- Annotation identity is a tuple: page identity/index at base revision, subtype, normalized rectangle/quad points, contents/reply relation, and optional object-number hint. Object numbers alone are never authoritative after rewrite.
- OCR text is derived data with engine/version/language/page-image digest and coordinates. It is never allowed to overwrite original text silently.
- Conversion output is a new document with its own authority. It never replaces the PDF and must carry a conversion report listing rasterized, omitted, and low-confidence regions.

## Comments and markup — GO design

Support standard PDF `Text`, `Highlight`, `Underline`, and `StrikeOut` annotations with explicit appearance streams for consistent rendering. Preserve third-party annotations byte/semantically where untouched. Replies use `/IRT` relationships; deletion first attempts the object-number hint only if all stable identity fields match, then performs a bounded page-local semantic match. Ambiguous matches fail closed.

Save stages are ordered and verified: guarded deletions, additions/updates, page operations, metadata/form operations, reopen, annotation inventory comparison, and atomic replace. Unsupported or encrypted/signature-protected files remain read-only unless save-as-copy is explicitly safe. Existing digital signatures are treated as invalidated by any write and require a blocking warning.

Acceptance requires round trips in Acrobat, Preview, Edge/Chrome PDFium, and PDF.js; correct rotated-page quad geometry; note/reply preservation; no deletion of a neighboring identical-looking thread; recovery after injected failure at every stage; and unchanged bytes or semantic inventory for unrelated annotations.

## OCR — conditional GO design

Initial OCR is a local assistive layer for scanned pages: render one bounded page image, invoke a platform engine, map line/character boxes into display coordinates, expose search/copy/accessibility with an “OCR-derived” label, and cache by page-image digest. It does not write an invisible OCR layer into the PDF in the first release.

The useful upstream pattern is a small helper protocol backed by macOS Vision or Windows.Media.Ocr, with Linux and unavailable-language states reported as unsupported rather than failure. WisWork must additionally enforce:

- No network OCR, telemetry text, or model context without a separate explicit user action.
- One helper child per job, fixed executable path under application resources, no shell invocation, bounded input/output, deadline, cancellation, and process termination.
- Maximum rendered dimension, pixels, page count, concurrent jobs, output lines/chars, and cache bytes. Suggested starting limits: 4096 px long edge, 20 megapixels, 2 concurrent pages, 30 seconds/page, 100,000 characters/page, and 250 MiB cache.
- Engine/language availability shown before the job. Windows results without confidence must not fabricate confidence; macOS confidence is a hint, not an accuracy guarantee.
- OCR boxes are untrusted and clipped to page bounds. Unicode is normalized only for indexing; copied text preserves engine output.

Writing a searchable OCR layer is a separate **NO-GO** until PDF/A, tagged-PDF/accessibility, font embedding, text positioning, signatures, redaction, and third-party compatibility have dedicated gates.

## Conversions — NO-GO design

Upstream's local pipeline is instructive: PDFium character/object extraction feeds a pure geometry IR, specialized layout analysis, then distinct DOCX, PPTX, and XLSX rebuilders; scanned pages optionally use platform OCR; uncertain/vector regions may be rasterized. That architecture is preferable to renderer scraping and should inform future research.

It is not ready to merge as a general feature because PDF has final-positioned graphics rather than editable source structure. Reading order, tables, lists, columns, formulas, clipping, transparency, embedded/subset fonts, bidi, rotations, forms, and scanned content are inferred. DOCX, PPTX, and XLSX need different semantic reconstruction and acceptance criteria. A page-image fallback preserves appearance but does not satisfy “editable conversion.”

A future preview may be enabled only as “experimental export” on a copy, with per-page quality reporting and raster fallback. Each target must have its own release gate; passing DOCX cannot authorize PPTX or XLSX.

## Format risks

- Linearized, incremental-update, object-stream, malformed, encrypted, signed, PDF/A, PDF/UA, XFA, AcroForm, portfolios, attachments, layers, and very large PDFs require distinct behavior.
- Annotation appearances, crop/media boxes, page rotation, nonzero origins, and viewer-generated default note icons can change identity and display.
- Redacted content must never reappear through OCR, comments, extraction, conversion, logs, thumbnails, or caches. True redaction detection and policy is a hard gate.
- Hidden/invisible OCR layers may be stale or malicious. Extracted text is never assumed to match visible ink without geometric validation.
- Conversion can lose reading order, accessibility tags, formulas, charts, vector art, fonts, links, comments, footnotes, headers/footers, and pagination.
- Any operation on a signed PDF invalidates signatures; UI must distinguish cryptographic signatures from visual stamps.

## Dependencies and licenses

The upstream reference uses `pdfjs-dist` for viewing, `pdf-lib` for object creation/save, `@embedpdf/pdfium` WASM for extraction/editing, HarfBuzz/font tooling, platform OCR frameworks, and Swift/C# helper binaries. The converter also depends on WisWork's document/presentation engines, `bidi-js`, and `jszip`.

Before implementation, legal and security review must record exact versions, source/build provenance, licenses and notices, transitive native/WASM components, Chromium/PDFium patent/security update policy, and reproducible helper builds. Platform Vision and Windows.Media.Ocr outputs are not redistributed models, but OS version/language-pack requirements must be documented. Unicode-derived tables require the applicable Unicode license notice. Test PDFs and fonts need redistribution clearance; proprietary customer documents never enter the corpus.

Current checkout measurements are planning signals, not packaged deltas: installed `@embedpdf/pdfium` is about 7.3 MiB, `pdfjs-dist` about 36 MiB, and `pdf-lib` about 24 MiB uncompressed on disk. CI must measure compressed installer and unpacked resource deltas per platform because tree size is not release size.

## Fixture corpus

Maintain a manifest with source/license, expected features, sensitivity class, and golden semantic/visual outputs. Include:

- Born-digital, scanned, and hybrid/searchable-scan PDFs in Latin, CJK, Arabic/RTL, Indic, mixed scripts, vertical text, and handwriting where the OS engine claims support.
- Rotations, crop boxes/nonzero origins, multi-column layouts, tables with spans, lists, forms, vector diagrams, transparency, clipping, images, links, bookmarks, tagged PDFs, and large pages.
- Notes/replies from Acrobat, Preview, PDFium, and PDF.js; identical nearby notes; missing appearances; incremental saves; deleted/reordered pages.
- Password-owner/user encryption, wrong-password retry, certified and signed PDFs, PDF/A and PDF/UA, corrupt/xref-repaired files, decompression bombs, huge page counts, and malicious nesting.
- Conversion goldens opened by current Word, PowerPoint, Excel, LibreOffice, and WisWork. Compare both format validity and target-specific semantic editability.
- Redaction fixtures proving removed text cannot be OCR'd or extracted, including raster redactions and hidden layers.

Public CI uses generated or permissively licensed fixtures. Private fidelity lanes may use licensed documents but retain no source or OCR text in logs/artifacts.

## Release size and runtime budgets

- Comments: less than 1 MiB compressed installer delta beyond already shipped PDF dependencies; save overhead below 2 seconds p95 for 100-page/50 MiB corpus files and bounded to 2× input size plus 64 MiB working memory where practical.
- OCR helper: less than 5 MiB compressed per platform, excluding OS frameworks; lazy execution; limits above; zero Linux payload when unsupported.
- Conversion preview: no release approval until compressed installer delta is measured and remains below 15 MiB incremental over the existing PDF stack, cold start regresses less than 10%, and a 500-page adversarial job stays within explicit memory/time ceilings. If PDFium is already shipped, packaging must deduplicate it rather than bundle another copy.
- Generated output, OCR cache, and temporary files use quotas and are cleaned after crash recovery without touching originals.

## Rollback and observability

Use independent flags: `pdf.comments_v2`, `pdf.local_ocr`, `pdf.convert_docx_preview`, `pdf.convert_pptx_preview`, and `pdf.convert_xlsx_preview`. Disabling a flag requires no file migration. Comment drafts remain exportable; OCR caches are disposable and versioned by engine; conversions create new files only.

Safe diagnostics include feature/stage, file-size/page-count buckets, PDF capability flags, engine/platform class, elapsed/limit reason, operation counts, verification result, and conversion quality buckets. Exclude paths, filenames, PDF bytes, comment/OCR text, passwords, annotation authors, document metadata, page images, and extracted content.

## Phases and acceptance gates

### Phase C1 — comments GO gate

- Typed IPC validation, base-digest stale detection, bounded identity matching, semantic reopen verification, atomic write, and fault-injection rollback all pass.
- Cross-viewer corpus has no critical display/relationship loss; unrelated annotations and page content remain unchanged.
- Signed/encrypted/redacted cases fail safely with correct user messaging.

### Phase O1 — OCR conditional GO gate

- macOS and Windows helpers are reproducibly built, signed, packaged, smoke-tested, cancellable, sandbox-compatible, and unavailable-platform behavior is explicit.
- Character accuracy is at least 98% for clean supported-language scans and measured—not promised—for degraded scans; box overlap/line-order thresholds are defined per script.
- Redaction leakage is zero in the security corpus; no content leaves the machine; all resource limits and crash recovery tests pass.
- Search/copy/accessibility distinguish OCR-derived text, and source bytes remain unchanged.

### Phase X0 — conversion research gate

- Freeze an IR and target-specific fidelity metrics before porting UI.
- Run at least 500 legally usable PDFs across all format-risk classes.
- Require valid outputs for 100%, no silent empty pages/sheets/slides, no redaction leakage, and an explicit report for every rasterized/omitted/low-confidence region.
- Separately require DOCX reading-order/table/style scores, PPTX editable-element/geometry scores, and XLSX cell/table/type/formula scores. Product review must validate that quality labels match user expectations.

Until all X0 criteria and packaged-size/security reviews pass for one target, that target remains **NO-GO**. The decisions in this document authorize comments implementation and an OCR prototype after its prerequisites; they do not authorize general PDF conversion runtime.
