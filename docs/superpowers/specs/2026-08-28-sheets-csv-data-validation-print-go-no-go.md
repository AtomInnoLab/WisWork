# Sheets CSV, Data Validation, and Print Settings: Go/No-Go Design

**Status:** design only; no runtime adoption is authorized by this document

**Upstream reference:** `genspark/main` at the 2026-08-26 snapshot (`0a2c25d`)

**Decision unit:** three separately releasable features, not the upstream Sheets snapshot

## Executive decision

| Capability                                  | Decision                                                                                   | Why                                                                                                                                                                                                                                                                   |
| ------------------------------------------- | ------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| CSV export                                  | **NO-GO for large/streamed workbooks; conditional GO for bounded fully loaded sheets**     | The upstream exporter reads the live display grid in 4,096-row blocks, but accumulates the entire CSV in renderer memory and sends it as one IPC string. It also requires preload completion. That is a useful small-file UX, not a large-workbook architecture.      |
| Data validation fidelity and active-cell UI | **Conditional GO**                                                                         | WisWork already has the OOXML reader, declarative writer, and Univer validation packages. The remaining UI/fidelity slice can be adopted without a new dependency, but only after fixtures prove that the complete file rule set is authoritative before any rewrite. |
| Print settings and PDF output               | **NO-GO as one feature; staged GO for settings preservation, then bounded preview/export** | OOXML print settings can be preserved and journaled safely. Chromium PDF layout is not Excel pagination, and manual breaks, print titles, fonts, printer metrics, and structural edits create correctness risks that need separate gates.                             |

Do not cherry-pick `0a2c25d` or any other snapshot wholesale. Extract reviewed slices behind independent flags and commits. The cross-highlight feature and unrelated renderer/native changes are out of scope.

## Current baseline and upstream delta

WisWork already contains:

- CSV byte decoding, delimiter sniffing, parsing, and conversion into a minimal XLSX package in `csv-import.ts`;
- Univer data-validation dependencies, native OOXML parsing, installation, and declarative `<dataValidations>` rewriting in `xlsx-dv.ts`;
- page orientation, paper size, scale/fit, margins, gridlines/headings, print area/titles, freeze state, and header/footer write support in `xlsx-page-setup.ts`;
- the streamed XLSX sidecar and structural edit journal on which safe large-workbook behavior depends.

The notable upstream additions are:

- a 157-line renderer CSV exporter plus main/preload/save-dialog wiring and loss warnings;
- active-cell data-validation dropdown and input-message chrome;
- a 294-line effective print-settings resolver, richer PDF layout, manual page-break preview/editing, row/column headings visibility, and native indexing of saved breaks;
- broad changes to shared IPC, preload, renderer, and the native sidecar mixed with unrelated snapshot work.

These additions are learning references, not an integration boundary.

## Authority model

### CSV

CSV has no workbook-preservation authority. It is a lossy interchange projection of exactly one sheet.

- **Import byte authority:** the selected file bytes. BOM is authoritative. Without a BOM, strict UTF-8 is tried first; heuristic legacy decoding is advisory and must be disclosed or user-selectable when ambiguous.
- **Import value authority:** parsed fields. Only unambiguous plain decimal tokens may become numbers; dates, formulas, booleans, locale-formatted numbers, identifiers with leading zeroes, and strings beginning with `=`, `+`, `-`, or `@` remain text unless the user explicitly selects conversion behavior.
- **Export value authority:** live calculated display values of the selected sheet, with formula view temporarily suppressed. The exporter must never silently substitute unloaded file cells with blanks or stale caches.
- **Scope authority:** the active sheet at command start, captured by stable workbook and sheet identity. A later tab switch must not redirect the export.
- **Output authority:** a newly created CSV file. CSV export must never overwrite the source XLSX through the normal save path.

For streamed workbooks, the renderer is not authoritative for the complete sheet. Large export therefore requires a sidecar streaming protocol that overlays the structural/value journal and emits bounded byte chunks directly to a temporary output file. Until that exists, streamed CSV export is disabled.

### Data validation

- **Untouched sheet authority:** original OOXML, byte-preserved.
- **Edited validation authority:** the complete live validation-rule model for that sheet, but only after native indexing has completed and all parsed rules have been installed successfully.
- **Structural authority:** the ordered structural journal maps rule ranges and same-/cross-sheet formulas. A validation snapshot taken before a row/column operation is not independently saveable.
- **Unsupported authority:** original OOXML. Any `x14:dataValidation`, unknown type/operator/error style, unresolved formula mapping, or Univer-only multi-select rule fails closed; it must not be approximated.
- **Checkbox exception:** converting a checkbox to a two-value list is a visible compatibility degradation, not preservation. It requires explicit product acceptance and a warning in the compatibility contract.

The save request carries workbook identity, sheet identity, index generation, rule-set fingerprint, and expected package hash. A mismatch produces `conflict`, never a rewrite from a stale model.

### Print settings

- **Untouched authority:** worksheet OOXML, workbook-scoped `_xlnm.Print_Area`/`_xlnm.Print_Titles`, related printer settings, and header/footer content remain byte-preserved.
- **Edited settings authority:** a page-setup journal layered over indexed file settings. `undefined` means preserve, `null` means clear, and a concrete value means replace.
- **Coordinate authority:** file-origin print areas, title rows, and manual breaks are mapped through the structural journal; session-origin values are already in screen coordinates. Ambiguous or partially deleted references fail closed for save and fall back visibly, never silently, for preview.
- **PDF authority:** Chromium output is a WisWork rendering, not an Excel-identical print artifact. Saved OOXML settings and exported PDF are separate promises.
- **Printer-dependent authority:** printer capability, font metrics, automatic page breaks, images in headers/footers, and Excel private extensions are not derivable from OOXML alone and remain unsupported.

## Format and security risks

### CSV risks

- Encoding guesses can produce plausible but wrong CJK text. Record the chosen encoding and offer retry when confidence is low.
- Delimiter sniffing can be wrong for one-column data, decimal commas, malformed quoting, or embedded newlines.
- Export loses formulas, styles, validations, comments, drawings, hidden sheets, additional sheets, and workbook metadata. A blocking loss dialog must name formula flattening and active-sheet-only scope.
- Display strings are locale-dependent and may not round-trip numerically.
- Cells beginning with spreadsheet formula markers can become CSV-injection payloads when opened elsewhere. Default export preserves visible values but warns for dangerous leading characters; any escaping policy must be explicit because prefixing changes data.
- A rectangular used range can be enormous even when sparse. `rows × columns`, estimated output bytes, deadline, and final bytes must all have independent caps.
- A single IPC string is unsuitable for large exports and can duplicate memory several times across renderer serialization, validation, and main-process write.

### Data-validation risks

- Rewriting the whole `<dataValidations>` section can drop attributes that the model did not parse.
- OOXML `showDropDown` has inverted semantics; list literals and formula references have different quoting/leading-`=` rules.
- Excel serial dates include the 1900/1904 date-system distinction and the 1900 leap-year compatibility behavior. Date conversion must use workbook date-system metadata, not a fixed epoch.
- Cross-sheet lists, named ranges, locale list separators, unions in `sqref`, prompts/errors, error styles, blank handling, and `time` rules require bijective fixtures.
- Structural edits can invalidate ranges or formulas. Partial overlap and deleted references must fail closed.
- Active-cell overlays can become stale after scroll, zoom, merge changes, workbook rebind, or disposal. Their lifecycle must be owned by the current document generation.

### Print risks

- OOXML schema order is strict. New `rowBreaks`/`colBreaks` must be inserted before the correct following elements and automatic cached breaks must not be mistaken for manual ones.
- Break limits differ by axis; off-by-one conversion between OOXML IDs and zero-based renderer coordinates is hazardous.
- Print areas may be unions, quoted names, full-row/full-column ranges, external/3-D references, or `#REF!`. Unsupported forms must be preserved for save and identified as unsupported for preview.
- Repeated columns are not covered by the upstream print-layout resolver; repeated rows are capped. This must be visible in capability reporting.
- Header/footer formatting codes, images, paths, first/even page variants, and printer relationships cannot be reduced to left/center/right plain text without loss.
- Font substitution and Chromium table layout change pagination. Tests must not claim pixel identity with Excel.
- Structural edits, hidden/filtered rows, merged cells, drawings, manual breaks, scale, fit-to-width/height, paper size, and margins interact combinatorially.

## Dependencies and licenses

No new runtime dependency is approved.

| Dependency/source                                  | Role                                                     | License/risk decision                                                                                                                                               |
| -------------------------------------------------- | -------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Upstream GenOffice code                            | Reference implementation                                 | Repository is Apache-2.0. Preserve required notices and provenance in commits; port reviewed slices rather than snapshot files.                                     |
| Univer data-validation packages already in WisWork | Validation model and UI                                  | Already shipped and pinned in the workspace. Confirm resolved package licenses in the release SBOM; do not add a second validation engine.                          |
| JSZip already in WisWork                           | Current minimal CSV-to-XLSX import                       | MIT; already shipped. It is not approved as the large CSV export path because it does not solve bounded streaming.                                                  |
| Rust XLSX sidecar and existing crates              | Indexed metadata, bounded reads, future streaming export | Already shipped. Prefer standard library streaming plus existing protocol primitives. Any new charset/CSV crate requires a separate license and binary-size review. |
| Chromium/Electron printToPDF                       | PDF rendering                                            | Already shipped. Treat its result as WisWork PDF output with documented fidelity limits.                                                                            |

Before implementation, capture `license`, integrity hash, and resolved version for every package touched by the slice. CI fails if the SBOM gains a new runtime package or license without approval.

## Fixture corpus

All fixtures are synthetic or redistributable and contain no customer data. Keep source CSV/XLSX, expected semantic manifest, and expected warnings together.

### CSV corpus

At minimum:

- UTF-8 with and without BOM; UTF-16LE/BE; GB18030, Shift_JIS, Big5, EUC-KR, and Windows-1252;
- comma, semicolon, and tab delimiters; quoted delimiter/newline/quote; CRLF/LF/CR; empty cells/rows; trailing newline; ragged rows;
- leading zeroes, very large integers, exponent notation, decimal comma, date-like strings, formula-marker strings, NUL/control characters, and malformed/unclosed quotes;
- one sheet with formulas/styles plus a multi-sheet workbook to assert every loss warning;
- dense and sparse boundaries immediately below and above cell, character, byte, and deadline caps;
- open/export/reopen in current Excel for Windows and Mac, LibreOffice Calc, and a strict CSV parser.

### Data-validation corpus

- each supported OOXML type and operator, `none/any`, allow-blank, dropdown visible/hidden, prompts, errors, and all error styles;
- inline and referenced lists, commas/quotes/non-ASCII items, cross-sheet and defined-name references;
- 1900 and 1904 date systems, date/time boundaries, custom formulas, unions in `sqref`, merged and hidden cells;
- multiple untouched rules plus one edited rule to prove the complete declarative rewrite;
- row/column insert/delete/move affecting rule ranges and formulas;
- `x14` rules, multi-select lists, invalid operators, unsupported extensions, and partial structural overlaps to prove fail-closed behavior;
- no-op package comparison, edited semantic comparison, Excel/Calc open-without-repair, and UI lifecycle tests for scroll/zoom/rebind/dispose.

### Print corpus

- Letter, A4, and one non-default paper size; portrait/landscape; scale and fit modes; three margin presets and boundary values;
- gridlines/headings, single and union print areas, repeated rows and columns, quoted sheet names, and cleared settings;
- manual row/column breaks at first/last valid positions, duplicates, automatic-break noise, and structural shifts/deletions;
- odd/even/first headers and footers with field codes, formatting, pictures, and unsupported codes;
- hidden/filtered rows, hidden columns, merges, variable dimensions, drawings/charts, RTL/CJK, and non-bundled fonts;
- no-op OOXML byte preservation, targeted-edit semantic diff, Open XML validation, Excel/Calc repair-warning checks, and PDF visual comparison on pinned Electron/OS/font images.

The visual gate uses region-based comparison and page-count assertions. It must not use a single full-page pixel threshold to imply Excel identity.

## Resource and release-size budgets

Measure signed production artifacts for macOS arm64/x64 and Windows x64 before and after each phase; record compressed artifact size, unpacked app size, renderer bundle chunks, and native sidecar size.

- No new runtime package in phases 1–3.
- Renderer compressed growth per capability: target ≤100 KiB, hard stop at 200 KiB.
- Native sidecar compressed growth per architecture: target ≤200 KiB, hard stop at 500 KiB.
- CSV export memory: ≤64 MiB incremental RSS over the source indexing baseline and no buffer larger than 8 MiB; renderer-to-main chunks ≤1 MiB.
- CSV default caps: 5 million cells, 256 MiB output, 30-second interactive deadline. Raising a cap requires benchmark evidence and cancellation tests.
- Validation installation/editing: bounded by 100,000 rules/ranges and 5 seconds; larger sheets remain view/preserve-only.
- Print preview/export: 250,000 rendered cells, 200 pages, 30 seconds, and cancellation. Larger jobs require a reduced preview or external Excel workflow.

Exceeding a hard stop is **NO-GO**, not a reason to silently truncate.

## Phased implementation and acceptance gates

### Phase 0 — evidence baseline

1. Freeze the fixture corpus and semantic manifests.
2. Record current signed artifact sizes and performance on release hardware.
3. Add capability flags and telemetry containing only safe counts, phase, duration, and error code; never cell content, paths, formulas, or validation messages.

**Gate:** all baseline fixtures reproduce; license/SBOM review passes; no runtime behavior changes.

### Phase 1 — CSV bounded export

Implement loss dialogs and active-sheet identity capture first. For fully loaded sheets, permit bounded export. In parallel, design a cancelable sidecar-to-temp-file chunk stream for streamed workbooks; do not route large content through a renderer IPC string. Commit by atomic rename only after byte count and hash validation.

**Accept when:**

- every encoding/quoting fixture passes and Excel/Calc reopen without silent column shifts;
- formula/multi-sheet/style loss warnings are blocking and tested;
- stale workbook/sheet identity, cancellation, disk-full, write failure, and app close leave the original untouched and remove the temp file;
- all applicable resource caps and signed-size budgets pass on release hardware.

Otherwise CSV large export remains **NO-GO**. The small fully-loaded exporter may ship behind its own flag.

### Phase 2 — data-validation completion

Port only the active-cell chrome and missing bijective mappings. Keep the existing fail-closed declarative writer. Add generation ownership, complete-index gating, fingerprints, and date-system-aware conversion before broadening support.

**Accept when:**

- supported fixtures pass no-op and edited semantic round trips in both date systems;
- all unknown/extended fixtures fail before package mutation;
- unrelated worksheet/package parts are byte-identical;
- Excel and Calc open without repair and enforce the expected supported rules;
- overlays never survive document rebind/dispose and remain aligned after scroll, zoom, and merges;
- performance and release-size gates pass.

Checkbox degradation is a separate product flag. Without explicit approval it remains **NO-GO** even if the rest ships.

### Phase 3 — print-setting preservation and editing

Add indexed manual breaks, headings visibility, journal mapping, and surgical OOXML save before changing PDF output. Preserve unsupported headers/footers and printer relationships unless the user edits the owning feature; then fail closed where a lossless mapping is unavailable.

**Accept when:**

- targeted edits produce schema-valid OOXML in required element order;
- automatic breaks are not serialized as manual breaks;
- structural edits correctly map or reject print areas, titles, and breaks;
- untouched unsupported print content and unrelated ZIP entries are byte-identical;
- Excel/Calc open with no repair and show the expected settings.

### Phase 4 — bounded preview and PDF export

Resolve the file settings plus session journal into an explicit effective-print model. Present unsupported constructs and pagination differences before export. Pin Electron, OS image, and bundled font set for deterministic CI comparisons.

**Accept when:**

- page count, crop area, orientation, margins, repeated-row behavior, headers/footers, and manual breaks pass the fixture matrix;
- cancellation and caps work without freezing the renderer;
- PDF output never claims Excel pixel fidelity;
- visual, performance, and signed-size gates pass on every release platform.

Repeated columns, header/footer images, first/even variants, and printer-specific pagination remain **NO-GO** until separately designed and fixture-gated.

## Rollout and rollback

Use independent flags: `sheets.csvExportV1`, `sheets.dataValidationUiV1`, `sheets.printSettingsV1`, and `sheets.printPdfV1`. Flags are read at document-open time so a session has one behavior generation.

- Roll out internal → 5% → 25% → 100%, with at least one full release day at each external stage.
- Stop on any file corruption/repair prompt, unexpected lossy rewrite, original-file overwrite, crash-rate regression, or hard-cap bypass.
- Rollback disables the affected flag and restores the previous command surface; it does not attempt to reverse already exported CSV/PDF files.
- OOXML writes use temp output, validation, and atomic replace. On failure the original package remains authoritative.
- Files saved by an enabled phase must remain readable after the flag is disabled. No proprietary persistent metadata is allowed.

## Final GO/NO-GO checklist

A capability is **GO** only if all of the following are true:

1. Its authority and stale-generation checks are implemented exactly as above.
2. The complete fixture corpus passes no-op, edited, structural, Office/Calc, and visual gates applicable to that capability.
3. Unsupported constructs are preserved or rejected before mutation; none are silently reduced.
4. Dependency license/SBOM and signed release-size budgets pass.
5. Performance, memory, byte, page/cell, deadline, and cancellation caps pass on release hardware.
6. The capability can be independently disabled without changing file readability.
7. Product documentation states every lossy or non-Excel-identical behavior.

If any item fails, that capability is **NO-GO** while the other two may proceed independently. Approval of this design authorizes fixture and prototype work only; production runtime still requires a phase-specific implementation plan and review.
