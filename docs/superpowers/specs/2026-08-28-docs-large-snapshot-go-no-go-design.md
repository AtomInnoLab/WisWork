# Docs Large Snapshot Go/No-Go Design

## Decision

**NO-GO for merging the upstream Docs snapshot wholesale.** The `genspark/main` delta combines
OOXML parsing and rewriting, a new pagination model, header/footer editing, table behavior, image
conversion, fonts, Electron upgrades, encryption, and broad UI changes. Its useful behavior is not
an independently reversible unit, and importing it would make fidelity regressions impossible to
attribute.

**GO for a staged evaluation and selective reimplementation**, one feature family at a time, only
after the gates below pass. The first production candidate is preservation-only OOXML hardening;
editable header/footer variants, nested-table editing, print, and new image conversion remain
NO-GO until their own stage is approved.

This document evaluates upstream snapshot commits through `0a2c25d` plus focused fixes such as
`8e3ad2a`, `d2e03aa`, and `6a2d1cd`. It does not authorize copying runtime code.

## Goal and non-goals

Goal: decide which upstream ideas can improve WisWork Docs print, header/footer, table, and image
fidelity without expanding authority, corrupting unsupported OOXML, or silently increasing the
desktop release and maintenance burden.

Non-goals:

- no wholesale cherry-pick or runtime change;
- no promise of Microsoft Word pixel parity;
- no support claim based only on parse success or a browser screenshot;
- no cloud conversion, document upload, shell command, dynamic dependency install, or arbitrary
  URL fetch;
- no password/encryption work, unrelated editor UI, AI tools, or Electron upgrade in these stages.

## Architecture boundary

The DOCX package remains authoritative for persistence. The editor projection is a bounded view,
not a lossless replacement for unknown OOXML. Untouched parts must stay byte-identical; an edit may
rewrite only the explicitly owned part and relationships, with semantic reopen verification.

Pagination and print are renderer projections. They may consume the parsed document and local
fonts, but they do not gain write authority over the DOCX. Image decoders accept only bytes already
inside the opened package (or inserted through the existing confirmed file flow) and return a
display derivative; original media remains authoritative for save.

## Global constraints

- Preserve existing file-open, save confirmation, atomic-write, dirty-state, and external-change
  behavior.
- Bound ZIP entries, inflated bytes, XML depth, table depth, image bytes, decoded pixels, page
  count, conversion time, and concurrent conversions before allocation.
- Reject path traversal, relationship escape, external image relationships, malformed dimensions,
  decompression bombs, and unsupported media with stable safe errors.
- Never log body text, header/footer text, table cells, media bytes, filesystem paths, or rendered
  pages. Diagnostics contain only stable codes, types, counts, stages, and bounded dimensions.
- Preserve unknown namespaces, `mc:AlternateContent`, relationships, and parts unless the approved
  operation explicitly owns them. A failed or uncertain verification leaves the original file in
  place.
- Feature flags and package boundaries must permit each stage to be disabled independently.

## Upstream evidence and risks

### Print and pagination

Upstream renders `.pv-page` nodes, clones the same page for its print preview, parses all/current/
custom ranges, hides unselected pages only in print CSS, and invokes Electron's system print dialog.
The same-source preview is a sound idea: it avoids maintaining a separate print renderer.

Risks are material. Browser layout depends on fonts, zoom, DPI, hyphenation, image decode timing,
CSS fragmentation, and platform print engines. A stable DOM page count is not proof that fonts and
images have settled. Headers, footers, floating objects, nested tables, widow/orphan rules, section
breaks, fields, notes, and unsupported shapes can move content across pages. System print cancel and
failure also differ by Electron/platform. Therefore print is NO-GO until pagination has visual and
cross-platform gates; printing an approximation without an explicit warning is not acceptable.

### Header/footer fidelity

Upstream models default/first/even variants, `titlePg`, `evenAndOddHeaders`, multiple sections,
page-number fields, rich paragraphs, layout tables, inline and floating pictures, crop, alignment,
wrap, and `mc:AlternateContent` fallback. Its tests also preserve unedited variants byte-for-byte
and keep schema order. Those are the right invariants.

Risks include linked-to-previous semantics, per-section references, `w:type="odd"` producer
variants, fields, nested tables, text boxes, watermarks, shapes, relationship allocation, part-name
collisions, root namespaces, and section inheritance. Flattening a rich part to text would be data
loss. Editing one shared part can unintentionally change several sections. Header/footer editing is
therefore NO-GO until identity, sharing, and per-part compare-before-write are explicit.

### Table fidelity

Upstream preserves table-cell paragraph properties during rebuilds, models nested tables, and caps
deep nesting by flattening content below a modeled depth while keeping the original package intact.
The 2,000-level adversarial test is a useful boundedness pattern.

Risks include grid spans, vertical merges, omitted cells, fixed versus autofit layout, percentage
widths, borders, conditional table styles, row splitting, repeating headers, floating tables,
bi-directional content, cell margins, nested relationships, revisions, and formulas/fields. A
flattened fallback is suitable for read-only display and model context, never as an editable tree.
Nested-table writes are NO-GO until an immutable source anchor and round-trip verifier exist.

### Image fidelity

Upstream handles inline/floating images, crop and wrap metadata, relationship fallbacks, TIFF through
`utif2`, and EMF/WMF through a vendored and locally patched `emf-converter`. It retains originals and
uses rendered derivatives for display, which is the correct authority split.

Risks include decoder memory exhaustion, malformed record loops, alpha/color-profile differences,
unsupported raster formats, SVG/script/external references, canvas nondeterminism, huge dimensions,
metafile font substitution, and silent partial rendering. The vendored converter has extensive
local WMF/EMF changes, creating an ongoing security and maintenance fork. Conversion output must
never replace original media unless a separate explicit conversion operation is approved.

## Authority review

| Capability          | Permitted authority                                                                        | Explicitly denied                                                                |
| ------------------- | ------------------------------------------------------------------------------------------ | -------------------------------------------------------------------------------- |
| Preservation parser | Read bounded entries from the already-open DOCX; retain opaque parts                       | network, filesystem traversal, external relationship fetch                       |
| Header/footer edit  | Rewrite one resolved internal part plus required internal relationships after confirmation | shared-part fan-out without preview, arbitrary OOXML, unrelated section mutation |
| Table projection    | Parse bounded structure; expose deeper content as read-only remnants                       | recursion without caps, editing flattened remnants                               |
| Image display       | Decode bounded internal bytes to an in-memory derivative                                   | executable codecs, remote images, decoder-selected files/URLs                    |
| Pagination/print    | Render local document projection and invoke the existing system print API                  | silent background printing, print-to-arbitrary-path APIs, DOCX mutation          |

No stage changes Agent permissions. If Agent tools are later added, mutations must use the existing
proposal/confirmation and atomic-save path; print remains a user UI action.

## Dependencies, licenses, and supply-chain gate

Current foundational dependencies (`jszip`, `fast-xml-parser`, `pdf-lib`, `opentype.js`) must be
re-reviewed at the exact locked versions rather than inherited from the snapshot. The upstream image
path additionally introduces `utif2` (MIT in its locked metadata) and an Apache-2.0 vendored
`emf-converter` 2.0.2 distribution with local patches. Upstream also adds `officecrypto-tool`, but
encryption is outside this design and must not enter with image or print work.

Before any dependency lands:

1. record direct and transitive licenses, source URL, exact version/integrity, maintainer activity,
   known vulnerabilities, install scripts, native/WASM/code-generation behavior, and browser/main
   process placement;
2. reproduce notices and source obligations in the repository's notice generator;
3. fuzz hostile inputs and verify time, bytes, pixels, recursion, and cancellation caps;
4. compare a clean packaged artifact before/after on macOS arm64/x64 and Windows x64;
5. reject any codec that needs subprocess, network, dynamic import, or relaxed CSP.

The upstream selected font/vendor/corpus assets total about **14.28 MiB uncompressed in Git**; the
largest additions are CJK fonts. This is evidence that bulk font import is not an incidental image
fix. Each stage has a release budget below and must report compressed installer and unpacked app
deltas, not source-tree estimates.

## Fixture corpus and oracle

Every fixture must have provenance and redistribution permission. Confidential customer files may
be used only in a private, access-controlled evaluation set with derived metrics; they never enter
Git or diagnostics.

The corpus must include:

- Word-generated DOCX from currently supported Windows and macOS versions, plus LibreOffice and
  standards-conformant synthetic producers;
- sections with default/first/even headers and footers, linked and unlinked variants, shared parts,
  page-number/date/complex fields, tables, images, text boxes, watermarks, RTL/CJK text, and
  `AlternateContent`;
- tables covering spans, merges, autofit/fixed/percentage widths, borders, conditional styles,
  repeated header rows, row splitting, nested depth 1/4/8/9/2,000, revisions, RTL, and images;
- PNG/JPEG/GIF/BMP/TIFF, valid and malformed EMF/WMF, transparency, crop, rotation, wrap, floating
  anchors, missing and external relationships, extreme dimensions, decompression bombs, and duplicate
  media references;
- print layouts covering A4/Letter/custom sizes, portrait/landscape, margins, columns, section/page
  breaks, headers/footers, tables crossing pages, floating objects, CJK/Arabic/RTL, missing fonts,
  fields, 1/10/100/500 pages, and range expressions.

Oracles:

- untouched save: complete ZIP entry inventory plus per-entry bytes must be identical;
- targeted save: only an allowlisted part/relationship/content-type set may differ, then Word and
  LibreOffice must open without repair and a parse/save/reopen semantic verifier must pass;
- visual: deterministic PDF or page raster from Word/LibreOffice compared by page count, text
  bounding boxes, anchors, and perceptual diff. Baselines are platform/version tagged; a human
  reviews all threshold changes;
- robustness: malformed/adversarial fixtures must hit stable bounded errors within quota and leave
  the original file unchanged.

## Staged implementation and acceptance gates

### Stage 0 — Corpus, metrics, and preservation harness

Deliver only fixtures, provenance, package-diff tooling, bounded resource metrics, and Word/
LibreOffice oracle instructions. No production behavior.

Gate: at least 60 curated documents, all format families above represented, deterministic CI
inventory checks, private corpus procedure documented, and baseline packaged sizes recorded.
Release delta: test-only; **0 bytes in packaged app**. Rollback: revert the scoped test-data/tooling
commit. **GO now.**

### Stage 1 — Preservation-only OOXML hardening

Selectively reimplement namespace retention, relationship/path normalization, schema ordering, and
byte-identical no-op save invariants. Rich unsupported headers, footers, tables, and images remain
opaque and read-only.

Gate: 100% byte-identical no-op round trip; zero unexpected changed entries on all fixtures; hostile
ZIP/XML limits pass; Word and LibreOffice report no repair; no new runtime dependency; packaged delta
under **100 KiB compressed**. Rollback: feature flag selects the existing parser/save path and files
saved by the new path remain standard DOCX. **Conditional GO after Stage 0 approval.**

### Stage 2 — Header/footer read projection

Add typed default/first/even, section, rich paragraph/table/image, field, crop/wrap, and sharing
identity for display/model context only. Do not rewrite header/footer parts.

Gate: 100% text/field/variant inventory match; at least 98% geometry/visual fixtures under the
approved tolerance; no duplicate images; shared-part identity proven; bounded failures; package
delta under **300 KiB compressed**, excluding separately approved fonts. Rollback: disable projection
and retain opaque parts. **Conditional GO; editing remains NO-GO.**

### Stage 3 — Header/footer targeted editing

Permit explicitly selected variant/section edits with immutable part identity, shared-impact preview,
confirmed clone-versus-shared choice, atomic save, changed-part allowlist, semantic reopen verification,
and compare-before-restore.

Gate: every sharing/inheritance fixture has deterministic preview and postcondition; injected failures
at prepare/write/verify restore only attributable state; no third state overwritten; Word/LibreOffice
open without repair; zero unintended section changes. Rollback: turn off editing while preserving read
projection; existing documents need no migration. **NO-GO pending a separate transaction design and
Stage 2 evidence.**

### Stage 4 — Table projection, then narrowly scoped editing

First add bounded nested-table read projection, marking content below the model depth as read-only.
Only later consider cell text edits whose durable OOXML anchor, spans, revisions, and complete table
postcondition are proven; structure/style edits are separate proposals.

Gate for projection: 2,000 levels completes within fixed time/memory, modeled depth at most 8 plus
one flattened read-only remnant, all text retained, and untouched bytes identical. Gate for editing:
grid/span/merge/style/revision corpus has zero unintended XML changes and canonical reopen verification.
Package delta under **250 KiB compressed**. Rollback: disable editable projection and retain original
opaque table XML. **Conditional GO for read projection; NO-GO for editing until its transaction design.**

### Stage 5 — Raster and metafile display derivatives

Land formats independently: existing safe raster path, then TIFF, then EMF/WMF. Original relationship
bytes remain authoritative. Unsupported or failed conversion shows an explicit placeholder.

Gate per decoder: license/notices complete; no native/subprocess/network authority; fuzz corpus and
resource caps pass; cancellation works; pixel oracle meets approved tolerance on both platforms;
unsupported records are surfaced rather than reported as exact; packaged delta under **500 KiB
compressed per decoder**. Font payload is a separate decision and cannot be charged to this budget.
Rollback: per-format feature flag disables derivative rendering without changing saved files.
**Conditional GO for TIFF evaluation; NO-GO for the locally patched EMF/WMF fork until ownership,
fuzzing, and upstreaming strategy are approved.**

### Stage 6 — Pagination and print

Build page layout against settled fonts/images; use the exact mounted page projection for preview and
print; support bounded all/current/custom ranges and system cancel/failure. The UI labels output as a
preview until parity gates pass.

Gate: page count exact on at least 95% of the full corpus and 100% of the release-blocking subset;
perceptual/text-geometry thresholds approved for each platform; font/image readiness cannot change
the count after print becomes enabled; 500-page document remains within memory/time caps; range parser,
cancel, failure, and cleanup pass; packaged delta under **300 KiB compressed**, excluding approved
fonts. Manual macOS and Windows print-to-PDF comparison is mandatory. Rollback: hide print UI and use
the previous pagination view; no document migration. **NO-GO until Stages 0–2 and image/font readiness
evidence pass.**

## Release and rollback policy

Each stage is one scoped PR with its own feature flag, fixture evidence, legal report, size report,
and manual acceptance record. Flags default off for one canary release, then ramp independently.
There is no document-format migration and no background rewrite. The emergency rollback is the prior
desktop build or disabling the single stage; documents saved by an approved write stage must still
open in current Word and LibreOffice. Any repair dialog, unexplained changed ZIP entry, unbounded
resource use, silent image substitution, or print page-count instability blocks release.

## Final go/no-go matrix

| Work item                         | Decision now       | Required next evidence                                  |
| --------------------------------- | ------------------ | ------------------------------------------------------- |
| Wholesale upstream Docs snapshot  | **NO-GO**          | Not reconsidered; decompose it                          |
| Fixture/oracle harness            | **GO**             | Scoped test-only implementation review                  |
| Preservation-only OOXML hardening | **Conditional GO** | Stage 0 plus byte/part invariants                       |
| Header/footer read projection     | **Conditional GO** | Identity, sharing, bounded visual corpus                |
| Header/footer editing             | **NO-GO**          | Dedicated transaction design and Stage 2 results        |
| Nested-table read projection      | **Conditional GO** | Depth/resource and byte-preservation proof              |
| Table editing                     | **NO-GO**          | Durable anchors and complete-table verifier             |
| TIFF display derivative           | **Conditional GO** | License, fuzz, size, cross-platform pixels              |
| Patched EMF/WMF converter         | **NO-GO**          | Maintainer/upstreaming plan, fuzzing, decoder budget    |
| Pagination-backed print           | **NO-GO**          | Cross-platform page/visual parity and settled resources |

The recommended next task is Stage 0 only. It produces the evidence needed to make later decisions
without exposing users to snapshot-scale runtime risk.
