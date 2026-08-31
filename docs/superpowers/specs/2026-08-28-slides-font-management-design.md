# Slides Font Management Design

**Status:** Stage 1 GO; stages 2 and 3 conditional GO; font embedding NO-GO

**Scope:** WisWork PC Slides only; no Taskpane, Relay, or Office.js changes

**Reference:** `genspark/main` font registry, private-font bridge, catalog/store, and font-manager tests. This document does not authorize copying that runtime.

## Decision

Proceed in three independently releasable stages: deterministic font discovery and substitution, document-scoped private-font registration, then an optional curated download/install manager. Do not ship a public font CDN, arbitrary remote URLs, or font embedding in PPTX until the corresponding legal and round-trip gates pass.

The immediate **GO** is for stage 1. Stage 2 is **conditional GO** only after its filesystem, platform sandbox, licensing, consent, and enterprise-control gates pass. Stage 3 is **conditional GO** only after every catalog file has recorded provenance, redistribution terms, immutable hash, and a controlled CDN origin. PPTX font embedding remains **NO-GO** in this proposal.

## Problem and authority

A presentation run names a requested family, but that name is not proof that the same face is available to both layout and Chromium drawing. Today a missing, Office-private, collection-based, or script-specific font can make measurement and rendering choose different substitutes. The result is wrapping, clipping, alignment, and screenshot drift.

Authority must be explicit:

- The PPTX run/theme/placeholder font name is authoritative for intent and must remain unchanged unless the user explicitly applies a font edit.
- The exact resolved local face bytes and variable-font instance are authoritative for both metrics and drawing during the current render. Family name or filename alone is not a face identity.
- A deterministic resolver record—requested family, script/language hint, weight/style/stretch, variable-axis coordinates, resolved PostScript/face name, collection face index, source class, byte digest, parser/shaper version, platform font inventory version, and alias-table version—is authoritative for diagnostics and fixture comparison.
- Fixture expectations name an exact face digest and instance, or an explicitly approved fallback tier and digest. “Same family class” is not sufficient authority.
- System, Microsoft Office private, Apple on-demand, Office cloud-cache, WisWork-managed, and bundled fallback directories are read-only inputs except the WisWork-managed store.
- The renderer never receives arbitrary filesystem paths. Main process exposes bounded face metadata and bytes for only deck-referenced faces through typed IPC.
- Font downloads are never authoritative over a better exact locally installed face; a catalog file is accepted only after SHA-256 and family/style inspection succeed.

## Proposed architecture

### Stage 1 — deterministic resolution and measurement

Build a main-process font registry with a cheap filename/family index and lazy parsing. Resolve by normalized family, localized aliases, style, script/language hint, and platform-specific fallback. TTC/OTC handling must select a face rather than treating the collection as a single font. Variable fonts require a fully specified, clamped axis instance; an implicit platform default is not deterministic authority. The same resolved face and instance must feed the shaping/metric provider and the drawing-family decision.

This stage should learn from upstream's useful separation of discovery, lazy parsing, aliases, script classification, and fallback, but WisWork should own a smaller reviewed, versioned alias table rather than importing the full table wholesale. Alias entries require fixture evidence and an owner; changes are fidelity changes and invalidate affected goldens and caches.

Fallback is tiered and observable: exact face, approved metric-compatible substitute, same-script substitute, then bundled emergency fallback. A lower tier may keep the deck usable but cannot count as an exact-match acceptance success or conceal a missing-font regression.

### Stage 2 — document-scoped private faces

For Office-private, cloud-cache, on-demand, and WisWork-managed files that Chromium cannot resolve, main process returns only deck-referenced face IDs. The renderer requests bytes by opaque ID, creates `FontFace` objects with the correct weight/style, waits for load, and invalidates the canvas exactly once per changed face set. Close/dispose removes document-scoped faces and clears byte references.

Discovery uses a compiled per-platform root allowlist, not configurable directory strings: documented OS font roots, the current user's documented font root, explicitly supported Microsoft Office font/cache roots, explicitly supported Apple font-asset roots, and the WisWork-managed store. Every candidate is canonicalized and must remain beneath its allowed root. Directory traversal does not follow symlinks; files are opened with no-follow semantics where available, then ownership, regular-file type, permissions, device/inode, size, and canonical location are revalidated on the open descriptor before and after reading. World-writable files, unexpected owners, links, network mounts, and files that change during the read are rejected. The digest is computed from the validated bytes, not from a prior path scan, closing symlink and TOCTOU gaps.

macOS sandbox/hardened-runtime and Windows packaging must be proven on signed release artifacts. Required entitlements/capabilities are least-privilege and documented; the feature must not request broad user-selected-file or full-disk access merely to scan fonts. If Office or Apple private roots are inaccessible under the supported sandbox, Stage 2 degrades to approved system/managed sources rather than weakening the sandbox.

IPC limits:

- Maximum 128 private faces per document.
- Maximum 32 MiB total face bytes per document and 8 MiB per face.
- SFNT magic, table bounds, declared length, and family/style metadata are validated before bytes cross IPC.
- No renderer-controlled path, directory scan, or URL is accepted.
- Resolution errors degrade to a deterministic bundled fallback and produce safe diagnostics, not raw paths.
- Private-root discovery is off by default until the user accepts a local-font access explanation, and administrators can disable Office/Apple/private-font discovery independently. Denial is sticky and does not trigger repeated prompts.

### Stage 3 — curated store and local install

The optional manager lists an immutable signed catalog, displays missing deck fonts, and permits a user-initiated family download. Files land in a versioned staging directory, are checked for expected length, SHA-256, parsability, family/style match, and license metadata, then atomically renamed into the WisWork font store. Concurrent requests for one family join a single operation. Partial families are not activated.

Local install is explicit user action through a native file picker. Accept `.ttf`, `.otf`, `.ttc`, and `.otc` only after structural parsing. Preserve the original file digest and license/provenance record; generate collision-safe internal filenames instead of trusting the source basename. Do not install fonts system-wide.

## Format and fidelity risks

- Theme fonts, Latin/East Asian/complex-script attributes, symbol fonts, vertical text, and per-run language hints can select different faces. The resolver must not collapse these into one family field.
- Variable fonts and synthetic bold/italic can produce metrics that differ from PowerPoint. They remain fallback-only until golden fixtures establish parity.
- Color, bitmap, Type 1, malformed, and encrypted/protected fonts are unsupported initially and must fail closed.
- A local substitution must never rewrite the PPTX's declared font. Saving without a user font edit must preserve the original OOXML bytes/semantics.
- Applying a font edit must update Latin, East Asian, and complex-script run properties consistently; mixed-script and inherited theme runs require explicit fixtures.
- Font embedding/subsetting changes package relationships, licensing rights, and PowerPoint behavior. It is deliberately outside scope.

## Dependencies and licenses

The upstream implementation relies on `opentype.js`, HarfBuzz shaping, Electron `FontFace`, bundled Carlito, platform font directories, and a curated OFL catalog. WisWork may reuse existing compatible parsing/shaping dependencies, but adding or upgrading either library requires security and deterministic-output review.

Every bundled or downloadable font needs: upstream URL, exact version/commit, copyright holder, license text, redistribution and modification flags, file SHA-256, and generated third-party notice entry. OFL reserved font names and modified-font naming requirements must be checked per family. Microsoft/Apple/OS fonts may be read for local rendering only after legal review confirms that the intended discovery and in-process use are permitted for each supported product/platform; they must never be copied into the WisWork store, redistributed, persisted in caches beyond the document session, uploaded, or exposed to another user. An inability to establish permission makes that source class NO-GO, not a reason to infer permission. User-installed font bytes remain local and are excluded from diagnostics, telemetry, and model context.

Upstream's catalog declares 65 files totaling about 69.1 MiB uncompressed, so mirroring it is not a zero-cost dependency and must not be silently bundled.

## Fixture corpus

Fixtures must be legally redistributable or generated from licensed test fonts and cover:

- Latin regular/bold/italic/bold-italic; variable weight; ligatures and kerning.
- Simplified/traditional Chinese, Japanese, Korean, Arabic, Hebrew, Devanagari, Thai, and emoji/symbol runs.
- Mixed scripts within one line, complex-script attributes, RTL, vertical text, theme fonts, tables, charts, groups, and masters/layout placeholders.
- Missing fonts, localized family aliases, TTC/OTC faces, Office-private fonts, user fonts, corrupt fonts, hash mismatch, interrupted download, and same-family collisions.
- Cross-platform reference screenshots from current PowerPoint on supported macOS and Windows versions.

Each golden records the resolver/alias versions, exact resolved-face digest, collection index, variable-axis instance, approved fallback tier, line breaks, glyph-run bounds, slide overflow, and a pixel comparison after fonts report loaded. Fixtures that rely on proprietary local fonts run only in a licensed private lane; public CI uses OFL fixtures.

## Release size and performance budget

- Stage 1/2 code plus four existing-compatible fallback faces: target less than 5 MiB compressed installer delta and less than 25 MiB peak lazy font memory per ordinary deck.
- Initial index: less than 250 ms p95 on supported developer reference hardware; no eager parse of all system fonts.
- First referenced-face parse and renderer registration: less than 150 ms p95 per face, off the interaction-critical path where possible.
- Stage 3 catalog metadata: less than 250 KiB compressed. Font files remain on-demand; installer delta must be zero beyond metadata/UI. Default managed-font disk cap is 250 MiB with user-visible removal.
- CI must publish before/after packaged size, cold-open time, and peak resident memory. A regression above 5 MiB, 10%, or 50 MiB respectively blocks release unless separately approved.

## Rollback and observability

Ship behind `slides.font_resolution_v2`, `slides.private_fonts`, and `slides.font_store` flags. Each stage can be disabled independently without migrating documents. The managed store uses a versioned directory and atomic manifest; rollback ignores the new version but does not delete user files. If a face fails after activation, quarantine its digest and fall back for that document.

Safe metrics include resolution outcome class, script, source class, elapsed time, byte-size bucket, and fallback reason. Never record font bytes, paths, user names, deck text, or unapproved family names. A user-facing inspector may show requested and resolved family locally.

## Acceptance gates

### Stage 1 GO gate

- At least 99.5% of the general corpus uses the expected exact face digest/instance or fixture-approved fallback digest, and 100% is deterministic across three opens on the same declared platform inventory.
- The release-blocking subset covering simplified/traditional Chinese, Japanese, Korean, Arabic/RTL, Indic/complex script, and mixed-script runs must resolve to its exact expected digest/instance, preserve shaping order, and meet line-break and overflow goldens at 100%. A fallback in this subset fails the gate unless that exact fallback is the fixture's explicitly approved expectation.
- Reports separate every fallback tier; aggregate pixel or line-break scores cannot mask an exact-face, CJK, RTL, or mixed-script failure.
- No save-without-edit fixture changes declared font semantics.
- Line-break agreement with PowerPoint is at least 99% for supported fixtures; no new clipped text or overflow.
- Parser fuzzing rejects malformed files without crash, hang, or unbounded allocation.

### Stage 2 conditional GO gate

- Measurement and canvas drawing use the same face digest for every private-face fixture.
- The platform root allowlist, canonical containment, no-follow open, descriptor revalidation, ownership/permission policy, and symlink/TOCTOU adversarial tests pass on every supported OS.
- Signed packaged applications pass sandbox/hardened-runtime/entitlement tests without broad filesystem permission; inaccessible roots fail closed.
- Legal approves each Office/Apple source class, and tests prove private bytes are neither persisted nor redistributed.
- User consent, denial persistence, and enterprise-disable controls pass product, privacy, and managed-device tests before private-root discovery is enabled.
- No path-capability escape, renderer filesystem access, cross-user or cross-document face leak, or post-close retained bytes.
- The limits above are tested, and missing/failed faces reliably fall back.

### Stage 3 GO gate

- Legal inventory and generated notices are complete for every catalog file.
- HTTPS origin allowlist, immutable versioned paths, SHA-256, size, family/style, atomic install, concurrency, cancellation, and corrupt-download tests pass.
- Offline, CDN outage, and rollback leave decks editable with deterministic fallbacks.
- Product review approves consent, disk usage, removal, and enterprise-disable controls.

If any stage misses its gate, keep the previous stage enabled and mark the failing stage **NO-GO**. Font embedding stays **NO-GO** until a separate licensing and OOXML design is approved.
