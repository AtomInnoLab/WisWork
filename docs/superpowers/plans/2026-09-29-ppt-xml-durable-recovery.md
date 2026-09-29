# 旧 XML 修改持久恢复实施单元（只读核对）

2026-09-29；依据原方案 §6.4/§14.2。母版原生修改交付后续，本文不计为 XML 恢复实现完成。

Scope: edit_slide_xml / edit_slide_chart / edit_slide_master_xml. Preserve tool names, version-1 declarative replace_xml schema, 1–32 unique paths, existing 32 KiB program and package limits, chart data/link protection, master layout relationship identity protection, and Mac master-XML exclusion. No implementation or expensive tests run.

## Current concrete gap

powerpoint-skill.ts proposePackageEdit validates the source ID, complete source package digest and full order, but stores before/applied only in the closure. replaceSlidePackage imports + deletes the source in one Office batch, may perform native recovery writes in catch, and for master applies the imported primary master's layouts to all old source-master pages in another batch. originalLayouts / originalLayoutIds / affectedLayouts are RAM maps. A lost callback/ACK or reopen has no persistent page/master/layout phase receipts.

## Unit 1: single-page slide/chart XML journal and independently observable SDK phases

Files: new presentation-package-change.ts and presentation-package-editing.ts; powerpoint-skill.ts forwarding; browser-powerpoint-adapter.ts or a narrow browser-presentation-package-edit-adapter.ts.

Use a dedicated document/change journal; do not force existing page v1 models whose beforeSlideIds are capped at 512. Original XML reads have complete order without that cap (600-page regressions already exist). Store full original and prepared packages/proofs in chunked paired-PC blobs; settings holds immutable refs, original native source ID, actual package source ID, original full order, source kind, current phase/pending intent, acknowledged new source ID/observed package digest and historical review refs. Use existing blob implementation, not another blob engine; add only a purpose-specific routing alias if required by capability naming.

First mutation only after all required PC savepoints are ready and reread. Persist import intent -> KeepSourceFormatting stage without deleting old source -> persist imported actual ID before later verification -> validate imported package with the existing accepted matcher and capture its actual SHA package proof -> persist delete intent -> delete old source -> prove exact final native order and target. No catch-triggered inverse writes in this durable path. Keep old adapter behavior for other unmigrated callers until their separate units.

Existing package hashes include some FNV summaries; do not use them alone as ownership proof. Preserve accepted verifyImportedPowerPointPackage behavior (including background normalization), but pair it with SHA-256 saved blob/package proofs. Existing page adapter exact replacementDigest semantics are narrower than that matcher; reuse stage/commit mechanics, not force the narrow verifier on all previously valid XML edits.

## Unit 2: master XML dependency snapshot and forward layout receipts

Files: presentation-package-change.ts master variant + new presentation-master-xml.ts / narrow SDK adapter; reuse full inspectStyleDependencies, master inventory reads, bounded XML relation parsing and package proof primitives.

Before import persist full old master/layout identity inventory, every original source-master dependent native page ID -> original layout ID and package backup, untouched/dependent package proofs, source-master layout relationship order, source package identity, full page order. Imported source gets new native page/master/layout IDs: discover from actual source page layout membership; validate new master/package relationship structure, layout count/order/content correspondence against frozen expected XML. Never map by name alone or SDK/package IDs interchangeably. Capture old-layout -> actual-new-layout mapping durably before apply.

Each applyLayout has its own intent, target current package/layout preimage, native context final business guard, readback/package proof and receipt. Skip the retained old source if it is still intended for deletion; the staged replacement already uses the imported layout. Source deletion and all dependency transitions must be journal phases regardless of sequencing. A source page verified successfully does not prove dependent pages applied. Preview/QA derives all actual affected native IDs, not verifySlides' first 20.

## Unit 3: master recovery/undo, then root history/UI integration

Files: presentation-master-xml.ts; dedicated restoration finalizer patterned on presentation-native-modify-restoration.ts; presentation-document.ts / presentation-change-history.ts / presentation-history.ts / presentation-changes.ts / presentation-changes-card.tsx / host-runtime.ts / ACP integration.

Undo must preflight every acknowledged dependent package/current layout plus source replacement/order. Reuse each saved old live layout only if the original master/layout still exists and is unchanged. Otherwise explicitly stage the saved original source package, discover and prove a reconstructed original master/layout mapping, then restore every original dependent association with per-page inverse receipts. Do not assume the old master survives after it becomes unused.

Applying the old layout is not a complete original-page proof: require each full original page package to match; when native applyLayout cannot restore page content, use original-page package restoration with exact provenance and actual restored native ID mappings, followed by the shared original master/layout association checks. Current existing-page restoration supports only single/batch provenance and has 512 limits; extend a separate XML source/large-order variant rather than reducing XML support. Final closure verifies all restored pages, complete order and all relevant dependencies. Importing/restoring page 0 alone is never master undo.

There is no existing adapter operation for deleting a newly created unused master. Record introduced master IDs and distinguish verified active page restoration from unverified leftover-master cleanup; do not claim the original complete master inventory is restored when it was not checked or cannot be removed safely.

## Uncertain outcome contract

- Persisted intent without known inserted ID: inspect only. Explicit reconciliation may claim one unique adjacent expected package under exact source/order/dependency proofs; no SDK import replay. Original full state incl master inventory must match before a user-confirmed no-write closure.
- Import persisted, deletion ACK unknown: old source present + both packages -> staged; old source absent + unique expected replacement/order -> deleted candidate. Explicit confirmation may close a receipt; never issue a second delete to resolve uncertainty.
- applyLayout ACK unknown: use full target package plus actual master/layout mapping and saved phase proofs. Field equality or page count alone cannot infer ACK, particularly no-op edits. Unknown/third-party state blocks replay and inverse writes; offer confirmed original restoration.
- Undo ACK unknown uses the same inverse-phase classification. Reopen can finish proven metadata receipts; SDK resume only starts an unstarted phase with no unresolved pending intent.
- Every SDK mutation rechecks document ID, capability, cancellation/epoch, exact journal CAS and native business preimage in the same context immediately before mutation; snapshots are copied before awaits.

## Required tests (synthetic local PC + real document binding + native SDK mocks)

Preserve current schema/allowlist tests and 600-page/end-index cases. Inject failure before backup, after backup, before intent, at import callback, after actual insert before ACK, after actual delete before ACK, after each layout step/receipt and each inverse step. Reopen inspect -> explicit reconcile/restore -> no repeated insert/delete/apply. Include two original masters with duplicate layout names/different IDs, several source-master layouts, every dependent page beyond index 20, all unchanged other-master pages, external package/layout/theme/master/order drift, original master disappearance, imported layout reorder/missing layout, absent source, ambiguous candidate, no-op XML, cancellation/disconnect/document switch/CAS after final await, source identity distinct from native ID, per-phase bounded storage reservation, screenshots historical only, exact affected-page QA. Mac has no master XML tool and no hidden fallback. Full content + package/dependency/order closure is required before claiming undone; real Office host verification remains separate.
