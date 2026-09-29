# 旧 XML 修改持久恢复实施计划

2026-09-29；依据原方案 §6.4/§14.2。母版原生修改交付后续；Unit 1 页面/图表本轮已实现并验证，Unit 2/3 母版 XML 待实施。

Scope: edit_slide_xml / edit_slide_chart / edit_slide_master_xml. Preserve tool names, version-1 declarative replace_xml schema, 1–32 unique paths, existing 32 KiB program and package limits, chart data/link protection, master layout relationship identity protection, and Mac master-XML exclusion. 本文建立时未实施；Unit 1 本轮已实现，最终验证与边界见阶段报告。Unit 2/3 尚未实现。

## 实施前缺口（Unit 1 已迁移页面/图表，母版 XML 仍待迁移）

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

## Unit 1 实施状态（2026-09-29）

页面/图表 XML 已接专有 package_xml 账本、presentation-package-backups.v1 六项配对 PC 操作、完整页序/全页 SHA 证明和独立 SDK 插入/删除阶段。原页与准备包先存 PC 并再次读验，持久 pending 后才写入；实际新页 ID 回执和未知结果仅观察/显式核对，无异常自动逆写。恢复按原包插入、证明、删除已应用页闭合。Document/CAS/settings、工作台、历史、准确受影响 QA、截图历史复核与 Runtime 真实确认链已接入。

导入所有条目按完整解压字节证明所有权，保留已允许的背景规范化；新增真实同长 FNV 碰撞反例，旧 matcher 接受但新所有权证明拒绝。600 页末页及 32 图表 XML 路径在合成本地 PC 测试验证；不作为真实 Office 验收。

最终证据参见 docs/product/wiswork-ppt-agent-xml-page-chart-durable-recovery-progress-2026-09-29.md。**母版 XML 尚沿用旧内存闭包，Unit 2/3 继续待实施**；不能将页面/图表恢复计为母版依赖、布局或完整母版库存恢复。
