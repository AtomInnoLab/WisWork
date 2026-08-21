# Office Agent Tool Parity Implementation Plan

## Goal and non-goals

Implement the approved Word, Excel, PowerPoint, VFS, sandbox-command, web, and installable-skill capability baseline in the Office task pane while reusing `@wiswork/agent-core`. Do not modify the PC Agent harness, expose PC capabilities, add Office OAuth/BYOK, or bypass confirmation-first writes.

## Architecture and global constraints

Compose one shared browser skill with exactly one host skill. Reads execute directly through bounded host adapters; writes create immutable generic proposals executed only by the task-pane confirmation controller. All tool inputs/results, VFS operations, script execution, Office API calls, and bridge traffic remain bounded, cancellable, host-gated, and fail closed.

## Files and responsibilities

- `apps/office-addin/src/agent/*`: shared tool validation, composed registry, generic mutation proposal integration, session lifecycle.
- `apps/office-addin/src/skills/shared/*`: VFS, `read`, sandbox commands, skill-package parsing, web command adapter.
- `apps/office-addin/src/skills/word/*`: Word tool contracts and browser Office.js adapter.
- `apps/office-addin/src/skills/excel/*`: Excel tool contracts and browser Office.js adapter.
- `apps/office-addin/src/skills/powerpoint/*`: PowerPoint tool contracts and browser Office.js adapter.
- `apps/office-addin/src/App.tsx` and task-pane components/styles: host registry and generic confirmation UI.
- `apps/office-addin/tests/*`: inventories, adapters, security limits, AgentLoop integration, confirmation behavior.
- build config, README, manifest, and design docs: feature gates, Office API requirements, deployment and manual acceptance.

## Task 1 — Shared skill runtime and generic confirmation

Create bounded exact-schema helpers, an in-memory VFS, strict `SKILL.md` parser/registry, shared `read` and sandbox-command tools, and a generic immutable proposal controller capable of previewing and confirming structured host mutations. Compose skills with the existing `composeSkills`; do not change PC harness behavior.

Acceptance: RED then GREEN tests cover VFS traversal/quota, bounded reads, skill metadata/files, forbidden sandbox globals/native access, command timeout/cancel, duplicate tool names, immutable proposals, stale state, single-flight confirmation, disconnect/logout cleanup, and a shared inventory that contains only real implemented commands. Scoped commit and independent review required.

## Task 2 — Word compatibility skill

Add Word tool definitions and adapter implementations for `get_document_text`, `get_document_structure`, `get_ooxml`, `screenshot_document`, and confirmation-gated `execute_office_js`. Add Word conversion/search command exposure only where Task 1 provides a real implementation.

Acceptance: RED then GREEN tests assert exact Word inventory, paragraph/style/list normalization, structure and OOXML bounds, screenshot feature detection, raw-code sandbox/confirmation, cancellation, stable errors, post-write verification, and absence of Excel/PowerPoint tools. Scoped commit and independent review required.

## Task 3 — Excel compatibility skill

Add all documented Excel read, range mutation, structure, resize, object, screenshot, and `eval_officejs` tools. Route every mutation through generic proposals and verify affected ranges/objects after confirmation.

Acceptance: RED then GREEN tests assert exact Excel inventory; range/value/formula/format normalization; CSV escaping; search pagination/bounds; object inventory; all mutation input schemas; impact previews; stale/confirm/cancel behavior; raw-code sandbox; stable unsupported errors; and absence of Word/PowerPoint tools. Scoped commit and independent review required.

## Task 4 — PowerPoint compatibility skill

Add all documented PowerPoint screenshot, shape, text, verification, text/XML/chart/master mutation, duplication, and raw Office.js tools. Capture bounded snapshots for XML/master changes and require post-write verification.

Acceptance: RED then GREEN tests assert exact PowerPoint inventory; shape geometry/text normalization; screenshot and verification bounds; mutation preview/confirmation; XML/chart/master/duplicate adapters; raw-code sandbox; stale/cancel behavior; and absence of Word/Excel tools. Scoped commit and independent review required.

## Task 5 — Host composition, task-pane UX, and release readiness

Select the composed skill by `Office.onReady()` host, render generic structured/code/XML previews, expose VFS uploads and installed-skill state, and preserve PC pairing/session behavior. Add feature gates and update deployment/manual-test documentation.

Acceptance: Office integration tests cover host selection, tool inventory, proposal UX, applying state, stop/logout/offline cleanup, and unsupported API messaging. Full Office/bridge/shell tests, root typecheck/lint/format/theme checks, configured/unconfigured builds, CSP/manifest/security searches, independent broad review, and Windows/macOS manual acceptance checklist are required.

## Security and dependency review

Any new evaluator, parser, archive, image, PDF, or document-conversion dependency requires license, browser-bundle, maintenance, CSP, size, and known-vulnerability review before adoption. No dynamic dependency install is allowed. Reference code may be adapted only when its license and transitive behavior are documented.

## Rollback and release

Keep the existing selection-only skill behind a build feature flag until all host acceptance checks pass. Disabling the new host/shared registries returns to the current behavior without changing PC Bridge or user identity. Release first to sideloaded desktop Office; Office Web remains unsupported until PNA and API-set acceptance passes.

## Completion phase — remove release-time unsupported gaps

The initial safe implementation exposed the reference inventory but deliberately failed closed where an audited browser implementation or reliable post-write verification was absent. Complete parity means these paths perform their documented core semantics; merely retaining a tool name that returns `office_api_unsupported` does not satisfy this phase.

### Task 6 — Shared declarative execution and browser conversion runtime

Implement a bounded, cancellable declarative Office-operation program format for the raw-code tool names. It must cover the reference project's documented examples without `eval`, `Function`, dynamic import, native shell, DOM/network/storage authority, or direct bridge access. Add audited browser-only PDF/DOCX/XLSX/CSV/image conversion and screenshot helpers where feasible, with fixed VFS inputs, byte/page/pixel/time/concurrency limits, licenses and bundle-size evidence. Web commands must use an explicit bounded PC bridge route; otherwise keep them absent rather than advertising success.

Acceptance: RED then GREEN tests cover parser/operation allowlists, authority denial, cancellation, quotas, conversion outputs, image/model transport, CSP preservation, and license/build checks. Scoped commit and independent review required.

### Task 7 — Complete Word and Excel functional parity

Word must provide real bounded screenshot output and confirmation-gated declarative execution. Excel must provide real screenshot output and every documented mutation variant, including styles, borders, notes, copy, clear formats/all, structural insert/delete, autofit, and chart/pivot create/update, with operation-specific snapshots and semantic post-write verification. Unsupported Office API sets may still return `office_api_unsupported`; supported hosts may not reject a valid documented payload merely because implementation or verification is missing.

Acceptance: contract and adapter tests exercise each documented payload family, idempotent success, stale/cancel/logout races, model-visible screenshot content, and failure-safe verification. Scoped commits and independent reviews required.

### Task 8 — Complete PowerPoint functional parity

Implement bounded slide package export/import for `edit_slide_xml`, `edit_slide_chart`, and `edit_slide_master`, preserving relationships and unrelated package parts. Implement confirmation-gated declarative execution for `execute_office_js`. Every write captures a bounded pre-state, validates the modified package/XML, re-reads affected slide/master state, and rejects false success.

Acceptance: fixture-backed package round trips, malformed XML/archive/path traversal/zip bomb bounds, chart/master relationship preservation, cancellation before irreversible writes, and semantic verification. Scoped commit and independent review required.

### Task 9 — Final integration and release gate

Remove placeholder unsupported branches for supported API sets, update skill prompts/docs to advertise only executable semantics, inspect CSP and production artifacts, and run Office/bridge/shell plus root verification. Manual Windows and macOS Word/Excel/PowerPoint acceptance remains mandatory before declaring production parity. The feature flag remains the rollback boundary.
