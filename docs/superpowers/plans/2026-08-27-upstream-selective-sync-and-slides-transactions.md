# Upstream Selective Sync and Slides Transactions

## Goal

Selectively absorb proven fixes and Slides concepts from `genspark-ai/genoffice` without merging its branch wholesale, replacing WisWork's shared Agent Harness, or weakening confirmation and recovery boundaries.

## Non-goals

- No wholesale upstream merge or replacement of `@wiswork/agent-core` / `@wiswork/agent-harness`.
- No default analytics, BYOK/arbitrary endpoints, upstream branding, star prompts, or packaging changes.
- No raw `object[]`, arbitrary JavaScript/OOXML, or model-callable partial-write mode.
- No Office.js permission expansion or confirmation bypass.

## Architecture

Keep one shared headless Agent Harness and host-specific executors. Add a pure `@wiswork/presentation-ops` package for bounded operation, target, receipt, error, and QC contracts. Desktop Slides and Office PowerPoint implement only the subsets they can safely plan, execute, verify, and classify.

Every write uses stable slide/element identities, expected fingerprints, an atomic transaction plan, operation-specific verification, and one of four receipts: `applied`, `unchanged`, `conflict`, or `uncertain`. A third state is never overwritten.

## Global constraints

- Strict discriminated operation schemas; no catch-all payload.
- Default and only model-facing mode is atomic.
- Sequential planning must account for dependencies between operations in one transaction.
- Recovery is allowed only for state proven attributable to the transaction.
- Bounded, content-safe receipts and diagnostics; no raw document data or host error strings.
- No new network, filesystem, or executable-code authority.
- Existing per-document harness ownership, cancellation generations, and Office confirmation remain authoritative.
- Each task below is an independently testable scoped commit/PR with RED then GREEN evidence.

## Implementation sequence

### 1. Small upstream reliability fixes

Implement separately rather than cherry-picking snapshots:

1. `packages/docx-engine/src/parse.ts`: prevent entity double decoding (reference `945c370`). Add literal/numeric/mixed entity round-trip fixtures. Commit: `fix(docx): decode document entities once`.
2. `apps/markdown/src/renderer/save-coordinator.ts` and `App.tsx`: serialize save/close using revision-owned promises, not upstream's 50 ms polling (reference `f81dd3d`). Test close during save, stale completion, rejection, and unmount. Commit: `fix(markdown): serialize close with pending saves`.
3. `packages/ai-provider/src/stream.ts`: release SSE readers exactly once on success, error, timeout, and cancellation (reference `b1a01e2`). Preserve strict malformed-frame handling. Commit: `fix(ai-provider): release completed stream readers`.

Acceptance: exact data preservation, close waits for the latest durable revision, and all streams release their reader without masking the primary error.

### 2. Immediate Slides improvements

Modify `apps/slides/src/renderer/ai/layout-audit.ts` and `slides-skill.ts`:

- Add deterministic left/right overflow checks (reference `9792915`).
- Add bounded `set_speaker_notes` using existing notes IPC (reference `2e3e97f`).
- Verify notes by stable slide target and readback.

Tests cover all four edges, rotations/boundaries, stale slides, over-limit notes, cancellation, and mismatch. Commit: `feat(slides): audit horizontal overflow and edit notes`.

### 3. Canonical presentation contract

Create `packages/presentation-ops`:

- `types.ts`: discriminated `PresentationOperation`, `PresentationTarget`, transaction and receipt unions.
- `schema.ts`: strict runtime parsing and action/string/geometry caps.
- `fingerprint.ts`: stable semantic fingerprints.
- Tests for unknown fields/kinds, prototype keys, non-finite values, duplicate client IDs, bounds, and serialization.

Initial operation families: text, geometry, fill/stroke, add/delete element, and speaker notes. No generic operation branch. Commit: `feat(presentation-ops): define typed transaction contracts`.

### 4. Durable presentation targets

Extend `packages/pptx-engine` to preserve slide part identity and `a16:creationId` element identity. Mint IDs only for new or explicitly edited objects; do not rewrite untouched decks. Pair identity with expected type and fingerprint, and fail closed on duplicate/ambiguous IDs.

Tests cover save/reopen, reorder, duplicate, group/ungroup, legacy decks, missing IDs, and ambiguous IDs. Commit: `feat(pptx-engine): preserve durable presentation targets`.

### 5. Desktop atomic transaction executor

Create `apps/slides/src/main/operations/{registry,planner,executor,receipts}.ts` and typed IPC in `apps/slides/src/shared/ipc.ts` / `main/ai-ipc.ts` / preload.

Flow:

1. Resolve stable targets against one authoritative snapshot.
2. Simulate operations sequentially during planning.
3. Capture one attributed transaction snapshot.
4. Apply through a bounded registry.
5. Verify operation-specific postconditions and resulting revision.
6. Classify applied/unchanged/conflict/uncertain.
7. Restore only an exactly attributable state; never overwrite a third state.
8. Publish one receipt and one undo/history entry.

RED tests: dependent operations, mid-operation throw, abort before/during/after apply, delayed visibility, stale revision, no-op, attributable intermediate state, concurrent third state, generated IDs, and retry idempotence. Commit: `feat(slides): execute atomic presentation transactions`.

### 6. Migrate existing Slides tools by family

Compile existing tools in `slides-skill.ts` to canonical operations in this order:

1. text;
2. geometry;
3. fill/stroke;
4. add/delete;
5. tables/charts/backgrounds.

Convert `execute_slide_script` from direct serial host edits to a bounded DSL compiler. Reject unsupported combinations before proposal and remove each legacy family only after real-file acceptance. Use one PR per family: `refactor(slides): route <family> edits through transactions`.

### 7. Structured deterministic and visual QC

Add a shared bounded issue contract and update `layout-audit.ts` / `slide-qc.ts`:

- Stable issue key: slide, element, rule, severity, bounded geometry evidence.
- Compare deterministic issue sets before/after.
- Use bounded screenshots only when deterministic checks cannot decide or presentation-quality review was requested.
- Keep quality warnings distinct from write verification.
- Auto-rollback only when the transaction remains attributable and critical severity worsened.

Commit: `feat(slides): issue structured quality receipts`.

### 8. Selection-scoped edit queue

Create `apps/slides/src/renderer/ai/edit-queue.ts` and `EditQueueCard.tsx`; integrate with `AiPanel.tsx` and `App.tsx`.

- Capture stable IDs/fingerprints at enqueue time; coordinates are evidence, not authority.
- Maximum 10 targets and bounded prompt context.
- One run-level outcome and undo boundary.
- Human-readable status/conflict UI; no raw JSON.
- Stop/new task/unmount invalidates pending launches.

Tests cover selection drift, stale targets, duplicates, limits, mixed slides, cancellation, and late callbacks. Commit: `feat(slides): add selection-scoped agent edit queue`.

### 9. Selective Sheets reliability work

Implement as separate PRs using upstream only as reference:

- bounded over-cap sidecar/read batching (`bc1dceb`);
- complete Ctrl+F (`afc6711`);
- filtered-hidden row handling (`7a814db`);
- complete bounded workbook error scanning (`7eb5d59`);
- safe shell error surfacing (`9711a45`);
- cross-highlight (`cc8cff4`) only after separate UX acceptance.

Large inputs must return either complete bounded results or explicit truncation receipts, and async work must remain bound to the correct workbook identity.

### 10. Evaluate large snapshots, do not merge them

Write separate go/no-go designs for Docs print/header/footer/table/image fidelity, Sheets CSV/data validation/print settings, Slides font management, and PDF OCR/comments/conversions. Each design must cover authority, format risks, dependencies/licenses, fixture corpus, release size, and rollback. No runtime code lands until its design is approved.

### 11. Office PowerPoint semantic convergence

Only after the desktop receipt contract is stable for one release, let `apps/office-addin/src/skills/powerpoint/**` import the pure operation/target/receipt/QC contracts.

- Desktop and Office share semantics, not executors or permissions.
- Publish an explicit capability map; never advertise unsupported desktop-only operations.
- Keep proposal suspension, user confirmation, Office.js verification/recovery, diagnostics, and Relay transport unchanged.
- Test partial Office sync, delayed visibility, stale target, third state, rejection, cancellation, and unsupported operations.

Commit: `refactor(office): adopt presentation operation receipts`.

## Verification

Each behavior change must demonstrate RED then GREEN. Final release-train gates:

- `npm run format:check`
- `npm run lint`
- `npm run typecheck`
- `npm test`
- `npm run build:all`
- affected app production builds
- `git diff --check`
- theme, licenses, notices, and branding gates

Slides fixture matrix must include real PPTX decks with tables, charts, images, groups, themes, masters, notes, and non-English fonts. Manual acceptance on packaged macOS and Windows builds must cover 10+ slide tasks, selected-element editing, cancellation at every phase, stale-target conflict, notes reopen, one-step transaction undo, and save/close/reopen identity.

## Security and privacy

- Reject unknown fields, unsafe prototype keys, non-finite numbers, and over-limit inputs.
- Contracts contain no arbitrary code, URLs, local paths, credentials, or raw package payloads.
- Screenshots stay within existing model image budgets and never enter diagnostics.
- Receipts contain only IDs, operation kinds, stable codes, counts, and bounded geometry.
- Office mutations always require confirmation.
- Any dependency addition requires license/source review.

## Migration and rollback

- No file-format migration. New durable IDs remain valid if code rolls back.
- Existing tool names and chat history formats remain stable.
- Route each tool family wholly to legacy or canonical execution; never mix both inside one transaction.
- Keep an operation-family rollout switch only until real-file acceptance, then remove the legacy path.
- Elevated conflict/uncertain rates disable the affected family; never respond by skipping verification or enabling partial writes.

## Release order

1. Reliability fixes: update WisWork PC only.
2. Slides immediate improvements and desktop transaction milestones: update WisWork PC only.
3. Sheets reliability: update WisWork PC only.
4. Large capability evaluations: no runtime release.
5. Office semantic convergence: redeploy Taskpane only.

Relay and Manifest changes are not expected. Any protocol or permission change requires a separate approved plan. Production PC releases use signed macOS packaging; unsigned artifacts remain CI smoke artifacts only.

## Explicitly excluded upstream changes

- upstream AgentLoop/provider/chat UI replacement;
- default analytics/privacy changes;
- BYOK or arbitrary endpoint authority;
- raw generic ops and default `per_op` writes;
- silently ignoring malformed SSE frames;
- branding/star prompts/build/package changes;
- wholesale Docs/Sheets/Slides/PDF snapshot commits.
