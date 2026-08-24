# Shared Agent Harness Implementation Plan

## Goal and non-goals

Implement the approved shared harness in `@wiswork/agent-harness` and migrate Office, Docs, Sheets, Slides, and Markdown without changing their tools, permissions, UI semantics, persistence formats, transports, or Relay protocol. Do not create shared live sessions or cross-document memory.

## Global constraints

- One harness instance per document/Taskpane.
- No document authority moves across process or network boundaries.
- Existing AgentLoop behavior and stable errors remain unchanged.
- All behavior changes use test-first RED/GREEN evidence.
- Every deliverable has a scoped commit and independent review.
- Base commit: `b70abcb`.

## Files and responsibilities

- `packages/agent-harness/`: shared lifecycle runtime, exports, tests, package metadata, TypeScript config.
- `package.json` / `package-lock.json`: workspace verification registration and dependency graph.
- `apps/office-addin/src/agent/use-office-agent.ts`: consume the harness while retaining Office presentation and confirmation policy.
- `apps/{docs,markdown,slides}/src/renderer/ai/AiPanel.tsx`, `apps/sheets/src/renderer/App.tsx`: consume the harness while retaining host UI/persistence callbacks.
- Host package manifests: declare `@wiswork/agent-harness`.
- Host tests: conformance/regression coverage for lifecycle behavior.
- `docs/`: architecture and release/rollback notes.

## Task 1 — Shared harness kernel

**Deliverable:** A dependency-light `@wiswork/agent-harness` package implementing the approved lifecycle contract above `AgentLoop`.

**Files:** create `packages/agent-harness/package.json`, `tsconfig.json`, `src/index.ts`, `src/harness.ts`, `tests/harness.test.ts`; modify root `package.json`, `package-lock.json` only as required.

**Sequence:**

1. Add failing tests for run state, empty/busy rejection, stop/reset, restore, late callback suppression, dispose, callback ordering, and two independent harnesses.
2. Implement the smallest wrapper using `AgentLoopOptions` without changing agent-core.
3. Register package tests/typecheck in root verification scripts and update the lockfile mechanically.

**Acceptance:** focused tests pass; public types compile; no DOM/Electron/Office dependency; existing agent-core tests stay green.

**Commit:** `feat(agent-harness): add shared agent lifecycle`.

## Task 2 — Markdown and Docs migration

**Deliverable:** The two lower-risk desktop editors create their interactive Agent through `createAgentHarness`, preserving chat persistence, snapshots, attachments, and autosave behavior.

**Files:** Markdown and Docs package manifests, their `AiPanel.tsx` files, focused Agent tests, and package lock if necessary.

**Sequence:**

1. Add conformance regressions for run/stop/reset and restored history/snapshots.
2. Replace direct AgentLoop refs with `AgentHarness` refs and map existing run/cancel/reset/messages/restore calls.
3. Keep project persistence, attachments, snapshot UI, and autosave in each editor.

**Acceptance:** Markdown and Docs full tests/typechecks/builds pass; transcript and rollback behavior are unchanged.

**Commit:** `refactor(editors): migrate Markdown and Docs harnesses`.

## Task 3 — Sheets and Slides migration

**Deliverable:** Sheets and the interactive Slides session instantiate the shared harness while preserving workbook/deck state, persistence, async apply behavior, quality checks, and tool registries.

**Files:** Sheets and Slides package manifests; Slides `AiPanel.tsx`; Sheets `App.tsx`; focused tests; lockfile if necessary.

**Sequence:**

1. Add one lifecycle regression per editor.
2. Replace direct interactive AgentLoop refs with `AgentHarness` refs.
3. Keep Sheets apply promises/planner fallback and Slides history/QC logic in host adapters; leave the internal Slides QC loop on agent-core.

**Acceptance:** both full test/typecheck/build suites pass; persisted transcript, autosave, rollback, and QC behavior are unchanged.

**Commit:** `refactor(editors): migrate Sheets and Slides harnesses`.

## Task 4 — Office migration

**Deliverable:** Office creates its Agent through `createAgentHarness`; confirmation, timeline, diagnostics, and lifecycle remain behaviorally identical.

**Files:** `apps/office-addin/package.json`, `src/agent/use-office-agent.ts`, `tests/agent-session.test.ts`, and package lock if necessary.

**Sequence:**

1. Add/adjust tests that assert stop preserves history, new task/reset clears it, logout/dispose blocks late callbacks, and confirmation suspension remains actionable.
2. Replace direct AgentLoop ownership with the harness and project harness lifecycle into the existing Office snapshot.
3. Keep proposal wrapping, diagnostics, and presentation events in Office code.

**Acceptance:** Office focused and full tests/typecheck/build pass; tool authority and Relay requests are byte-compatible; no UI copy or permission change.

**Commit:** `refactor(office): use shared agent harness`.

## Task 5 — Integration, documentation, and release proof

**Deliverable:** Cross-host conformance coverage, truthful architecture documentation, and a verified rollback/release checklist.

**Files:** shared integration tests and relevant README/architecture docs only.

**Sequence:**

1. Add a cross-host synthetic test proving independent histories/cancellation with different skills/transports.
2. Document shared-code versus shared-instance boundaries and deployment impact.
3. Run full tests/typechecks/builds, lint, format, theme, license, branding, and diff checks.
4. Request broad independent review; fix all Critical/Important findings and re-run affected gates.

**Acceptance:** all automated gates pass; no protocol/manifest change; manual concurrent-host acceptance checklist is explicit.

**Commit:** `docs(agent-harness): document shared runtime boundaries` (only if documentation changes remain after earlier commits).

## Security and privacy review

- Verify histories never cross harness instances.
- Verify reset/dispose prevent late callbacks and tool execution.
- Verify Office confirmation cannot be bypassed by the shared layer.
- Verify no Relay/PC capability expansion and no new network authority.
- Verify persisted transcript schema is unchanged.

## Rollback

Each host migration is independently reversible by restoring direct `AgentLoop` construction and removing its harness dependency. The kernel may remain unused safely. There is no data migration; rollback requires only rebuilding affected desktop apps or Taskpane.

## Release order

1. Merge shared kernel and all host migrations together after full verification.
2. Update WisWork PC for desktop editors and the shared package.
3. Deploy Taskpane after PC compatibility is confirmed.
4. Relay and Manifest remain unchanged.
