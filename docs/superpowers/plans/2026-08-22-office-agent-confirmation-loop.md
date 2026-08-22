# Office Agent confirmation-loop implementation plan

## Goal and architecture

Deliver a first-class confirmation suspension in the shared AgentLoop, wire it to all Office host proposals, and rebuild the Taskpane approval/composer presentation around that state. The model must not continue while a write awaits user input; after the decision it resumes the same turn with the real result.

## Constraints

- Preserve existing proposal validation, exactly-once execution, recovery, auth, Relay, and host API-set gates.
- No raw proposal JSON in default UI.
- No manifest or deployment contract change.
- Maintain exact rollback flags and generation-safe cleanup.

## Task 1 — AgentLoop suspension protocol

Files: `packages/agent-core/src/types.ts`, `packages/agent-core/src/skill.ts`, `packages/agent-core/src/loop.ts`, `packages/agent-core/tests/loop.test.ts`.

Add the smallest typed suspension interface that lets a tool publish a waiting state and later resolve to a final `ToolExecution`. Write RED tests proving the next provider turn does not start before approve/reject, resolution resumes the same history/tool pair once, and cancel/reset prevents late continuation. Implement bounds and duplicate/failure handling. Run targeted and full agent-core tests plus typecheck. Scoped commit.

## Task 2 — Office proposal/session integration

Files: `apps/office-addin/src/agent/proposal-controller.ts`, `apps/office-addin/src/agent/use-office-agent.ts`, host skill composition as required, and focused tests.

Expose an observable pending proposal without giving UI execution authority beyond confirm/reject. Adapt Word, Excel, PowerPoint and import/media write tools to return a suspended execution through a shared wrapper. RED tests must prove the proposal is visible and actionable while AgentLoop is waiting, approval executes once then resumes, rejection does not write/retry, and logout/new task/Stop abort safely. Run focused host/session tests and Office typecheck. Scoped commit.

## Task 3 — Taskpane approval and fixed composer UX

Files: `apps/office-addin/src/App.tsx`, `apps/office-addin/src/styles.css`, UI/presentation/style tests.

Replace raw preview/code JSON with bounded host-aware summaries and human-readable targets/diffs. Confirm/Reject remain enabled during suspended waiting but all conflicting actions remain disabled during actual apply. Make the workspace grid reserve the final row for the composer, constrain the timeline to independent scrolling, and cover narrow/short panes. RED interaction/SSR/style tests first, then implementation. Scoped commit.

## Task 4 — Integration and release validation

Run all host write-path tests, Office and agent-core full suites/typechecks, production Office build, lint, format, theme, license and diff checks. Verify build artifacts contain no source maps or unsafe-eval changes. Independently review each deliverable and the combined diff; fix Critical/Important findings within two rounds. Document real Word/Excel/PowerPoint desktop manual acceptance as the remaining release gate. No manifest update unless build-contract evidence proves otherwise.

