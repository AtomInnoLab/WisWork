# PowerPoint Native Master Editing Implementation Plan

## Goal and non-goals

Implement the approved native PowerPointApi 1.10 master inspection and editing contract, with confirmation, semantic verification, and bounded recovery. Do not restore Mac master-package import, alter Relay/PC/pairing, or silently fall back to slide-local edits.

## Architecture

The adapter owns host-specific PowerPointApi reads and writes and returns normalized, serializable master state. The skill owns exact schemas, program parsing, proposal lifecycle, stale checks, orchestration, and error mapping. Existing XML master editing is renamed and remains capability-gated separately.

Global constraints: exact bounded inputs; no raw JavaScript; no image bytes in proposal previews; PowerPointApi 1.10 gate; no write without recoverable pre-state; explicit semantic verification; no overwrite of concurrent changes.

## Task 1: Public contract and normalized state

Files:

- `apps/office-addin/src/skills/powerpoint/browser-powerpoint-adapter.ts`
- `apps/office-addin/src/skills/powerpoint/powerpoint-skill.ts`
- `apps/office-addin/src/agent/host-runtime.ts`
- `apps/office-addin/tests/powerpoint-skill.test.ts`
- `apps/office-addin/tests/host-runtime.test.ts`

Deliverable: `inspect_slide_masters`, native `edit_slide_master` v2 schema, renamed `edit_slide_master_xml`, normalized master/layout/background/theme state, and VFS plumbing.

Acceptance: Mac and Windows expose the native tools when 1.10 is available; Mac never exposes XML master editing; exact schemas reject unknown/contradictory fields; inspection is bounded and contains no document package data.

TDD: add inventory/schema/inspection tests, run targeted Vitest and observe missing-tool failures, then implement the smallest state and schema surface to pass.

Commit: `feat(office): add native PowerPoint master contract`

## Task 2: Native mutation, verification, and recovery

Files:

- `apps/office-addin/src/skills/powerpoint/browser-powerpoint-adapter.ts`
- `apps/office-addin/src/skills/powerpoint/powerpoint-skill.ts`
- `apps/office-addin/tests/powerpoint-skill.test.ts`

Deliverable: solid/gradient/pattern/picture master background operations, theme-color operations, layout inheritance operations, one structured proposal, stale validation, semantic verification, and reverse recovery.

Acceptance: confirmed operations call only native PowerPointApi; stale proposals do not write; successful writes read back exactly; partial failure restores transaction-owned values in reverse order; concurrent state is not overwritten; unrecoverable original picture fills reject before write.

TDD: add one behavior test at a time for success, stale, partial failure, concurrent state, recovery failure, and cancellation; confirm RED reason; implement to GREEN; keep surrounding PowerPoint tests green.

Commit: `feat(office): execute native PowerPoint master edits`

## Task 3: Compatibility, documentation, and release verification

Files:

- `apps/office-addin/src/skills/powerpoint/powerpoint-skill.ts`
- `apps/office-addin/tests/host-runtime.test.ts`
- design and plan documents

Deliverable: capability-aware prompt and compatibility behavior, with the rejected batch-only design removed.

Acceptance: Taskpane-only release scope is explicit; full Office tests, typecheck, lint, and formatting checks pass; final diff contains no raw XML exposure or unbounded proposal fields.

Verification commands:

- `npm test --workspace @wiswork/office-addin -- --run tests/powerpoint-skill.test.ts tests/host-runtime.test.ts`
- `npm run typecheck --workspace @wiswork/office-addin`
- `npm run lint`
- `npm run format:check`
- `npm test --workspace @wiswork/office-addin`

Commit: `test(office): verify native PowerPoint master editing`

## Rollback and release

Code rollback is a Taskpane rollback to the prior static build. The old Mac behavior already hides master XML editing, so rollback does not require document migration. Deploy only `apps/office-addin/dist`; then restart/reload and pair PowerPoint. Relay and WisWork PC remain unchanged.
