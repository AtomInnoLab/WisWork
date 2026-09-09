# Presentation Production Recovery Plan

## Goal and non-goals

Deliver the approved recovery design without introducing a parallel deck workflow. Keep existing authorization, bounded media, and legacy contract compatibility intact.

## Global constraints

- Never log or render a secret value.
- Never translate an upstream failure into a successful empty result.
- DESIGN.md revision and page acceptance IDs are authoritative across planning, production, repair, and verification.
- All asynchronous mutations have bounded, idempotent cleanup.

## Task 1 — Image search truth and secure configuration

Files: `packages/ai-search/src/index.ts`, its tests, and the smallest existing Shell settings/IPC files needed for encrypted key save/clear/test.

Acceptance: typed provider failures reach the Slides tool; confirmed empty success remains distinguishable; environment configuration stays compatible; persisted keys are encrypted and never readable by the renderer.

Sequence: add failing backend/error propagation and secret-store tests; implement the shared search result/error contract; add save/clear/status/test integration; run ai-search and Shell tests.

Commit: `fix(search): surface failures and persist provider config`.

## Task 2 — DESIGN.md readiness and contract-bound production

Files: `packages/agent-core/src/presentation-design-workflow.ts`, `apps/slides/src/renderer/ai/slides-skill.ts`, `apps/slides/src/renderer/ai/AiPanel.tsx`, presentation design UI/sidecar tests.

Acceptance: empty DESIGN.md is never emitted; a complete ready contract with ready assets is required before build; normalized equivalent page fields bind; semantic drift still fails; repair tools remain available while a page awaits review.

Sequence: add failing workflow/UI/build-gate tests; implement readiness and normalized binding; change timeline creation behavior; run agent-core and Slides presentation tests.

Commit: `fix(slides): enforce the DESIGN.md production contract`.

## Task 3 — Screenshot, script, and transaction recovery

Files: Slides screenshot/script execution and `packages/codex-bridge/src/tool-router.ts` with their focused tests.

Acceptance: screenshot unavailability is bounded and retryable; review failure can be repaired and re-reviewed; malformed script does not mutate; all mutation terminal paths release the pending gate; later calls proceed.

Sequence: add timeout/cancel/failure reproductions; implement idempotent settlement and repair-safe gates; run Slides script/review tests and codex-bridge tests.

Commit: `fix(agent): recover presentation review transactions`.

## Final verification and release

Run full `npm test`, `npm run typecheck`, `npm run lint -- --quiet`, `npm run format:check`, and `git diff --check`. Independently review the complete diff for contract, security, concurrency, and compatibility. Build PC/Taskpane only if requested after review. PR/deployment are separate user-authorized finishing actions.
