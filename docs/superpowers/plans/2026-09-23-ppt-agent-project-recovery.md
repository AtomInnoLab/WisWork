# PPT project recovery / phase 2

Base: d56e86d. Approved v1.3 plan, O1 project visibility and durable recovery. Continue existing isolated worktree; preserve original working tree and existing ACP baseline.

## Goal and architecture

Expose persisted compilation history, slide inventory, honest checks and direct recovery in Taskpane. Reuse v1 receipts; add bounded status and explicit-request resume operations to PC service. A browser controller binds queries and commands to the current document and runtime epoch. React renders a compact project card. No new dependencies.

This batch does not claim per-page production, full research/brief schema, large attachment transport or actual Office rendering QA. Resume recompiles the saved full deck only when its request is pending; compiled requests return their original output. Import still uses its separate confirmed flow.

## Units

1. Durable status/resume: modify packages/project-store/src/presentation-store.ts and tests; apps/shell/src/main/presentation-service.ts and tests. Add bound request lookup/history without changing on-disk v1. Status response is bounded, without binary/asset data. Resume selects explicit requestId; same lock and idempotency as compile. RED then GREEN tests for cancelled compile -> recreated service -> resume, cross-document refusal, unknown fields and old completed results preserved. Scoped commit.
2. Browser/controller/UI: modify presentation-generation.ts to remember a validated project before sending compile (for first-response loss), support resume with explicit request ID. Add presentation-project.ts controller and presentation-project-card.tsx; wire host-runtime/App. Query status on mount/after agent work/manual refresh. Cancel, new-task/dispose, disconnect and changed-document invalidate late results. Show history and slide list, never equate compiled with verified. Test before implementation, including direct controls and stale async responses. Scoped commit.

## Contract

status request: {operation:'status',documentId,projectId}; response {projectId,title,status:'pending'|'compiled',latestRequestId,latestCompiledRequestId?:string,slideCount,slides:[{id,title}],history:[{requestId,sequence,status,slideCount}],checks?:{structure,geometry,render,sources,roundTrip}}. slides describe latest request; checks ONLY if latest is compiled. History newest first <=20.
resume request: {operation:'resume',documentId,projectId,requestId}; response same compiled artifact as compile/get. No input deck from browser on resume. Unknown request => not_found. Existing compile/get remain compatible.

## Verification/release

Independent review of server and final integration. Run targeted regressions, full npm test and typecheck, changed TS lint, addin/shell builds, diff check. Add stage report with remaining scope. No deployment or merge. New operations are additive; old PC returning invalid_request gives a user-readable upgrade message, not a false empty project. Records unchanged so rollback preserves history. All requests retain host document binding and bounds; no arbitrary file access. Stage report must include achieved work and next steps.

## Execution evidence

- Backend delivered in `2176156`, UI in `b60b519`; independent backend review found no actionable issue.
- Generation RED: first-response loss did not preserve the project; explicit resume was unsupported. Runtime RED: recovery controller absent. All targeted tests pass after wiring.
- Final review identified selection overwrite on failed restore/resume; two RED regressions reproduced it. Fix limits early persistence to compile and rechecks recovery identity around successful persistence; focused re-review accepted.
- Real end-to-end recovery executes the actual PptxGenJS compiler, then parses the recovered 8-page output. It does not constitute Office render or RoundTrip acceptance.
- Full verification and remaining scope are recorded in `docs/product/wiswork-ppt-agent-project-recovery-progress-2026-09-23.md`.
