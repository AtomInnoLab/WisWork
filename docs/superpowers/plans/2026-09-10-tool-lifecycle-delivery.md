# Tool lifecycle follow-up implementation plan

Base: `000dfe7`, isolated worktree `/tmp/wiswork-design-draft`, branch `codex/fix-design-asset-recovery`.

## 1. Shared gateway lifecycle identity (root)

Files: `packages/codex-bridge/src/dynamic-mcp-gateway.ts`, its tests, `apps/shell/src/main/codex-runtime.ts`, and engine filtering/tests if required.

Add a bounded opaque turn ID per gateway grant to semantic tool start/end events, assign matching invocation identity to executable calls, and attach only stable enumerated failure codes to error outcomes. Prove start/end delivery on router rejection and isolation across revoked grants. Do not change execution authority.

## 2. Office lifecycle projection (Office implementer)

Files: `apps/shell/src/main/office-codex-proxy.ts`, `apps/office-addin/src/agent/transport.ts`, `use-office-agent.ts`, and their existing tests.

Drive all semantic tool attempts through lifecycle projection, keep retrieval previews as bounded enrichment, and deduplicate against the remote execution handler. Align enhanced lifecycle capacity with the host's bounded session capacity. RED cases: 33 searches; pending-write router rejection; normal execution duplicates; cancellation and stale events. GREEN preserves byte/privacy guards and existing tool behavior.

## 3. Native PC cancellation and receipts (PC implementer)

Files: `apps/shell/src/main/pc-codex-hosts.ts`, `packages/agent-runtime/src/enhanced.ts`, PC renderer/type files only where needed, `apps/slides/src/renderer/ai/agent-controller.ts`, and associated tests.

Propagate cancellation to reads, clean pending work before terminal events, forward missing gateway outcomes, and ensure Slides surfaces rejected/expired proposals without executing them. RED cases include old read resolving during the next run and duplicate local/remote receipts. Reuse existing event contracts with additive identity metadata; coordinate any shared type changes with root.

## 4. Integration and review (root + independent reviewer)

Reconcile metadata and older-client compatibility across both paths; independently review all three units. Run fresh full `npm test`, `npm run typecheck`, `npm run lint`, `npm run licenses`, formatting checks, affected builds, and `git diff --check`. Preserve shared Cargo test output mapping already configured in this worktree. Record actual skips/limits.

Integration found one additional visible-delivery defect: Docs, Sheets, and Slides replaced the last running chip instead of the matching invocation. Root adds a shared identity-based chip updater, optional identity fields in these three consumers, and regressions for parallel completion, repeated start, and no-start failures. Existing history remains readable. The PC renderer's cancellation guards make a separate `EnhancedSession` event-schema change unnecessary.

Independent review status: gateway/runtime identity unit, native PC unit, Office unit, and chip projection unit each reviewed with no remaining Important/Critical findings. The gateway test fixtures initially failed typechecking due to missing summaries; these were corrected and rechecked. Broad final review found a previously unhandled late mutation-claim settlement after cancellation/timeout. Aborted-claim guards and a handled settlement chain resolve it; six additional regressions cover cancelled/timed-out late success and failure, live failure, and invalid receipts. Changed-scope re-review and the original probe are clean (17 proxy tests pass; zero unhandled rejections). Full repository verification is rerun after this final fix.

Save scoped local commits after verification; no version bump, PR push, or deployment in this task. Revert the follow-up commit(s) to roll back without discarding prior fixes. Do not remove user worktrees or overwrite unrelated changes.

## Final verification — 2026-09-10

- `CARGO_TARGET_DIR=/mnt/cargo-target npm test`: exit 0 after the final mutation-pump fix; Vitest reports 7,323 passed and 21 existing conditional skips. Rust tests and the repository's Node policy/release gates also pass.
- `npm run typecheck`: exit 0 across all workspaces after the final fix.
- `npm run lint`: exit 0; 13 existing warnings, no errors. `npm run licenses`: exit 0.
- `npm run format`, `npm run format:check`, `npm run format:check -- --base origin/main`, and `git diff --check`: pass.
- `npm run build:all`: exit 0; the final main-process-only delta was rebuilt with `npm run build -w @wiswork/shell`, also exit 0. The initial sandbox-only Cargo lock failure was resolved by using the existing shared target with approved local build access.
- `npm run test:enhanced-seven-host`: exit 0; all seven host golden reports verified/restored.
- Independent unit and broad final reviews are clean after the cancellation-edge fix and re-review. Real gateway/session/proxy/parser tests cover the 33-search path; cancellation probes report zero unhandled rejections.

No live macOS Office acceptance, installer packaging, release, push, or deployment was performed. The source changes must be released to PC and Taskpane before users receive the full fix. No new dependencies or document migrations were added.
