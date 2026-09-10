# Office production recovery plan

Goal: fix the reproduced image, remote tool, and visual-review failures, and deliver the previously requested DESIGN.md reading/editing workflow. Baseline and constraints are in the accompanying design specification.

1. **Image candidates and composition.** `apps/shell/src/main/office-retrieval-proxy.ts`, `office-image-handoff.ts`, corresponding image tests, and Office import-media files. Reproduce original decode/pixel failure skipping a valid backup, then normalize within each candidate attempt; retain source safety and bounded deadlines. Make image fit preserve proportions with regression coverage. Commit this unit independently.
2. **Remote tool lifecycle.** `apps/shell/src/main/office-relay-client.ts`, `office-codex-proxy.ts`, corresponding Relay/proxy/router integration tests. Reproduce two simultaneous reads and cancelled/late tool replies; introduce bounded sequential dispatch and deterministic cleanup. Do not replay writes on unknown outcomes. Commit this unit independently.
3. **Review recovery and native screenshot failures.** PowerPoint skill/adapter and tests. Reproduce negative re-review ignored and applied-unverified repair blocked. Bind acceptance to a fresh screenshot after repair, return useful native screenshot recovery, preserve batch expansion limits. Commit this unit independently.
4. **DESIGN.md UI and connected-PC editing.** `apps/office-addin/src/App.tsx`, styles, session wiring/tests; new bounded PC design-document handler and main wiring; Relay capability negotiation and tests. Reuse Markdown UI, place the current design entry prominently, open PC-owned files in WisWork, synchronize saves with revision/session checks and apply via existing validated contract workflow. Commit this unit independently.
5. **Integration and independent review.** Review every unit and the combined diff, fix important findings, run fresh full `npm test`, `npm run typecheck`, lint, formatting, Taskpane/PC builds, and Relay tests/clippy for affected protocol code. Record commands, results, limits and deployment dependencies here. Preserve the branch for integration.

No existing user changes are to be modified. No new dependencies unless existing capabilities cannot serve the required flow. Real-host screenshot success is not inferred from mocked Office tests.

## Completed implementation and review

- Image candidates: a decoded/pixel-limited primary no longer skips its authorized backup; normalization remains inside the same bounded deadline. Both insertion tools use cover/contain fitting; large handoffs retain at least a 960px long edge. Independent review passed 125 targeted tests and a real Chromium pixel check (center crop / transparent letterbox).
- Runtime: bounded FIFO remote tools, per-call cancellation, deterministic queued/in-flight cleanup, and authenticated retired Relay results ignored without disturbing the next request. Independent review found no outstanding Important/Critical issue.
- Review recovery: a negative re-review revokes prior acceptance; applied-but-unverified repairs need fresh valid screenshot evidence; pending writes cannot be accepted. Native empty/invalid screenshot variants try the next supported native request. Failures expose same-page recovery without allowing unreviewed batch expansion or final completion. Follow-up screenshot tests now confirm actual mock-host mutations, stale-capture rejection, and dirty final verification.
- DESIGN.md: prominent rendered reader, connected-PC-owned Markdown files, bounded capability transport, explicit apply, stale-base protection and local-file retention. Independent review found and closed rolling-timeline draft loss, closed historical reader reopening, and long-session polling exhaustion. Regression coverage includes 105 rolling events, new tasks/replaced sessions, 5,760 requests, replay rejection, pause/resume while closed, and user turns preempting read-only synchronization.
- No dependencies, version bump, PR, or production deployment were added in this repair turn.

## Verification ledger

Commands ran in the isolated worktree on 2026-09-10. Logs and browser fixtures: `/tmp/wiswork-office-recovery-checks-oy8UgU` (local temporary evidence, not shipped).

| Check                                                                   | Result                                                                                                                                                         |
| ----------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Full `npm test`                                                         | Passed; 7,621 Vitest cases at the full-run checkpoint, plus Node test suites. Later reviewed changes were re-tested in the complete affected workspaces below. |
| Final `npm run test -w @wiswork/office-addin`                           | 45 files, 965 passed.                                                                                                                                          |
| Final `npm run test -w @wiswork/shell`                                  | 61 files, 637 passed, 11 conditional tests skipped.                                                                                                            |
| Full `npm run typecheck`                                                | Passed; changed Office/PC workspaces checked again after review.                                                                                               |
| `npm run lint`                                                          | No errors; pre-existing warnings in unrelated workspaces. The new panel warning was fixed; changed-file lint is clean.                                         |
| Formatting and `git diff --check`                                       | Passed against baseline `867cc93`.                                                                                                                             |
| Taskpane / PC production builds                                         | Passed; refreshed after the reviewed runtime/UI fixes. Existing chunk-size warnings remain.                                                                    |
| `cargo test --locked --manifest-path services/wiswork-relay/Cargo.toml` | 32 unit + 55 integration tests passed, including real WebSocket late-result recovery.                                                                          |
| Relay `cargo clippy --locked --all-targets -- -D warnings`              | Passed.                                                                                                                                                        |
| Real Chromium DESIGN reader                                             | Actual React workspace at 320px and 400px: header entry visible, 44px button target, full-height scrollable reader, no horizontal overflow.                    |
| Independent review                                                      | Image, runtime/gates, and DESIGN/file-channel reviews complete; all reported Important findings closed and re-reviewed.                                        |

The initial full run exposed three legacy screenshot tests expecting raw error strings; these were updated to assert the new structured recovery and stronger evidence guards. RED checks for reviewer findings preceded their fixes. No Mac PowerPoint end-to-end acceptance was available in this Linux environment. The release/real-host checklist in the design specification remains mandatory before claiming the deployed plugin can reliably finish a complete deck.
