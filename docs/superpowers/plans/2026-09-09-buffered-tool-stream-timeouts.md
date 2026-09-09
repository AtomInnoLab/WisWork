# Buffered tool stream timeout repair

## Goal and boundaries

Keep a PC/Office turn alive while its authenticated upstream is actively generating a buffered tool call or hidden reasoning. Do not dispatch incomplete tool arguments, expose private stream contents, increase existing timeout/size limits, change document authority, or automatically restart a completed user turn.

Baseline: `366568c`, PR #190, PC 0.6.59 (not yet released). Work stays in the existing isolated worktree. The user requested investigation and repair if this PC failure differs from the earlier Office failure.

## Evidence and architecture

The new PC 0.6.58 report contains six interrupted streams after a successful batch of reads, over about 261 seconds. Its retained recordings end amid input JSON deltas and are marked interrupted, not parser-rejected. Exact source arguments and full frame timing are unavailable because export is redacted and capped.

The local HTTP response currently has a 30-second socket-idle timer, while validated tool output is deliberately buffered until the complete upstream message. Upstream activity does not refresh that downstream socket timer. The PC also has a separate 60-second turn-idle timer that sees user-visible events but not buffered upstream progress.

Refresh the existing socket timer on each nonempty upstream chunk, and send an optional content-free activity callback carrying the already-bound turn ID. Refresh the PC deadline only for that exact active turn. Keep request-body, stalled-upstream, slow-client/backpressure, absolute-duration and byte/frame protections in place. This avoids adding heartbeat output, increasing limits or disabling socket protection.

## Deliverables

1. Local bridge: `packages/codex-bridge/src/local-server.ts` and `tests/local-server.test.ts`.
   - Add `onStreamActivity(turnId?: string)` as an optional host callback; observer failures are fail-open.
   - Reproduce active buffered input exceeding socket idle with a real local HTTP connection and production converter before implementing the fix.
   - Positive upstream activity refreshes the existing response/socket timer; partial input is never emitted as an executable call.
   - Verify true upstream stall, stream cancellation, and absolute request limits still terminate.
2. PC host: `apps/shell/src/main/codex-engine.ts` and relevant engine lifecycle tests.
   - Connect stream activity to the exact active turn's idle deadline, without single-document fallback for missing/stale identifiers.
   - Test matching activity lasting beyond 60 seconds, no-activity expiry, wrong/missing IDs, cancellation and no user-visible fake tool events.
3. Integration and release.
   - Independent review of the combined diff, full bridge suite and affected host suites; real pinned-runtime integration where available; type checks, PC build, formatting and diff checks.
   - Commit and update the existing open PR #190. PC remains 0.6.59; no tag or production deployment is created before merge.

## Rollback and ownership

Revert the scoped timeout-repair commit(s) to restore the preceding transport behavior; no stored-data migration is involved. Keep the earlier input-recovery and proposal-order fixes. The developer owns code verification and the PR; a live Mac PowerPoint run remains a post-install check, not a claim from local mocks.
