# Office screenshot lifetime and terminal diagnostics

Base: `73a5476380ae35b30fd296359a7912336decec78` in the existing isolated production-recovery worktree.

## Evidence and scope

The supplied trace comes from older Taskpane build `acee1328e36e`. Its final `screenshot_slide / transport / agent_run_failed / 600000` does not identify a native screenshot error or prove a ten-minute timeout: diagnostic recording clamps elapsed duration to 600000 and inherits the most recent tool label. The original terminal cause is not recoverable from this export alone.

Two concrete current defects are reproduced: native screenshot reads do not settle when Office run/sync hangs, even after cancellation; safe transport failure categories and long-run elapsed durations are lost before reaching diagnostics. Fix these under the already-approved native screenshot recovery design in `../specs/2026-09-10-office-production-recovery-design.md`. Do not change write transaction synchronization, fabricate screenshots, weaken visual gates, retry writes, increase run deadlines, change versions, or deploy.

There is a separate confirmed lifetime mismatch: PC issues a 15-minute Enhanced statement with no renewal, while an Enhanced run may last 30 minutes. Taskpane explicitly revokes on statement expiry. This screenshot-only patch did **not** implement renewal or claim long sessions were fixed. Earlier tests use 40-minute statements and therefore miss the real authorization boundary. The user subsequently approved authenticated short-lived renewal; its separate implementation and verification are tracked in `2026-09-11-office-enhanced-lease.md`.

## Deliverables

1. **Bound native screenshot reads.** `browser-powerpoint-adapter.ts` owns one 15-second budget for run admission, native sync/export, fallback attempts and cleanup. Local cancellation abandons only the read, observes late promises, and prevents late native retries or screenshot acceptance. `powerpoint-skill.test.ts` exercises stalled admission, three sync boundaries, cleanup, caller abort, shared fallback budget, delayed timers and the real visual-review gate.
2. **Preserve safe failure evidence.** `relay/session.ts` retains allowlisted request errors and labels authorization expiry without changing revocation behavior. `transport.ts` and `use-office-agent.ts` map finite safe categories to actionable messages. Run diagnostics use `agent_run`, while screenshot tool errors retain their tool identity. `office-diagnostics.ts` records elapsed duration up to the existing Relay 24-hour bound; new local categories project to already-supported wire codes. No raw error body or document content is exported.
3. **Verify and review.** Run new tests RED before implementation, then full Office tests, Office typecheck, changed-file lint/format, production bundle build and independent review. Keep this patch local for user review; no external PR update or deployment is implied.

## Test-first evidence

- Native lifecycle tests: 11 initial expected failures, including late PNG incorrectly passing visual review. Additional delayed-timer test separately failed before adding wall-clock checkpoints. Thirteen lifecycle regressions pass after implementation.
- Diagnostics/transport/session: initial run had 13 expected failures (duration clamp and generic error mapping), then four expected Relay category/expiry failures. Additional explicit session-expiry and plain/structured screenshot-error cases guard the actual public outputs.

## Separately approved authorization design

Use an optional negotiated capability with the existing session-state frame. Only the still-authenticated PC may renew the same runtime, policy and session generation; every authority field stays unchanged and the short expiry moves forward. Renewal must not cancel in-flight tools/proposals, extend the Relay idle TTL, revive an expired authorization, or survive account/policy changes. Existing persistent bindings need explicit re-pairing to add the capability. This requires coordinated PC, Taskpane and Relay changes and its own regression/security review, recorded in the separate approved plan above.

## Screenshot-only verification completed before renewal approval

- `npm run test -w @wiswork/office-addin`: 45 files, 1002 tests passed. The first full run caught one legacy assertion expecting the old generic screenshot-read category; its precise-category expectation was updated and the complete suite rerun successfully.
- Office and Shell typechecks passed. Taskpane production build passed (existing large-chunk warning only).
- Changed-file ESLint, `npm run format:check`, and `git diff --check` passed.
- Independent review approved the final runtime changes. Review caught and tests reproduced an expiry-diagnostic change bypassing persistent reconnect; the correction preserves the existing resume path, cancels the old request exactly once, and neither replays writes nor publishes late tool results after the next authenticated handshake.
- Changes remain local and uncommitted. No PC version bump, PR update, deployment, or authorization renewal was performed.

## Rollback and limits

Revert this patch to restore the previous Taskpane behavior; no data or protocol migration is required. A native Office read cannot be forcibly killed by JavaScript, so its eventual result is ignored. If the host itself cannot render, that page still cannot pass final visual acceptance. Real macOS PowerPoint deck completion and any historical terminal cause remain unverified by automated tests alone.
