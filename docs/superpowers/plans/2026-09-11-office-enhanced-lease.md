# Office Enhanced lease implementation plan

Goal: remove the reproduced healthy-run authorization cutoff without extending any individual authority lease or bypassing revocation. Non-goals: changing model/provider budgets, replaying mutations, rewriting pairing persistence, upgrading or deploying components.

Approved design: `../specs/2026-09-11-office-enhanced-lease-design.md`. Work in the existing isolated `codex/fix-office-production-recovery` worktree. Preserve all uncommitted screenshot-lifetime and diagnostic fixes from the previous turn; base HEAD `73a5476380ae35b30fd296359a7912336decec78`.

## Independently testable deliverables

1. **PC renewal authority and lifecycle.** Runtime `renewOfficeSessionStatement(previous)` revalidates login/policy/runtime and returns only an expiry update. `office-relay-client.ts` negotiates the optional capability, schedules bounded account-checked renewal, fences asynchronous completion, and cleans up on clear. Wire through `index.ts`; allow storing the optional capability without broadening old bindings. Files: Shell codex-runtime/client/binding-store/index and corresponding tests. RED: real 15-minute statement expires despite continued work. GREEN: same-generation renewal survives old expiry; account/policy changes, stalled renewal and cancellation remain fail-closed. Scoped commit: `fix(office): renew authenticated PC authorization leases`.
2. **Taskpane lease acceptance without cancellation.** Add capability to Relay/binding allowlists and App offer. Recognize exact same-authority renewal and rearm expiry only; preserve active tools, requests, pending proposals and replay state. Make repeated elevated-disable idempotent when already disabled. Reject using the control capability as a request. Files: Taskpane relay/session, binding-store, App, host-runtime and related tests. RED: same-generation state rejects and pending work is cancelled. GREEN: real continued work and pending proposals survive; changed/expired/unnegotiated lease is rejected. Scoped commit: `fix(office): retain active work across authorized lease renewal`.
3. **Relay compatibility and bounded control forwarding.** Add optional capability; disallow invoking control capability as a request; session-state control must not refresh idle TTL. Files: Relay `src/lib.rs`, `tests/relay.rs`; documentation and integration tests. RED then GREEN for negotiation, non-callability, unchanged expiry/idle and real WebSocket forwarding. Scoped commit: `fix(relay): negotiate short-lived authorization renewal`.

## Integration and release checks

Run targeted tests first, then `npm test`, `npm run typecheck`, lint/format, Taskpane/PC builds, Relay tests and clippy. Independently review each domain and the combined change, including live request/tool traffic crossing a 15-minute boundary, exactly-once pending mutation and late renewal after logout. Keep old pairings safe and document explicit re-pairing requirement. No DB migration. Keep local commits/worktree unless the user requests PR/release/deployment.

## Test-first and review evidence

- PC: reproduced missing runtime renewal, binding negotiation, scheduling and expiry watchdog behaviors. Tests use actual runtime-issued 15-minute statements and continued multi-step proxy/tool traffic across the original expiry. A deliberately tiny extension reproduced an unbounded scheduling loop; a one-minute minimum scheduling interval bounds malformed-short-extension attempts (normal renewal remains every ten minutes).
- Taskpane: reproduced same-generation rejection, pending semantic proposal cancellation and rejected optional binding capability. Same-authority renewal now preserves pending tools/proposals, request IDs, consumed call IDs and conversation state. Expired, replayed, altered and unnegotiated renewal remains rejected.
- Relay: reproduced missing capability negotiation and unintended idle-TTL renewal by session control. All 90 Relay tests pass, including actual WebSocket renewal while a tool result is pending; clippy passes with warnings denied.
- Independent compatibility review found that using the expanded current offer during resume invalidated old saved grants. Four new failing regressions cover old-grant resume and altered resumed approvals; the fix sends and validates the exact saved capability list without persistence mutation. Independent re-review passes.
- Independent security review reproduced runtime invalidation during the final asynchronous account lookup after renewal. Seven failing regressions cover runtime/policy invalidation and absent freshness validation. Publication now synchronously revalidates runtime-owned authority after the final await, and capability offer/scheduling require both callbacks. Independent re-review confirmed that the original real-runtime crash reproduction disconnects without publishing a stale renewal; combined review approved with no remaining concrete findings.

## Release boundary

This work does not change version numbers, update PRs, merge branches, or deploy services. To activate after a separately approved release, update Relay and Taskpane plus the PC app, then explicitly re-pair the plugin to negotiate `enhanced-lease.v1`. Existing saved pairings continue with their original grants and original no-renewal behavior. No database migration or data rewrite is needed. The 30-minute whole-run limit and visual acceptance requirements remain unchanged; actual macOS PowerPoint end-to-end deck completion is still required before claiming the reported production experience is resolved.

## Final verification

- Full `npm test` passes after the final security fix: 645 passing Vitest files, 7,731 passing tests, including Shell 681 and Taskpane 1,026. The native and repository-policy checks in the command also pass; optional real-provider/platform cases remain skipped. The first attempt needed a local-test-only sandbox escalation because loopback listeners returned `EPERM`; no production endpoint was exercised.
- Full Relay `cargo test --locked`: 34 unit tests and 56 integration tests pass. `cargo clippy --locked --all-targets -- -D warnings` and Rust formatting checks pass.
- Full repository `npm run typecheck` passes. Taskpane and PC production builds pass; Taskpane retains the existing large-chunk warning. Repository ESLint reports zero errors and 13 existing warnings; Prettier and diff whitespace checks pass.
- Independent PC, Taskpane and Relay review completed. Compatibility and runtime-revocation findings were reproduced, fixed test-first and independently rechecked. No outstanding findings in the reviewed scope.
