# Office ↔ PC Pairing Implementation Plan

## Goal and non-goals

Implement the approved loopback pairing design so the Office add-in reuses WisWork PC authentication and Agent transport. Do not add a cloud pairing service, expose PC credentials, or move Office.js document mutation into PC.

## Architecture and global constraints

Create a framework-neutral pairing/bridge package with in-memory state and strict origin/capability boundaries, integrate it into the Electron shell through loopback HTTP plus trusted IPC approval, then replace Office browser OAuth with a pairing client and PC-backed transport. Bind only `127.0.0.1`; exact CORS/PNA; no wildcard; no credential persistence in Office; stable safe errors; bounded requests/streams; writes remain confirmation-first.

## Files and responsibilities

- `packages/office-bridge/*`: pairing state machine, loopback request handler, proxy contract, limits, tests.
- `apps/shell/src/main/*`, `apps/shell/src/shared/*`, `apps/shell/src/preload/*`: bridge lifecycle, auth adapter, pairing approval IPC and types.
- `apps/office-addin/src/pc-bridge/*`: loopback client, polling session, transport adapter.
- `apps/office-addin/src/App.tsx` and tests: connection UX and Agent session lifecycle.
- package manifests, deployment manifest/config, and docs: explicit local origin/port and operational guidance.

## Task 1 — Pairing state machine and loopback handler

Deliver a standalone `@wiswork/office-bridge` package containing cryptographically random pairing creation, internal approve/reject, one-time capability issuance, expiry/replay/logout revocation, exact-origin CORS/PNA request handling, and an authenticated bounded messages proxy interface.

Acceptance: tests first fail then pass for loopback-only options, origin rejection, PNA preflight, pending/approved/rejected/expired transitions, one-time poll redemption, invalid secrets, capacity/rate bounds, capability expiry/revocation, safe proxy failures, and absence of credential fields. Commit as one scoped package deliverable.

## Task 2 — WisWork PC integration

Consume Task 1 from the Electron shell. Start/stop the bridge with the app lifecycle, adapt the existing `AuthClient`/provider request path without exposing tokens, send pending pairing details to the renderer, and add trusted IPC approve/reject actions with signed-in checks.

Acceptance: shell tests prove localhost binding configuration, lifecycle cleanup, approval requires a valid PC account, logout revokes capabilities, IPC inputs are validated, and upstream credentials/bodies never appear in responses. Existing PC authentication and Agent tests remain green. Commit as one scoped integration deliverable.

## Task 3 — Office pairing client and Agent transport

Replace independent Office OAuth with an in-memory PC bridge session. Implement health/create/poll/cancel behavior, safe connection states, capability-authenticated streaming messages, bounded parsing, cancellation, and logout/disconnect reset. Keep AgentLoop and Office tools in the task pane.

Acceptance: tests first fail then pass for PC offline, pending, rejection, expiry, approval, one-time capability handling, PNA/CORS-safe request shape, stream/cancel behavior, capability non-persistence, disconnect clearing history/proposals, and confirmation-first edits. Remove obsolete deployed OAuth entry points only when no consumer remains. Commit as one scoped Office deliverable.

## Task 4 — Deployment, compatibility, and documentation

Generate a manifest/config that permits only the deployed task-pane origin, fixed WisUsage origin where still needed, and the exact loopback endpoint required by the bridge. Document PC-required operation, port conflicts, origin configuration, revocation, and manual Windows/macOS acceptance.

Acceptance: configured/unconfigured builds fail closed appropriately; CSP/manifest contain no wildcard, placeholder, or obsolete auth callback; full root tests, typecheck, lint, formatting, build, diff checks, security searches, and manual HTTP contract probes pass. Commit as one release/readiness deliverable.

## Security review and release

Each task receives independent specification/code review; Critical and Important findings are fixed and re-reviewed. A final broad review covers cross-layer auth loss, replay, CORS/PNA, localhost exposure, token leakage, stream bounds, and Office confirmation. Rollback is disabling bridge startup and deploying the prior Office unavailable/login build; restart revokes all in-memory grants and no migration is required.
