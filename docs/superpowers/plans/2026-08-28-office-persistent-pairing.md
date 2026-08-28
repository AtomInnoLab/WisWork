# Office Persistent Pairing Implementation Plan

## Goal and non-goals

Implement the approved durable-binding/ephemeral-session design in `docs/superpowers/specs/2026-08-28-office-persistent-pairing-design.md`. Do not persist existing session capabilities, weaken write confirmations, put credentials in document settings/localStorage, or break protocol v1/v2 clients.

## Architecture and global constraints

Relay owns the authoritative OIDC subject and durable public-key binding. Office proves possession of a non-exportable WebCrypto private key; PC proves the current WisWork login on every resume. Every resume creates fresh connection-bound session capabilities and never replays interrupted requests.

Global constraints: exact frame schemas; control feature separated from callable capabilities; maximum twelve PC bindings; bounded challenge/retry state; no identifier/secret logging; transactional database migration; old-client fallback; explicit logout revocation; normal shutdown preservation.

## Files and responsibilities

- `services/wiswork-relay/src/{lib.rs,binding_store.rs,main.rs}`: feature negotiation, enrollment, challenge/proof, subject matching, ephemeral session creation, SQLite lifecycle/configuration.
- `services/wiswork-relay/tests/relay.rs`, `Cargo.toml`, `Cargo.lock`, deploy service/README: protocol, restart, replay, migration, service state and operations.
- `apps/office-addin/src/relay/{session.ts,binding-store.ts}` and `App.tsx`: WebCrypto/IndexedDB binding, enhanced enrollment, resume, auto-reconnect/fallback and UI states.
- `apps/office-addin/tests/relay-session.test.ts` plus focused store/UI tests: RED/GREEN client behavior and unsupported-host fallback.
- `apps/shell/src/main/{office-relay-client.ts,office-relay-pool.ts,office-relay-binding-store.ts,index.ts}` and shared status types: encrypted records, startup resume, backoff, initial enrollment, logout/auth-loss revocation, shutdown preservation.
- `apps/shell/tests/*office-relay*`: storage, wrong account, pool concurrency, reconnect and lifecycle tests.
- Office/Relay README and deployment documentation: changed guarantees, rollout, rollback and smoke matrix.

## Task 1: Relay durable binding and resume protocol

Acceptance: an enhanced explicitly approved pairing commits a durable public-key binding; valid Office challenge proof plus a matching freshly authenticated PC creates distinct fresh sessions across disconnects and Relay restarts. Wrong subject/host/capability, replay, forged proof, revoked binding, limits, and unknown schema fail closed. Existing v1/v2 tests remain unchanged and green.

Sequence:

1. Add failing integration tests for enhanced negotiation/enrollment, restart recovery, concurrent sessions, replay/forgery/wrong subject/revocation, legacy compatibility, and database schema failure.
2. Implement a small transactional SQLite binding store and production state-path configuration.
3. Implement control-feature segregation, conditional enhanced frames, challenge/proof and PC matching, reusing the existing Session constructor with fresh capabilities.
4. Update service hardening and operational docs; run Relay test/clippy/license checks.
5. Commit the independently reviewable Relay deliverable.

## Task 2: Office taskpane key persistence and silent resume

Acceptance: supported hosts generate/store a non-exportable P-256 key in IndexedDB, enroll once, resume automatically, preserve bindings while PC is offline, cancel retry on explicit disconnect, and fall back to ordinary pairing on old Relay or unsupported storage. No exportable private key or document setting/localStorage credential is used.

Sequence:

1. Add failing store tests and relay-session tests for enrollment, signing transcript, auto-resume/retry, old-Relay fallback, corrupt/mismatched records and explicit forget.
2. Implement dependency-injected WebCrypto/IndexedDB storage with strict schema/bounds.
3. Extend the relay state machine and App connection UI without replaying interrupted work.
4. Run Office focused suite, typecheck, production build and relevant UI tests.
5. Commit the independently reviewable taskpane deliverable.

## Task 3: WisWork PC encrypted binding lifecycle and auto-resume

Acceptance: initial enhanced approval stores bounded account-scoped metadata in `office-pairings.enc`; app/login startup resumes each record with a fresh access token; network failure retries; normal shutdown preserves records; logout/auth loss/account change deletes records and sends or queues revocation; multiple bindings keep the current twelve-session bound.

Sequence:

1. Add failing encrypted-store, client, pool and bootstrap lifecycle tests.
2. Implement the separate safeStorage store and strict account filtering/tombstones.
3. Extend client/pool APIs for initial binding capture, background resume/backoff, revoke/forget, and shutdown distinction.
4. Wire shell auth/bootstrap/status behavior and run Shell focused/full tests, typecheck and build.
5. Commit the independently reviewable PC deliverable.

## Task 4: Cross-component compatibility, release controls and documentation

Acceptance: legacy clients pair normally with the new Relay; new clients safely fall back against an old Relay; feature flags can disable enrollment/auto-resume without deleting bindings; documented deployment order and rollback are executable; real-host smoke steps cover Windows, macOS and Word Web.

Sequence:

1. Add cross-version fixtures/tests and kill-switch validation.
2. Align protocol constants/status vocabulary across all components and update user-facing reconnect states.
3. Update README, deployment environment, migration/backup/rollback and smoke checklist.
4. Run full Node and Relay verification plus diff/format/lint/security-sensitive inspections.
5. Commit integration/documentation, request broad independent review, and fix all Critical/Important findings before branch handoff.

## Migration, security and release verification

- Back up the Relay database before binary upgrade; initialize schema in a staging path and test restart before production cutover.
- Verify the systemd dynamic user can access only its `StateDirectory`; database mode is not group/world-readable.
- Verify database and logs contain no token, private key, document content, request body, session capability, challenge, signature, binding id or subject hash output.
- Deploy Relay -> PC -> taskpane. Roll back in reverse by feature flags first; preserve the SQLite database and encrypted PC store.
- Run fresh full `npm test`, `npm run typecheck`, affected production builds, Relay `cargo test --locked`, `cargo clippy --locked --all-targets -- -D warnings`, license checks where available, and the real-host smoke matrix before claiming production readiness.
