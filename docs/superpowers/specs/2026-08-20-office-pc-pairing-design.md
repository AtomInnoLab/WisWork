# WisWork Office ↔ PC Pairing Design

## Product thesis

WisWork Office reuses the authenticated WisWork PC session and Agent transport through a consented loopback bridge, while Office.js document reads and writes remain inside the task pane and every write still requires an explicit preview confirmation.

## Goals

- Reuse the existing Wispaper user, membership, credits, refresh, and logout behavior already working in WisWork PC.
- Remove Office's independent Wispaper OAuth flow.
- Keep Wispaper access and refresh tokens inside the PC process.
- Let the existing Office AgentLoop stream messages through the PC-authenticated provider.
- Require visible PC approval before an Office origin can use the bridge.

## Non-goals

- The PC process will not execute Office.js tools or mutate Office documents.
- The task pane will never receive the PC access or refresh token.
- V1 will not support Office on a different machine from WisWork PC.
- V1 will not silently approve or persist an unbounded pairing.

## Architecture

WisWork PC starts an HTTP loopback service bound only to `127.0.0.1`. It tries an ordered, explicitly configured port pool and selects the first available port; the default pool contains 64 ports (`43127` first, then the remaining ports in `43120–43183`). The Office task pane discovers the selected port through bounded health probes before creating a short-lived pairing request. The PC UI approves or rejects it through internal IPC, and the task pane polls with an unguessable polling secret. Once approved, the task pane receives a short-lived opaque capability that can call only the local streaming messages endpoint. The bridge invokes the existing PC auth client's authenticated request path; credentials never cross the loopback boundary.

The Office AgentLoop, Office skill, proposal preview, and confirmation boundary stay in the task pane. Only provider transport moves behind the PC bridge.

## Protocol

All response bodies are bounded JSON except the streaming messages response, which preserves the existing bounded provider stream contract.

### `POST /v1/office/pairings`

Allowed only from the exact configured Office HTTPS origin. Creates a 128-bit-or-stronger random pairing ID and a separate polling secret. Returns `202` with `{ pairing_id, polling_secret, expires_in }`. The PC UI receives an internal pairing event containing the pairing ID and declared Office host label. No account data is returned.

### `GET /v1/office/pairings/:id`

Requires `Authorization: Pairing <polling_secret>`. Returns one of `pending`, `approved`, `rejected`, or `expired`. The first successful approved poll atomically rotates the pairing into a short-lived bridge capability and returns it once. Replays fail closed.

### Internal PC approval

Approval and rejection are not exposed as unauthenticated loopback HTTP routes. The shell handles them through its trusted main/renderer IPC boundary. Approval is allowed only while the PC auth client reports a valid Wispaper session.

### `POST /v1/office/messages`

Requires `Authorization: Bridge <capability>`, exact allowed Origin, JSON content type, and the existing request size/concurrency/time budgets. The PC bridge forwards the fixed WisUsage messages request through the existing authenticated PC session and streams only the provider response. It never returns credentials or upstream error bodies.

### Browser preflight

The loopback server handles `OPTIONS` with an exact origin allowlist, explicit methods/headers, `Access-Control-Allow-Private-Network: true` when requested, `Vary: Origin`, and no wildcard. Requests without an allowed `Origin` fail closed except internal health diagnostics that disclose no session state.

### `GET /v1/office/health`

Office probes only the explicitly built-in numeric loopback endpoints, in stable bounded batches. A bridge is accepted only when it returns status 200 and the exact bounded JSON object `{ "service": "wiswork-office-bridge", "version": 1 }`. The response contains no account, session, or credential data. The selected endpoint is retained for the entire pairing and messages session.

## Security and trust boundaries

- Bind to the numeric loopback address only; startup fails if binding would expose a non-loopback interface.
- Pairing IDs, polling secrets, and bridge capabilities are generated with cryptographic randomness, stored only in memory, bounded in count, and compared without early-exit string equality.
- Pairing requests expire after two minutes; bridge capabilities expire after a short configurable lifetime and are revoked on PC logout or bridge shutdown.
- The task pane stores its capability only in memory. Reload requires re-pairing.
- PC approval identifies the requesting Office origin/host and is never automatic.
- The messages proxy permits only the fixed configured WisUsage destination and stable safe errors.
- Rate, size, stream, timeout, and concurrency limits are enforced at both Office and PC boundaries.
- Office read operations may run automatically after pairing. Replace/append remain proposals requiring confirmation in Office.

## User experience states

- PC unavailable: show “Open WisWork PC” and retry.
- PC signed out: PC approval is disabled and Office shows “Sign in to WisWork PC first.”
- Pending: Office shows “Approve this connection in WisWork PC.”
- Rejected/expired: Office offers a new pairing attempt.
- Connected: normal Agent conversation appears.
- PC logout/offline: active run stops, history/proposals clear, and Office returns to disconnected state.

## Failure handling

- `EADDRINUSE` advances to the next configured loopback port. Exhausting the pool or encountering any other bind error fails closed and is surfaced in PC diagnostics without falling back to a public bind.
- Unknown origins, malformed bodies, invalid/replayed secrets, expired grants, and unsupported methods return stable status/stage-only errors.
- Upstream 401 follows the existing PC refresh-once behavior; terminal auth loss revokes Office capabilities.
- A disconnected stream never applies an Office proposal automatically.

## Rollback

The PC bridge is feature-gated and can be disabled without changing existing PC login or Agent behavior. The Office build can revert to an unavailable/connect-PC screen. All pairing state is in memory, so process restart is a complete revocation and requires no migration.

## Verification and acceptance

- Unit tests cover origin/PNA handling, loopback-only bind configuration, expiration, replay, approval authorization, logout revocation, safe errors, and bounded proxy streaming.
- Shell integration tests cover protocol lifecycle and internal approval IPC without opening real network listeners where avoidable.
- Office tests cover offline/pending/rejected/approved/logout UI session transitions and verify no credential persistence.
- Manual Windows/macOS acceptance verifies bounded WebView discovery/fetch to loopback from the deployed HTTPS task pane, PC confirmation, streaming Agent response, credit-bearing PC identity, write confirmation, logout revocation, single-port conflict recovery, and full-pool exhaustion.

## Success criteria

- A Wispaper-authenticated PC user connects Office without another Wispaper login.
- No PC credential appears in Office storage, network responses, or logs.
- Agent requests are charged and authorized as the same Wispaper user as PC.
- Document writes remain observable and confirmation-first.
