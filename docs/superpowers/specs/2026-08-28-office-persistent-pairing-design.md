# WisWork Office Persistent Pairing Design

## Status

Approved in conversation on 2026-08-28. This document records the agreed product behavior and the security/compatibility details discovered during implementation planning.

## Goal

Keep an Office Add-in trusted for as long as the matching WisWork PC login remains valid, so transient network loss, taskpane reloads, PC restarts, eight-hour relay session rollover, and relay restarts recover automatically without another six-digit pairing.

The first connection still requires the current six-digit code and explicit PC approval. Every restored connection still receives fresh short-lived, socket-bound session capabilities.

## Non-goals

- Persisting or reusing a relay `session_id`, session capability, Wispaper access token, or refresh token.
- Letting the relay execute the agent, retain document content, or resume when no matching signed-in PC is present.
- Sharing a trusted binding through Office document settings or a document file.
- Removing existing per-write confirmation, capability negotiation, quotas, or request limits.
- Requiring old clients to understand the new binding protocol.

## User-visible behavior

1. The first connection shows a six-digit code in Office. The user enters it in WisWork PC and approves the matching host/origin.
2. When all three components support persistent pairing, that approval also enrolls the current Office installation as trusted.
3. Taskpane reload, Office restart, PC restart, transient disconnect, session expiry, and relay restart show a bounded `Reconnecting to WisWork PC...` state and recover silently.
4. If PC is offline, Office keeps the trusted binding and shows `Waiting for WisWork PC`; it does not show a new code.
5. Explicit WisWork logout, terminal refresh-token loss, account switch, manual forget/revoke, invalid key material, or a changed capability scope invalidates silent recovery and returns to the first-pairing flow.
6. A host without persistent WebCrypto/IndexedDB support remains on the existing one-time pairing path and never stores a plaintext fallback credential.

## Binding scope and concurrency

- A binding is scoped to one Office storage profile, one Office host (`Word`, `Excel`, or `PowerPoint`), one WisWork account subject, one approved data-capability set, and the public taskpane origin.
- The Office storage profile may share one binding across documents in the same host. Concurrent documents create independent short-lived sessions from the same binding; they never share a session capability or active request.
- WisWork PC stores at most twelve binding metadata records per account, matching the current relay pool bound.

## Architecture

The system has two layers:

1. **Durable binding:** Relay stores only a binding id, the OIDC subject hash it computed itself, an Office public key, host, approved data capabilities, timestamps, and revocation state. Office stores the matching non-exportable private key in IndexedDB. PC stores only binding metadata in a separate `safeStorage`-encrypted file.
2. **Ephemeral session:** after initial approval or successful resume, Relay generates new independent Office/PC capabilities bound to the two current WebSocket connections. The existing 30-minute renewable idle lease and eight-hour absolute connection lease remain unchanged.

The relay database never stores document data, model traffic, Bearer tokens, refresh tokens, private keys, or ephemeral session capabilities.

## Cryptographic enrollment and resume

### Enrollment

- Office generates a non-exportable WebCrypto ECDSA P-256 private key. Only the public SEC1 key is sent to Relay.
- Persistent pairing is negotiated as the control feature `pairing-resume.v1`, separate from callable Agent/Web data capabilities.
- Relay creates the durable binding only after the existing authenticated PC explicitly approves the first pairing, using Relay's OIDC `sub` hash as owner.
- Database commit must succeed before either endpoint is told enrollment succeeded. If it fails, the current short-lived session may continue without persistence, and both clients report that the connection was not remembered.

### Resume

1. Office sends the binding id and host. Relay loads a live binding and returns a random, connection-bound, single-use challenge with a 30-second deadline.
2. Office signs a domain-separated transcript containing protocol name, binding id, challenge, exact origin, and host. Relay verifies it with the stored public key.
3. PC independently connects with a freshly loaded Wispaper Bearer token and requests the same binding id. Relay compares the authenticated OIDC subject with the durable owner.
4. Only after both proofs succeed does Relay create a normal in-memory Session with new random capabilities.

Challenges are random, connection-bound, single-use, limited per connection/IP, and removed on success, timeout, or disconnect. A client-chosen nonce is never accepted as proof.

## Protocol compatibility

- Existing protocol v1 and v2 frame schemas and behavior remain accepted unchanged.
- Enhanced v2 enrollment is opt-in through `pairing-resume.v1`. The control feature is never accepted as an `office.request` capability.
- A new Office client first attempts enhanced v2. If an old Relay rejects the enhanced exact schema, it reconnects once using existing v2 and completes an ordinary one-time pairing.
- A new Relay issues binding fields only when Office and PC both negotiated `pairing-resume.v1`; old clients therefore never receive unknown fields.
- New resume frame types fail closed on old Relay and trigger the bounded fallback to first pairing. No client treats protocol failure as authorization.

### Exact enhanced-v2 contract

`capabilities` always contains callable data capabilities only. The only control feature is the exact array `features: ["pairing-resume.v1"]`.

- Enhanced Office enrollment starts with `office.create` using the existing v2 fields plus `features` and `binding_public_key`. The public key is strict unpadded Base64url of a 65-byte uncompressed P-256 SEC1 point whose first byte is `0x04`. An enhanced `office.created` echoes `features`; a legacy response has exactly its old fields.
- Enhanced PC `pc.negotiate`, `pc.claim`, and `pc.approve` use their existing fields plus `features`; `pc.negotiated` and `pc.claimed` echo only the feature intersection.
- When both endpoints negotiated the feature, PC approval first yields exact `office.binding_offer {version,type,pairing_id,binding_id,capabilities,features}` only to Office. Relay has not yet written SQLite or created a session. Office persists the binding as pending and sends exact `office.binding_ready {version,type,pairing_id,binding_id}`; local persistence failure sends exact `office.binding_abort {version,type,pairing_id,binding_id}`.
- A matching ready frame from the original Office connection writes a non-resumable pending SQLite row and yields exact `office.binding_commit {version,type,pairing_id,binding_id}`. Office atomically activates its pending record and responds with exact `office.binding_committed {version,type,pairing_id,binding_id}`; activation failure uses the existing exact `office.binding_abort`. Relay activates the SQLite row only after the matching committed frame, and only then do initial `office.approved` and `pc.approved` add `features:["pairing-resume.v1"]` and `binding_id`. Office abort, offer/commit timeout, enrollment database/limit failure, or compensated delivery failure sends exact `office.binding_aborted {version,type,pairing_id,binding_id}` so Office deletes pending state. Abort or enrollment failure degrades to a short session whose approved frames contain explicit `features:[]` and no `binding_id`. Ordinary v1 and v2 approved schemas remain exact and unchanged.
- Office resume uses `office.resume {version,type,binding_id,host,capabilities}`, receives `office.challenge {version,type,binding_id,challenge,expires_in}`, and answers `office.prove {version,type,binding_id,challenge,signature}`. The signature is strict unpadded Base64url of the raw 64-byte P-256 ECDSA signature.
- The signed UTF-8 transcript is exactly `wiswork-office-resume-v1\n${bindingId}\n${challenge}\nhttps://office.8-216-134-194.sslip.io\n${host}`.
- A proved Office endpoint without its PC receives `office.waiting_for_pc {version,type}`. PC resume uses `pc.resume {version,type,binding_id,capabilities}` and receives `pc.waiting_for_office {version,type}` when Office proof is absent.
- A matched resume receives the existing standard v2 approved frames with callable data capabilities only. Resume state, rather than an extra frame field, identifies this path.
- Authenticated PC revocation uses `pc.revoke_binding {version,type,binding_id}` and receives `pc.binding_revoked {version,type,binding_id}`. Lost-ack retries are idempotently acknowledged only when the retained tombstone belongs to the same authenticated subject.
- Binding absence, revocation, host mismatch, or wrong subject is deliberately collapsed to `binding_unavailable`. Other stable resume errors are `invalid_proof`, `challenge_expired`, `resume_rate_limited`, `resume_limit`, `peer_unavailable`, and `capability_not_negotiated`.

## Persistence and migrations

### Relay

- SQLite schema version 1 stores durable binding metadata and public keys.
- The systemd unit uses `StateDirectory=wiswork-relay` and a mode that is writable only by the dynamic service user.
- Schema creation/migration is transactional. Unknown future schema versions fail startup without mutating the database.
- Enrollment-pending rows are excluded from resume lookup and pruned at startup, closing the crash window before final activation acknowledgement.
- In tests, an in-memory database is allowed; production persistent pairing requires an explicit database path.
- With `WISWORK_RELAY_PAIRING_RESUME=0`, Relay does not require, open, initialize, migrate, prune, or revoke against the SQLite store.

### WisWork PC

- `office-pairings.enc` is independent from `auth-session.enc` and uses Electron `safeStorage`, the existing Linux secure-backend allowlist, atomic temporary-file rename, strict schema validation, account filtering, and a twelve-record bound.
- Explicit logout/terminal auth loss deletes active records before closing sockets. If Relay is unreachable, encrypted revocation tombstones are retained and retried only after the same account signs in again.
- Normal application shutdown or network loss closes ephemeral sessions but preserves binding records.

### Office taskpane

- IndexedDB v2 stores independent `Word`, `Excel`, and `PowerPoint` slots, migrating the legacy single active record into only its recorded host. Each slot contains a schema-versioned binding and a non-exportable `CryptoKey`; it never stores a Wispaper token, relay capability, or exportable private key.
- Initial enrollment takes an expiring per-host IndexedDB lease and stages with compare-and-set. A second same-host taskpane cannot overwrite the first key, and a stranded pending key is never eligible for resume and expires locally after the enrollment window.
- Incoming WebSocket control frames are processed serially. Every asynchronous continuation checks its connection generation before publishing state, and enhanced approval is accepted only after the exact offer/ready/commit/committed sequence.
- Explicit forget first writes a durable per-host blocked marker, then atomically deletes the matching binding and marker. Reload cannot resume a blocked binding; deletion failure keeps the taskpane disconnected and presents a retry action, and cross-taskpane invalidation is broadcast only after durable deletion succeeds.
- Corrupt, expired-pending, mismatched-host, mismatched-capability, or unusable key cleanup rechecks the record identity inside one read-write transaction, so stale cleanup cannot delete a newer live binding.
- Document settings are not used, so bindings cannot travel with a shared document.

## State and failure handling

- Automatic retries use exponential backoff with jitter, one retry loop per binding, a bounded maximum delay, and cancellation on explicit disconnect/unmount.
- `waiting_for_pc` preserves the binding. `binding_revoked`, invalid proof, wrong subject, or capability mismatch removes the local active record and starts a fresh pairing.
- Active agent work is cancelled on socket loss exactly as today. Reconnection does not replay a request or silently assume that a write completed.
- Account status and access token are re-read before every PC resume connection.
- Protocol errors remain content-free in logs. Binding ids, public keys, challenges, signatures, and subject hashes are not logged.

## Revocation and login alignment

- Explicit logout, account switch, terminal refresh failure, or user-initiated forget revokes all matching local bindings and attempts server revocation.
- A PC that is merely closed/offline retains its encrypted records and resumes after the same stored WisWork login is available.
- Relay never accepts a PC resume solely because a user id was supplied by the client; only the subject from Relay's own OIDC validation is authoritative.
- If remote revocation cannot be delivered, loss of the PC binding record still prevents the normal client from resuming. A same-account tombstone is sent on the next authenticated connection.

## Security limits

- Persistent pairing extends availability, not authority: all existing capability checks and user confirmations remain.
- A compromised signed-in PC or compromised Office profile remains able to act within that user's approved capabilities; this design does not claim hardware attestation.
- A stolen public database does not contain a usable Office private key or Wispaper credential.
- A host that cannot prove safe key persistence falls back to one-time pairing instead of plaintext localStorage.

## Rollout and rollback

1. Deploy Relay support, database/state directory, and the feature kill switch while existing clients continue v1/v2.
2. Release WisWork PC with encrypted binding storage and resume support disabled until Relay support is confirmed.
3. Deploy taskpane support and enable `pairing-resume.v1`.
4. Observe resume success, unexpected revocation, and fallback metrics before enabling by default for all hosts.

Rollback disables Office feature advertisement and PC auto-resume. The Relay continues serving unchanged v1/v2 pairings; disabling the feature does not delete the database. Rolling back the Relay binary leaves the SQLite file untouched. Re-enabling can reuse valid bindings after schema compatibility is confirmed.

## Verification and acceptance

- Existing v1/v2 Relay, PC, and taskpane tests stay green.
- Enhanced initial pairing persists one binding only after explicit approval.
- Reload/restart/disconnect/session-expiry/Relay-restart paths create new ephemeral capabilities without a new code.
- Wrong account, wrong host, changed capability set, expired/replayed challenge, forged signature, revoked binding, corrupt stores, and unsupported WebCrypto all fail closed.
- Two concurrent Office documents can resume the same binding into distinct sessions.
- Logout/auth loss prevents further resume and clears PC active records; ordinary shutdown preserves them.
- A real-host smoke matrix covers Windows WebView2, macOS Office, and Word Web key persistence. Unsupported hosts remain on first-pairing behavior.

## Success metrics

- Silent-resume success rate after transient disconnect or reload.
- Repeat six-digit pairing rate per weekly active paired user.
- Median recovery time and unexpected binding-revocation rate.
- Resume fallback rate by Office host/platform.
- Security counters for replayed challenge, wrong-subject resume, and invalid signature, without logging identifiers.
