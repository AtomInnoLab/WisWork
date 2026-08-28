# WisWork Relay

Standalone Rust WebSocket relay for pairing the Office task pane with a signed-in WisWork PC. It never executes the agent or stores document content.

## Development

```bash
CARGO_TARGET_DIR=/tmp/wiswork-relay-target cargo test --locked --manifest-path services/wiswork-relay/Cargo.toml
CARGO_TARGET_DIR=/tmp/wiswork-relay-target cargo clippy --locked --manifest-path services/wiswork-relay/Cargo.toml --all-targets -- -D warnings
cargo deny --manifest-path services/wiswork-relay/Cargo.toml check licenses
```

The process binds only `127.0.0.1`. `WISWORK_RELAY_PORT` defaults to `43190` and must be a decimal port from 1 through 65535 when set. When persistent pairing is enabled, the binary requires an absolute `WISWORK_RELAY_BINDING_DB` path; tests may use an in-memory database through `Config::default()`. Setting `WISWORK_RELAY_PAIRING_RESUME=0` disables enrollment, resume, and revocation dispatch without requiring, opening, migrating, pruning, or otherwise mutating a binding database.

## Persistent pairing protocol

Protocol v1 and the original v2 schemas are unchanged. Persistent pairing is an opt-in v2 control feature and is never a callable data capability:

- Enhanced `office.create` adds `features:["pairing-resume.v1"]` and `binding_public_key`, which is unpadded base64url of a 65-byte uncompressed P-256 SEC1 point. Enhanced `pc.negotiate`, `pc.claim`, and `pc.approve` add `features`; their responses echo only the supported intersection. After explicit PC approval, Relay sends only Office an exact `office.binding_offer { pairing_id, binding_id, capabilities, features }` frame and does not yet create a session. Office persists the key as pending and answers with exact `office.binding_ready { pairing_id, binding_id }`, or with `office.binding_abort { pairing_id, binding_id }` if local persistence failed. Relay writes a non-resumable pending SQLite row and sends exact `office.binding_commit { pairing_id, binding_id }`. Office atomically activates its pending record and answers exact `office.binding_committed { pairing_id, binding_id }`; Relay then activates the SQLite row and sends enhanced `office.approved` and `pc.approved` frames containing `features:["pairing-resume.v1"]` and `binding_id`.
- Office abort, offer timeout, database/limit failure, or post-commit delivery failure produces exact `office.binding_aborted { pairing_id, binding_id }` cleanup. When both peers remain available, pre-commit abort and enrollment failure establish a short session using approved frames with explicit `features:[]` and no `binding_id`. Protocol v1 and original v2 frames remain unchanged.
- Office resumes with `office.resume { binding_id, host, capabilities }`, receives a single-use `office.challenge`, and sends `office.prove { binding_id, challenge, signature }`. Resume capability arrays are exact and callable-only; unknown, duplicate, or control-feature names fail closed. The signature is unpadded base64url of raw 64-byte P-256 ECDSA/SHA-256 over the exact UTF-8 transcript `wiswork-office-resume-v1\n${bindingId}\n${challenge}\nhttps://office.8-216-134-194.sslip.io\n${host}`.
- A freshly authenticated PC sends `pc.resume { binding_id, capabilities }`. Until its peer arrives, Relay sends `office.waiting_for_pc` or `pc.waiting_for_office`. A match receives the standard v2 approved frames with a fresh session id and fresh connection-bound capabilities; those resume-approved frames do not contain durable binding fields. An unmatched waiter receives `relay.error { code:"peer_unavailable" }` when its bounded wait expires, so a healthy socket never waits indefinitely.
- An authenticated owner revokes with `pc.revoke_binding { binding_id }` and receives `pc.binding_revoked { binding_id }`. Retrying the same revocation is idempotently acknowledged only for the same authenticated subject. Revocation also terminates live sessions for that binding. Binding lookup, host, revocation, and wrong-subject failures use the same content-free `binding_unavailable` error.

The SQLite schema stores binding ids, Relay-computed OIDC subject hashes, public keys, hosts, origins, approved data-capability lists, timestamps, bounded revocation tombstones, and non-resumable enrollment-pending rows. Startup prunes pending rows left by a crash before activation. A subject retains at most 24 recent tombstones so authenticated revocation retries survive lost acknowledgements and restarts without unbounded per-subject growth. It does not store tokens, private keys, document content, requests, challenges, signatures, session ids, or session capabilities.

## Production

1. Build with `cargo build --release --locked --manifest-path services/wiswork-relay/Cargo.toml`.
2. Install the binary as `/opt/wiswork-relay/wiswork-relay`.
3. Install `deploy/journald@wiswork-relay.conf` as `/etc/systemd/journald@wiswork-relay.conf`, restart `systemd-journald@wiswork-relay.service`, then install `deploy/wiswork-relay.service`. The unit creates `/var/lib/wiswork-relay` with mode `0700` for its dynamic user and explicitly configures `/var/lib/wiswork-relay/bindings.sqlite`. The database file is forced to mode `0600`. The service uses the isolated `wiswork-relay` journal namespace, capped at 64 MiB persistent / 16 MiB runtime storage and seven days. Before enabling remote diagnostics, verify `journalctl --namespace=wiswork-relay --until '7 days ago'` returns no service entries and confirm the namespace limits with `systemd-analyze cat-config systemd/journald@wiswork-relay.conf`. Roll back remote uploads with `VITE_WISWORK_OFFICE_REMOTE_DIAGNOSTICS=0`; local bounded export remains available.
4. Enable the Relay service only after the journal namespace is active.
5. Install `deploy/nginx-http-limits.conf` in nginx's `http` context and include `deploy/nginx-location.conf` inside the existing Office TLS server block. `deploy/nginx-office-site.conf` is the complete configuration used by the current development server. Reload nginx only after `nginx -t` succeeds.

The public endpoint is `wss://office.8-216-134-194.sslip.io/office-relay`; the health check is `/office-relay/health`. The service validates PC Bearer tokens only against the fixed Wispaper OIDC userinfo endpoint, immediately discards them, and must never log credentials or relay payloads.

Before an upgrade, stop Relay and copy `bindings.sqlite` to protected backup storage. Start the new binary against a staging copy first; startup transactionally creates schema v1 and refuses unknown future schema versions without mutation. Test one enrollment, process restart, resume, and revocation before production cutover. For rollback, first set `WISWORK_RELAY_PAIRING_RESUME=0`; keep the database untouched. A previous binary can then be restored only if it does not open or modify this file. Re-enable after schema compatibility is confirmed.

Coordinate the persistent-pairing rollout in the order Relay, WisWork PC, then taskpane. Roll back
in reverse by setting the taskpane build flag `VITE_WISWORK_OFFICE_PAIRING_RESUME=0`, the PC process
flag `WISWORK_OFFICE_PAIRING_RESUME=0`, and finally `WISWORK_RELAY_PAIRING_RESUME=0`. These switches
disable enrollment and automatic resume without deleting IndexedDB, `office-pairings.enc`, or
`bindings.sqlite`; re-enable in forward order after confirming schema compatibility. Windows
Office WebView2, macOS Office, and Word Web real-key persistence smokes remain mandatory manual
release gates and are not completed by the automated suites.
