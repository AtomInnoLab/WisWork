# WisWork Relay

Standalone Rust WebSocket relay for pairing the Office task pane with a signed-in WisWork PC. It never executes the agent or stores document content.

## Development

```bash
CARGO_TARGET_DIR=/tmp/wiswork-relay-target cargo test --locked --manifest-path services/wiswork-relay/Cargo.toml
CARGO_TARGET_DIR=/tmp/wiswork-relay-target cargo clippy --locked --manifest-path services/wiswork-relay/Cargo.toml --all-targets -- -D warnings
cargo deny --manifest-path services/wiswork-relay/Cargo.toml check licenses
```

The process binds only `127.0.0.1`. `WISWORK_RELAY_PORT` defaults to `43190` and must be a decimal port from 1 through 65535 when set.

For the cross-runtime presentation business smoke, run
`WISWORK_REAL_RELAY_SMOKE=1 pnpm exec vitest run apps/shell/tests/presentation-pc-business-smoke-integration.test.ts` from the repository root. The opt-in test starts `examples/local_business_smoke.rs` on loopback with a test-only local account endpoint, then exercises the real Relay, PC client, presentation service and attachment service. It uploads and deletes only generated fixtures under a temporary PC data directory. This is a local protocol check; deployed Relay and PowerPoint host acceptance still require separate runs.

`node tools/ppt-agent-electron-real-relay-smoke.mjs` runs the same business path with the PC client and services inside a real Electron process, including Electron's native PNG decoding. On Linux it requires `xvfb-run`; it creates temporary project and attachment data and removes the uploaded fixtures. This still uses the local test Relay/account and does not exercise the PowerPoint host.

## Production

1. Build with `cargo build --release --locked --manifest-path services/wiswork-relay/Cargo.toml`.
2. Install the binary as `/opt/wiswork-relay/wiswork-relay`.
3. Install `deploy/journald@wiswork-relay.conf` as `/etc/systemd/journald@wiswork-relay.conf`, restart `systemd-journald@wiswork-relay.service`, then install `deploy/wiswork-relay.service`. The service uses the isolated `wiswork-relay` journal namespace, capped at 64 MiB persistent / 16 MiB runtime storage and seven days. Before enabling remote diagnostics, verify `journalctl --namespace=wiswork-relay --until '7 days ago'` returns no service entries and confirm the namespace limits with `systemd-analyze cat-config systemd/journald@wiswork-relay.conf`. Roll back remote uploads with `VITE_WISWORK_OFFICE_REMOTE_DIAGNOSTICS=0`; local bounded export remains available.
4. Enable the Relay service only after the journal namespace is active.
5. Install `deploy/nginx-http-limits.conf` in nginx's `http` context and include `deploy/nginx-location.conf` inside the existing Office TLS server block. `deploy/nginx-office-site.conf` is the complete configuration used by the current development server. Reload nginx only after `nginx -t` succeeds.

The public endpoint is `wss://office.8-216-134-194.sslip.io/office-relay`; the health check is `/office-relay/health`. The service validates PC Bearer tokens only against the fixed Wispaper OIDC userinfo endpoint, immediately discards them, and must never log credentials or relay payloads.

## Presentation generation capability

Relay v2 supports the optional `presentation.v1` capability. PC advertises it only when a presentation handler is configured; pairing negotiates the intersection of the Office and PC capabilities. Generation uses the existing request and response frames: JSON object requests up to 256 KiB, ordered chunks up to 64 KiB, and responses up to 16 MiB. One request may be active per paired session; independent document sessions can run concurrently. Relay does not interpret presentation payloads or persist projects. V1 remains agent-only.

Deploy this Relay allowlist update before enabling the presentation handler in PC and requesting the capability from the task pane. An older Relay filters the capability out, so clients must check negotiated capabilities before offering generation.
