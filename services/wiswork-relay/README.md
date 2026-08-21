# WisWork Relay

Standalone Rust WebSocket relay for pairing the Office task pane with a signed-in WisWork PC. It never executes the agent or stores document content.

## Development

```bash
CARGO_TARGET_DIR=/tmp/wiswork-relay-target cargo test --locked --manifest-path services/wiswork-relay/Cargo.toml
CARGO_TARGET_DIR=/tmp/wiswork-relay-target cargo clippy --locked --manifest-path services/wiswork-relay/Cargo.toml --all-targets -- -D warnings
```

The process binds only `127.0.0.1`. `WISWORK_RELAY_PORT` defaults to `43190` and must be a decimal port from 1 through 65535 when set.

## Production

1. Build with `cargo build --release --locked --manifest-path services/wiswork-relay/Cargo.toml`.
2. Install the binary as `/opt/wiswork-relay/wiswork-relay`.
3. Install `deploy/wiswork-relay.service`, then enable it with systemd.
4. Install `deploy/nginx-http-limits.conf` in nginx's `http` context and include `deploy/nginx-location.conf` inside the existing Office TLS server block. `deploy/nginx-office-site.conf` is the complete configuration used by the current development server. Reload nginx only after `nginx -t` succeeds.

The public endpoint is `wss://office.8-216-134-194.sslip.io/office-relay`; the health check is `/office-relay/health`. The service validates PC Bearer tokens only against the fixed Wispaper OIDC userinfo endpoint, immediately discards them, and must never log credentials or relay payloads.
