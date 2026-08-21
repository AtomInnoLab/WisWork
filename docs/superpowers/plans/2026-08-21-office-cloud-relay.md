# Office Cloud Relay Implementation Plan

1. Build the standalone Rust relay in `services/wiswork-relay` with strict protocol parsing, in-memory pairing/session state, quotas, cancellation, health endpoint, and tests.
2. Add a WisWork PC relay client that uses the existing authenticated account boundary, supports code claim/approval UI, proxies agent streams through the existing harness, and revokes on logout.
3. Replace the Office task pane's default loopback discovery with the WSS relay session/transport while retaining the loopback bridge only as an explicit rollback mode.
4. Add nginx/systemd deployment assets, CI/build commands, environment validation, documentation, and deploy the relay behind the existing Office TLS origin.
5. Run focused and full verification, independent security review, then smoke-test desktop Word/Excel/PowerPoint and Word Web.
