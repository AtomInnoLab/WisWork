# Isolated test authentication — configuration decision pending

This is a proposed next step, not deployed configuration. Do not register production `wiswork://` for a test app or copy production session files.

## Recommended first step: Dogfood only

Use an independent test OAuth client with the exact redirect URI:

```text
wiswork-dogfood://oauth/callback
```

The service owner must confirm:

- Test client ID, authorization endpoint and exact authorization-response issuer.
- This exact redirect URI is allowed for that client.
- Callback exchange and refresh endpoints, and how they select the test client. Current WisWork sends code/redirect URI to the callback endpoint, not a client ID; changing only the local client ID is insufficient.
- Whether PKCE is required. Current implementation has no challenge/verifier flow. If required, agree and implement the client-to-gateway contract before attempting login.
- Any confidential client secret stays on the server, never in the app, build metadata, diagnostic report or repository.
- Test account permissions, target model-service environment and any billing/data boundaries.

Do not send secrets in task messages. Client IDs and endpoint/issuer/redirect configuration can be supplied as non-secret configuration.

## Client-side implementation after approval

One main-process-owned identity/config selects the callback consistently for authorization, exact callback validation, the early deep-link queue, OS scheme registration, exchange and refresh. Production defaults must remain unchanged. Missing test configuration must disable test sign-in rather than fall back to production. Retain state binding, expiration, one-time consumption, issuer/query validation and profile-scoped encrypted storage.

Required regressions: cross-scheme rejection without consuming state; redirect consistency across authorization/exchange/refresh; missing config fail-closed; cold-start and second-instance callback delivery; wrong issuer and duplicate query rejection; production defaults unchanged; no session sharing.

## PR Preview decision

Do not assign all PRs one `wiswork-preview://` scheme: macOS cannot reliably route it to the initiating PR. Per-PR schemes require per-PR server registration and one installation per PR identity. A loopback callback bound to a random `127.0.0.1` port scales better but requires authorization-server/gateway support, transaction-bound redirect state and a separately reviewed local receiver. Preview login remains disabled until this choice is agreed.
