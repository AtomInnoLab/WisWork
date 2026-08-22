# Office Multi-Session Relay Plan

## Deliverable 1: PC Relay pool

Files: `apps/shell/src/main/office-relay-pool.ts`, `apps/shell/src/main/office-relay-client.ts`, `apps/shell/src/main/index.ts`, and focused Shell tests.

Implement a bounded pool of unchanged single-session clients. Reserve slots before asynchronous token/socket work; route pending/approve/reject by pairing ID; flatten pending snapshots; aggregate status; isolate ordinary failures and cancellation; globally revoke on auth loss/logout/shutdown.

RED: tests prove a second claim currently closes the first socket, and three claims cannot remain paired. Add concurrency, cancellation isolation, exact routing, capacity, failure release, and global revoke cases.

GREEN: targeted pool/client tests, full Shell tests, Shell typecheck and build.

Commit independently.

## Deliverable 2: Relay claim accounting

Files: `services/wiswork-relay/src/lib.rs`, `services/wiswork-relay/tests/relay.rs`.

Count invalid guesses, not both phases of one valid v2 negotiation. Bind a successful negotiation to the authenticated subject and pairing for the subsequent claim without weakening global or subject rate limits. Preserve expiry and disconnect cleanup.

RED: three valid v2 pairings for one subject exceed the current five-attempt limit; invalid guesses across reconnects remain bounded.

GREEN: Relay integration tests, `cargo test --locked`, `cargo clippy --locked --all-targets -- -D warnings`, and cargo-deny.

Commit independently.

## Deliverable 3: Cross-layer integration and release

Run Shell and Relay suites together. Verify multi-pending IPC and renderer queue behavior, auth-loss revocation, status diagnostics, and no protocol/manifest change. Perform independent unit reviews, fix Critical/Important findings in at most two rounds, then run a broad final review and repository gates. Push the PR, deploy Relay if its commit changed production behavior, and publish a WisWork PC build; taskpane redeployment is unnecessary unless integration changes it.
