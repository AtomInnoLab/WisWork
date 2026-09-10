# Office Enhanced authorization lease renewal

Approved by the user (`允许`) after the proposed scope: only a still-connected, signed-in and policy-authorized PC renews short-lived authorization; each lease remains 15 minutes, revocation remains effective, and the whole run remains bounded to 30 minutes. This is implementation approval, not version bump, PR update, merge, or deployment approval.

## Architecture and constraints

Negotiate optional control capability `enhanced-lease.v1`, and reuse authenticated `pc.session_state` / `relay.session_state` frames. Renewal keeps runtime instance, component version, host, raw-Office permissions, policy generation and session generation identical; only expiry may advance. A generation change remains a replacement authority and still cancels old work. A lease must never revive after expiry, account change, policy change, runtime replacement, logout, explicit disconnect or session revocation.

The PC owns renewal scheduling, starting before expiry (normally five minutes remaining). It rechecks the paired account and runtime policy across asynchronous waits, keeps one renewal in flight, and clears timers when the session is disposed. Runtime renewal validates the current authority rather than calling the constructor that increments generation. Rejected, failed, stalled or late renewals remain fail-closed under the old expiry; no writes are replayed. The original request/idle/absolute lifetime bounds remain independent.

Taskpane accepts same-generation renewal only when the optional capability was actually negotiated and every non-expiry field matches. It updates expiry without aborting tools, deleting replay-protection state, clearing the conversation or cancelling pending proposals. Repeated disable of already-absent elevated tools must not cancel semantic proposals merely because a renewed immutable statement was published. Expiry and true authority replacement keep their existing cancellation behavior.

Relay adds capability negotiation and control validation using its existing authenticated forwarding. Lease-only state updates must not renew Relay idle TTL; ordinary authenticated work still renews idle TTL as before. The control capability is not a callable agent/retrieval operation.

## Compatibility and rollout

Old peers never negotiate `enhanced-lease.v1`, so they retain the original semantics and receive no same-generation lease frames. Existing persistent binding capabilities are never silently broadened: re-pair after all three components are updated to opt in. Do not migrate or overwrite stored bindings. Rollout order is Relay + Taskpane, then PC; a new approved pairing enables renewal. Roll back by deploying previous components and re-pairing without the optional capability. No deployment is performed in this repair turn.

## Verification

Tests use the real 15-minute issued statement, not a 40-minute fixture that bypasses the defect. A still-active multi-step turn and a pending approved-once write must survive renewal across the original expiry; cancel/logout/policy/account changes and late responses must not revive them. Reject absent negotiation, expired resurrection, replayed/decreasing expiry, permission/identity changes and direct use of the control capability. Keep 30-minute run, 30-minute idle and eight-hour absolute bounds. Verify native screenshot/diagnostics repairs from the previous turn remain intact. Independent review and full repository/Relay verification precede completion; actual Mac PowerPoint end-to-end visual acceptance is reported separately.
