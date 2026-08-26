# Office Multi-Session Relay Design

## Goal

Allow one signed-in WisWork PC process to keep Word, Excel, and PowerPoint Office Agent sessions connected and usable at the same time. Pairing, requests, cancellation, expiry, and failures must remain isolated per Office session.

## Non-goals

- Do not change the Office taskpane wire protocol, manifest, identity, or document capabilities.
- Do not multiplex multiple Office sessions through one PC WebSocket in this release.
- Do not share request IDs, AbortControllers, capabilities, or timers between Office sessions.

## Architecture

The existing `createOfficeRelayClient` remains a single-session, fail-closed protocol implementation. A new PC-side Relay pool owns up to 12 independent child clients and WebSockets. It reserves a slot before each claim, routes approve/reject by exact pairing ID, aggregates pending requests and safe status, and revokes every child on logout, terminal authentication loss, or shutdown.

The Rust Relay already stores sessions by random session ID and binds each session to its Office connection, PC connection, and separate capabilities. It will retain that model. Its claim accounting will be corrected so a valid v2 `pc.negotiate` followed by `pc.claim` does not consume two password-guess attempts; invalid codes remain subject- and globally-rate-limited across reconnects.

## Global constraints

- Maximum 12 PC Relay slots; slot reservation is atomic across concurrent claims.
- Every child owns its socket, generation, pending pairing, session capability, active request, request replay cache, and timers.
- A child failure, Office disconnect, request cancellation, or expiry cannot clear or abort another child.
- `auth_required`, explicit logout, and shutdown revoke all children.
- Pairing approval routes only by exact Relay-provided pairing ID; unknown IDs fail closed.
- Existing IPC payloads stay unchanged. `listOfficePairings` returns the flattened pending set.
- Aggregate status never reports all sessions disconnected merely because a new claim failed while an existing session remains paired.
- No credentials, access tokens, capabilities, or request bodies enter renderer state or diagnostics.

## User flow

The user opens Word, Excel, and PowerPoint and enters each six-digit code in WisWork PC. Each claim opens an independent authenticated Relay connection. Pending approval dialogs remain a queue; approving one does not remove the other pending requests. After approval, all three taskpanes can stream Agent requests concurrently.

## Failure and rollback

- Slot failure before pairing releases only that slot.
- Reject/expiry removes only the exact pending pairing and emits one expiry/removal event.
- Network or protocol failure closes only the affected child.
- Global auth loss calls pool `revoke` and aborts every active child request.
- Rollback is a single PC construction change from pool factory back to the existing single client; the child protocol and Relay wire format are unchanged.

## Acceptance

1. Word, Excel, and PowerPoint can be claimed and approved without an earlier socket closing.
2. The three sessions can run requests concurrently, including identical request IDs, without cross-talk.
3. Cancel or disconnect of one session leaves the other two paired and usable.
4. Logout/auth loss/shutdown revokes all sessions.
5. Three v2 negotiate+claim sequences succeed within one pairing window, while repeated invalid-code guessing remains bounded.
6. Shell, Relay, typecheck, lint, format, and production build gates pass; real three-host desktop acceptance remains required before release.
