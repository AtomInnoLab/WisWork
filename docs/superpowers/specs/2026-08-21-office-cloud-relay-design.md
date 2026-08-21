# WisWork Office Cloud Relay Design

## Goal

Connect desktop and web Office task panes to the signed-in WisWork PC agent without relying on browser access to an HTTP loopback server.

## Deployment

- A standalone Rust service lives in `services/wiswork-relay`.
- It binds to loopback on the server and is exposed by the existing nginx TLS origin as `wss://office.8-216-134-194.sslip.io/office-relay`.
- The Office task pane and WisWork PC both create outbound secure WebSocket connections. No inbound port is opened on the user's computer.
- The existing local HTTP bridge remains behind its current feature gate as rollback only.

## Pairing protocol

1. Office opens a WebSocket with the exact deployed Origin and sends `office.create` with a supported host label.
2. Relay returns a random opaque pairing id, a random polling secret, a six-digit verification code, and a short expiry.
3. The user enters the six-digit code in WisWork PC. PC must have a valid local Wispaper session before it can claim.
4. PC sends `pc.claim` with the code. Relay returns the Office host/origin and waits for explicit approval.
5. PC sends `pc.approve` with the pairing id. Relay binds the two live sockets and sends each side an opaque, short-lived session capability.
6. Office sends bounded agent requests through that session. Relay forwards them to PC; PC uses its existing Wispaper access token, agent harness, and credits, then streams bounded events back.

The six-digit code is a human verification value, not the authorization secret. Pairing ids, polling secrets, and capabilities are independent cryptographically random values. Codes are one-use, short-lived, rate-limited, and claim attempts are capped.

## Protocol and limits

- JSON control frames use `version: 1`, exact message types, exact keys, and stable error codes.
- Client frames are exactly:
  - Office: `office.create {version,type,host}`; `office.request {version,type,session_id,capability,request_id,body}`; `office.cancel {version,type,session_id,capability,request_id}`.
  - PC: `pc.claim {version,type,verification_code}`; `pc.approve|pc.reject {version,type,pairing_id}`; `pc.start {version,type,session_id,capability,request_id,status,content_type}`; `pc.chunk {version,type,session_id,capability,request_id,sequence,data}`; `pc.done {version,type,session_id,capability,request_id}`; `pc.error {version,type,session_id,capability,request_id,code}`.
- Server frames are exactly:
  - Office: `office.created {version,type,pairing_id,polling_secret,verification_code,expires_in}`; `office.approved {version,type,session_id,capability,expires_in}`; `office.rejected|office.expired|office.pc_offline`; `relay.start|relay.chunk|relay.done|relay.error` with the corresponding request/session fields. `relay.start` carries status/content type and resolves the streaming Response before any chunks; chunks are then enqueued immediately and `relay.done` closes the stream.
  - PC: `pc.claimed {version,type,pairing_id,host,origin,verification_code,expires_in}`; `pc.approved {version,type,session_id,capability,expires_in}`; `relay.request {version,type,session_id,request_id,body}`; `relay.cancel {version,type,session_id,request_id}`.
- `polling_secret` authorizes only the originating Office socket and is never sent to PC. PC and Office capabilities are independent and bound to their respective live sockets; a socket reconnect requires a new pairing in version 1.
- Binary frames are not accepted in the first version.
- Chunk `data` is strict standard Base64, decodes to at most 64 KiB, and `sequence` starts at zero and increases by one without gaps.
- Control frame: 16 KiB maximum.
- Agent request: 256 KiB maximum; streamed response: 16 MiB maximum; chunk: 64 KiB maximum.
- One active agent request per session; 120 second request deadline.
- Pairing TTL: 120 seconds; approved session TTL: 30 minutes, renewed only while both endpoints remain connected.
- Disconnect, logout, rejection, timeout, or protocol violation revokes the session and cancels active work.
- Relay stores only in-memory connection state and never logs document content, secrets, capabilities, access tokens, or upstream bodies.

## Trust boundaries

- nginx terminates public TLS; relay trusts only its local reverse proxy and validates the WebSocket Origin itself.
- Office is unauthenticated before pairing. Therefore pending requests are never broadcast to PCs; the user explicitly enters the code on one signed-in PC.
- Relay does not receive Wispaper credentials. PC remains the identity, billing, and model-execution boundary.
- TLS protects the first version from network observers. Application-layer end-to-end encryption is deferred and must be added before treating the relay operator as untrusted.

## Failure behavior

- Unknown origin, malformed frame, duplicate identifier, invalid/expired code, excess attempts, oversize data, or unsupported version fails closed with a stable error.
- PC offline shows `waiting_for_pc`; Office may retry within the pairing deadline.
- PC logout or auth loss closes its relay session and Office returns to disconnected state.
- Relay restart drops in-memory sessions; both clients reconnect and pair again.

## Acceptance

- Desktop Word, Excel, and PowerPoint connect without mixed-content or PNA exceptions.
- Word Web also connects through the same protocol.
- A signed-out PC cannot claim or approve a pairing.
- The same six-digit code and Office host are visible on both sides before approval.
- Agent streaming, cancellation, logout revocation, size/time limits, and reconnect are covered by tests and a real Office smoke test.
