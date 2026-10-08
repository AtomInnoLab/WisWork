# ACP Agent–Client Event Design

## Design thesis

WisWork will use Agent Client Protocol (ACP) v1 as the shared semantic contract between its Agent runtime and every frontend, while retaining the current provider transports, Office Relay, host adapters, tools, and document safety controls.

The product goal is not protocol compliance for its own sake. It is to give users a coherent, observable Agent experience: streamed messages, plans, tool progress, permission decisions, cancellation, recovery, and terminal outcomes all have stable meanings across Office, Docs, Sheets, Slides, PDF, Markdown, and LaTeX.

## Scope and non-goals

In scope:

- ACP `session/update` events for assistant message chunks and tool-call lifecycle.
- Stable session, message, and tool-call identifiers.
- ACP-compatible prompt stop reasons and cancellation semantics.
- A compatibility layer so current UI callbacks continue to work during migration.
- A common frontend reducer/view model for plans, tool progress, notices, permissions, and messages.
- Progressive migration of every WisWork client to the ACP event surface.

Out of scope for this migration:

- Replacing Office Relay pairing, authentication, WebSocket framing, or deployment.
- Replacing Anthropic-compatible provider SSE or the existing model/tool loop.
- Giving ACP file-system or terminal capabilities to Office clients.
- Sending chain-of-thought. `agent_thought_chunk` is reserved for concise user-safe status only.
- Adding custom root-level ACP fields. WisWork extensions use namespaced `_meta` only.

## Architecture

`AgentLoop` remains the execution engine. `@wiswork/agent-harness` becomes the ACP boundary: it converts loop events into official SDK `SessionNotification` values and exposes those updates to clients. Existing callbacks remain an adapter during migration. Each frontend reduces the ACP event stream into its own visual presentation; it no longer infers lifecycle state from ad-hoc callback order.

For Office, ACP is carried inside the existing authenticated Relay request rather than replacing it. The Relay continues to be a content-blind tunnel and pairing authority. Once both endpoints support the event contract, Relay chunks carry ACP JSON-RPC notifications/responses; the outer Relay framing remains unchanged until a separate remote-transport decision is made.

## Canonical lifecycle

1. Client creates or restores a logical session and retains its `sessionId`.
2. A prompt begins. The Agent emits `agent_message_chunk` updates with a stable `messageId`.
3. A requested tool emits `tool_call` with `pending` or `in_progress`, a stable `toolCallId`, programmatic `name`, human title, kind, and bounded `rawInput` only where safe.
4. Progress uses `tool_call_update`; terminal status is exactly `completed` or `failed`.
5. Risky operations use `session/request_permission`; the existing proposal controller supplies the decision. Cancellation settles outstanding permission requests as `cancelled`.
6. The prompt response ends with one ACP stop reason: `end_turn`, `max_tokens`, `max_turn_requests`, `refusal`, or `cancelled`.
7. Transport failures remain JSON-RPC errors/notices, not successful stop reasons.

## Product interaction model

- The normal timeline shows assistant messages and human-readable tool titles.
- Tool rows expand to reveal safe input, progress, output, affected locations, and diagnostic code.
- A plan is a first-class checklist via `plan`/`plan_update`, not prose disguised as a tool result.
- Permission appears inline at the relevant tool row with allow/reject choices and a visible consequence.
- Cancel immediately changes the turn to “stopping,” then to ACP `cancelled`; it never appears as generic failure.
- Reconnect restores the same session where supported and reconciles in-progress tool calls before enabling retry.
- UI-specific display data such as image grids is carried in namespaced `_meta` until ACP has an equivalent standard field.

## Autonomy and trust boundaries

Read-only inspection, research, deterministic validation, and reversible formatting inside the user-authorized document scope may run automatically. Deletion, broad replacement, uncertain recovery, external publication, and writes outside the authorized scope require permission. ACP reports the decision boundary; it does not weaken existing host validation, proposal review, or rollback rules.

## Compatibility and rollout

The first release adds ACP events without changing visible behavior. Office then becomes the reference client. Remaining clients migrate one at a time. Only after all clients and diagnostics consume ACP may the legacy callbacks be removed. Each release can roll back by switching the UI consumer to the compatibility callbacks; no document migration is required.

## Success measures

- 100% of visible Agent activities map to an ACP update or prompt response.
- No tool row remains `in_progress` after completion, cancellation, error, reset, or reconnect.
- One stable tool ID connects permission, execution, diagnostics, and UI history.
- Cancellation is displayed as cancellation, not failure.
- The same recorded ACP fixture renders equivalently in at least Office, Docs, Sheets, and Slides.
- Existing model/tool behavior and document mutation verification remain byte- or behavior-compatible.

## Alternatives rejected

- Replacing Relay with ACP WebSocket now: remote ACP support is still evolving and this would mix UX semantics with security/network migration.
- Rewriting the Agent loop around the SDK immediately: high regression risk with no user-visible benefit over a typed boundary adapter.
- Keeping one custom event vocabulary per app: preserves current inconsistency and makes permissions/recovery harder to reason about.
