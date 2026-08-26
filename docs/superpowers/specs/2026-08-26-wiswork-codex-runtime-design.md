# WisWork Codex Runtime Design

## Status

Approved in the product/architecture discussion on 2026-08-26.

## Goal

Replace WisWork's model-orchestration loop with a pinned `codex app-server` runtime while preserving WisPaper as the only user identity, WisUsage as the only model gateway, and WisWork's document engines as the only mutation boundary.

## Non-goals

- Do not add ChatGPT login or user-supplied OpenAI keys.
- Do not change WisUsage.
- Do not let Codex edit OOXML or arbitrary user files directly.
- Do not enable hosted Responses tools, arbitrary shell, arbitrary network, multi-agent, background responses, or cloud thread sync in the first release.
- Do not remove the legacy `AgentLoop` until Codex has passed per-editor evaluation and a stable rollback window.

## Architecture

Electron Main owns a pinned `codex app-server` child process and a loopback-only bridge. Codex speaks OpenAI Responses API to that bridge; the bridge translates the validated subset to the Anthropic Messages request/stream shape already accepted by WisUsage and attaches the current WisPaper access token through `AuthClient.fetchWithAuth`. Codex reaches editor capabilities through a document-scoped MCP endpoint; Electron Main routes each validated tool call over IPC to the owning renderer, where the existing `AgentSkill` and document engine execute it.

The renderer never receives model credentials or controls the model, endpoint, auth header, Codex executable, sandbox, or MCP session identity. Codex never receives the WisPaper access or refresh token.

## Global constraints

- Model presented to Codex: `gpt-5.6-sol`.
- Upstream model sent to WisUsage: `openai/gpt-5.6-sol`.
- Upstream endpoint: fixed WisUsage `/v1/messages` endpoint.
- Codex transport: app-server stdio JSON-RPC. WebSocket app-server transport is out of scope.
- Local bridge: random loopback port, per-process Responses credential, per-document MCP credential, constant-time credential comparison.
- Codex environment: allowlist only; no inherited WisPaper token or ambient secrets.
- Codex workspace: an isolated empty directory. Default sandbox is read-only with network disabled.
- Unknown protocol variants fail closed and expose only redacted diagnostics.
- Every document mutation requires optimistic revision validation, an explicit user approval, a pre-mutation snapshot, validation, and an undo entry.
- A runtime failure never silently retries a mutating turn through the legacy runtime.

## Components

### `@wiswork/agent-runtime`

Defines the UI-facing runtime contract and adapts the existing `AgentLoop` as `LegacyAgentRuntime`. This keeps the AI panels independent of a specific harness during staged migration.

### `@wiswork/codex-bridge`

Owns:

- app-server JSON-RPC client and generated protocol bindings for the pinned Codex version;
- child-process lifecycle and redacted stderr diagnostics;
- loopback HTTP server;
- Responses request to Messages request conversion;
- Anthropic Messages SSE to Responses SSE conversion;
- abort, limits, usage, stop-reason, tool-call ID, and opaque reasoning state handling;
- document-scoped MCP server and tool routing.

### Electron Shell integration

Starts the bridge before Codex, creates a minimal Codex configuration, binds Codex threads to document sessions, forwards normalized events to the correct renderer, and tears all resources down on logout, document close, or application quit.

### Renderer integration

Registers each composed `AgentSkill` as a document tool session. Read tools may execute automatically. Mutating tools wait for renderer-owned approval, reject stale document revisions, capture a snapshot, execute through the domain engine, validate, and return a structured result.

## Supported protocol contract

First release supports:

- instructions/system text;
- user/assistant text and image inputs;
- function tool definitions;
- function calls, argument deltas, tool outputs, and parallel calls;
- pinned Codex 0.147 Code Mode custom `exec` only as a constrained carrier for one direct document-scoped WisWork MCP call per turn;
- streaming text;
- request cancellation and upstream disconnect;
- input/output/cache usage when supplied upstream;
- normalized max-token and error termination;
- opaque reasoning/thinking blocks only after the upstream Messages service provides a specified, round-trippable representation. The pinned initial bridge validates Codex reasoning request metadata but rejects reasoning history and upstream `thinking`/`redacted_thinking` blocks; it does not invent a Messages thinking dialect or silently discard reasoning content.

Unsupported fields, tools, and stream events return an explicit compatibility error. Hosted tools, general programmatic tool calling, background responses, server-side conversations, and `previous_response_id` are disabled until separately specified and tested. The sole programmatic exception is the pinned Code Mode `exec` wrapper above: the bridge derives an allowlist from validated `mcp__wiswork` metadata, publishes a generated one-call syntax, and accepts only a direct call to one allowlisted document MCP method with a JSON-object argument and an optional `text(...)` result wrapper.

Codex 0.147 cannot fully configuration-disable `apply_patch` and `view_image`. Task 3 must still disable shell/unified exec, `update_plan`, and multi-agent features, while the protocol translator exposes neither those built-ins nor arbitrary JavaScript. Runtime document-MCP policy remains the critical enforcement boundary: computed properties, multiple calls/statements, shell, filesystem, `apply_patch`, `view_image`, `update_plan`, non-WisWork namespaces, and every unadvertised method fail closed.

## Autonomy and approval

- Automatic: inspect current artifact state, read selections and structure, run non-mutating validation, and use approved WisWork search tools.
- Approval required: document mutations, save/export/print, destructive structural edits, and external actions.
- Refused: arbitrary shell, arbitrary filesystem, direct OOXML mutation, credential access, model/endpoint overrides, cross-document tool access, or permission escalation from model text.

## Failure handling

- Bridge bind/config/start failure: do not start Codex; keep legacy runtime available behind the feature flag.
- Codex crash: terminate the active turn, show a restartable error, and never duplicate a tool action.
- WisUsage 401: refresh once through `fetchWithAuth`, retry once, otherwise return `auth_required`.
- Unknown Responses or Messages event: fail the turn with a redacted compatibility code.
- Renderer/document closed: abort all pending tools and the associated Codex turn.
- Revision mismatch: return `document_changed`; the agent must re-read state before proposing another mutation.
- User denial: return a normal denied tool result so the model can finish without retrying the action.

## Privacy and persistence

Initial Codex threads are ephemeral. Document content and model transcripts are not added to the project file. A future persistent-history feature requires a separate privacy specification, user controls, retention policy, and clear-history operation.

## Rollout and rollback

Roll out behind an authoritative `legacy | codex` runtime flag. Begin with an internal protocol contract test, then LaTeX, Docs, Sheets, Slides, and PDF as separate releases. Rollback changes only the runtime flag; the legacy transport and tools remain intact for at least one stable release after the final editor migration.

Never switch runtimes in the middle of a turn. On Codex failure, preserve the user's prompt and offer an explicit new legacy turn.

## Verification

- Golden and adversarial unit tests for both protocol directions and arbitrary SSE chunk boundaries.
- Contract test using the pinned real Codex app-server against a fake local upstream.
- Auth tests covering missing session, 401 refresh, logout, and token non-disclosure.
- Security tests for non-loopback bind refusal, invalid bridge/MCP tokens, endpoint/model override, cross-document calls, and unknown protocol input.
- Per-editor end-to-end tests for read, approval, mutation, validation, and undo.
- Packaging tests for supported macOS and Windows architectures, binary hash validation, signing, and shutdown cleanup.

## Risk ownership

- Protocol compatibility and pinned Codex upgrades: WisWork desktop runtime owner.
- WisPaper token lifecycle: existing auth package owner.
- Model availability, usage, and billing: existing WisUsage owner; no server change in this project.
- Artifact correctness and undo: each editor/domain-engine owner.
- Rollout flag and incident rollback: WisWork release owner.
