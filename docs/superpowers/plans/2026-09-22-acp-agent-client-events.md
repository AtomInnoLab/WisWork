# ACP Agent–Client Events Implementation Plan

## Goal and non-goals

Adopt ACP v1 as WisWork's Agent-to-frontend event and interaction contract across all clients. Preserve providers, AgentLoop behavior, tools, Office Relay, authentication, document transactions, and deployment topology.

## Architecture and constraints

The shared harness publishes official `@agentclientprotocol/sdk` types. Legacy loop callbacks remain operational until all clients migrate. ACP updates must be bounded, safe for display, use stable IDs, never expose chain-of-thought, and put WisWork-only data under namespaced `_meta`.

## Deliverable 1 — Shared ACP event boundary

Files: `packages/agent-harness/src/acp-events.ts`, `packages/agent-harness/src/harness.ts`, `packages/agent-harness/src/index.ts`, package manifests/lockfile, and focused tests.

- Add the official SDK as a type/runtime dependency.
- Give every harness a stable session ID and an ACP notification subscription.
- Convert cumulative assistant text to true ACP message deltas with stable message IDs.
- Convert tool start/result to `tool_call` and `tool_call_update` with terminal status.
- Prevent stale events after reset/dispose and isolate throwing listeners.
- Keep existing callbacks behavior-compatible.

Acceptance: focused tests prove message deltas, tool lifecycle, session IDs, listener isolation, reset/dispose suppression, and unchanged legacy state behavior. Package typecheck passes.

Commit: `feat(agent): expose ACP session updates`

## Deliverable 2 — ACP prompt and permission semantics

Files: shared harness/core interfaces, Office proposal controller/session, UI reducer, and focused tests.

- Add typed `session/prompt` completion with ACP stop reasons and `session/cancel` behavior.
- Map proposal suspension to `session/request_permission` without bypassing existing validation.
- Settle pending permission requests as `cancelled` on stop/reset/disconnect.
- Add user-safe notices for transport errors; keep diagnostics details bounded and separate.

Acceptance: cancellation never renders as failure; every permission decision is linked to its tool ID; no permission promise or tool row remains pending after terminal paths.

Commit: `feat(agent): adopt ACP prompt and permission lifecycle`

## Deliverable 3 — Office reference client

Files: `apps/office-addin/src/agent/use-office-agent.ts`, presentation timeline/reducer, App components/styles, diagnostics, and tests.

- Replace ad-hoc callback-driven timeline mutations with an ACP reducer.
- Render message chunks, plan entries, tool status, notices, permissions, and prompt completion.
- Preserve current proposal UI and host mutation safeguards.
- Record ACP method/update type in diagnostics without recording document content.

Acceptance: Word/Excel/PowerPoint test fixtures cover streaming, multi-tool turns, permission, cancellation, retry, disconnect, and recovery; existing Office unit/typecheck/build gates pass.

Commit: `feat(office): render agent progress from ACP`

## Deliverable 4 — In-process desktop clients

Files: shared UI reducer/components plus Docs, Sheets, Slides, PDF, Markdown, and LaTeX controllers/panels and tests.

- Reuse the ACP reducer and activity components.
- Migrate one client at a time, preserving app-specific content blocks and rollback hooks.
- Use namespaced `_meta` only for current image/link display data that ACP cannot represent directly.

Acceptance: a shared recorded ACP fixture produces equivalent lifecycle UI across clients; per-app tests, typechecks, and builds pass.

Commit: one scoped commit per client or coherent client group.

## Deliverable 5 — Office Relay payload adoption

Files: Office taskpane relay/session, PC relay client/agent bridge, Rust Relay tests only if framing validation changes, and protocol documentation.

- Keep pairing/auth/outer Relay frames unchanged.
- Carry ACP JSON-RPC messages inside request/chunk bodies after capability negotiation.
- Support mixed-version rollout: new PC + old taskpane and old PC + new taskpane fail closed or use the explicit legacy capability.
- Preserve request IDs, cancellation, timeouts, quotas, and no-payload logging.

Acceptance: protocol tests cover version negotiation, duplicate IDs, cancellation, disconnect, reconnect, bounded payloads, and mixed versions. Deploy PC/Relay before enabling the new taskpane capability.

Commit: `feat(relay): carry ACP agent events`

## Deliverable 6 — Remove legacy event vocabulary

Files: agent-core/harness legacy types, all callers/tests, and documentation.

- Confirm no production consumer uses legacy callbacks.
- Remove compatibility adapters and dead translations.
- Generate a final event inventory and upgrade notes.

Acceptance: repository-wide search finds no legacy lifecycle consumers; full tests, typecheck, lint, format, builds, and real Office acceptance pass.

Commit: `refactor(agent): remove legacy client events`

## Release, rollback, and validation

Release behind a negotiated capability, with Office as canary. Roll back by disabling ACP consumption while retaining the shared adapter. No stored document format changes. Before broad enablement, replay deterministic event recordings and run real Mac/Windows Office checks for multi-document sessions, long tool calls, approval, cancellation, disconnect, and restart recovery.
