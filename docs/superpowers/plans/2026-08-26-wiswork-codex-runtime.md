# WisWork Codex Runtime Implementation Plan

## Goal and boundaries

Implement the approved design in `docs/superpowers/specs/2026-08-26-wiswork-codex-runtime-design.md`. Keep WisUsage unchanged, keep WisPaper as the only login, and retain the legacy runtime for rollback. Work is based on commit `ccbc6b5` and branch `codex/wiswork-codex-runtime`.

Known baseline exception: `apps/sheets/native/xlsx-engine` test `archive::tests::extracts_requested_entries_and_reports_manifest` fails before this work because the test and implementation both create the same output directory. The user approved proceeding without changing that unrelated test.

## Task 1: Establish the Responses/Messages protocol contract

**Files:**

- Create `packages/codex-bridge/package.json`, `tsconfig.json`, `vitest.config.ts`.
- Create `packages/codex-bridge/src/types.ts` and `src/index.ts`.
- Create `packages/codex-bridge/tests/responses-to-messages.test.ts`.
- Create `packages/codex-bridge/tests/messages-to-responses.test.ts`.
- Modify root `package.json` test/typecheck scripts.

**Interfaces produced:** validated request/response types and pure conversion functions for the supported subset.

**Sequence:**

1. Write failing tests for text, instructions, model mapping, tools, tool results, invalid model, unsupported fields, and tool ID preservation.
2. Implement the smallest request converter.
3. Write failing tests for text/tool SSE event conversion, usage, stop reason, upstream error, unknown event, and arbitrary chunk boundaries.
4. Implement the smallest streaming converter.
5. Refactor with all tests green.

**Acceptance:** targeted tests and package typecheck pass; unsupported inputs fail closed without including prompts in errors.

**Commit:** `feat(codex): add responses messages protocol bridge`

## Task 2: Add the authenticated loopback Responses server

**Files:**

- Create `packages/codex-bridge/src/local-server.ts`.
- Create `packages/codex-bridge/src/security.ts`.
- Create `packages/codex-bridge/tests/local-server.test.ts`.
- Reuse `AuthClient.fetchWithAuth` through an injected callback; do not import Electron.

**Interfaces produced:** `startResponsesBridge(options)` returning loopback URL, secret, and idempotent `close()`.

**Sequence:** write failing tests for loopback binding, bearer auth, fixed path/model/upstream, streaming, cancellation, 401 callback behavior, size limits, and redacted errors; implement minimally; run surrounding suite.

**Acceptance:** invalid tokens never reach the injected upstream callback; bind address cannot be overridden; abort closes upstream; no token or body appears in errors.

**Commit:** `feat(codex): add authenticated local responses server`

## Task 3: Add pinned app-server protocol and process lifecycle

**Files:**

- Create `packages/codex-bridge/src/json-rpc.ts`, `app-server-client.ts`, and `process-manager.ts`.
- Generate version-matched stable bindings under `packages/codex-bridge/src/generated/`.
- Create process and JSON-RPC tests using fake child streams.
- Add a schema-generation/check script that fails on drift.

**Interfaces produced:** `CodexProcessManager` and `CodexAppServerClient` supporting initialize, thread/start, turn/start, turn/interrupt, notifications, shutdown, and crash reporting.

**Sequence:** test JSON-RPC correlation and malformed messages; test lifecycle, minimal environment, strict config, duplicate start, crash, and idempotent stop; implement; validate bindings against the pinned local binary.

**Acceptance:** no ambient secret inheritance; every pending request rejects on crash; stderr diagnostics are bounded/redacted; no orphan remains after stop.

**Commit:** `feat(codex): host pinned app server runtime`

## Task 4: Add the document-scoped MCP tool bridge

**Files:**

- Create `packages/codex-bridge/src/mcp-server.ts`, `tool-router.ts`, and tests.
- Add tool-session IPC types under `apps/shell/src/shared/codex-api.ts`.
- Add main-process routing under `apps/shell/src/main/codex-ipc.ts`.

**Interfaces consumed:** `AgentSkill`, `AgentToolDef`, `AgentToolCall`, and `ToolExecution` from `@wiswork/agent-core`.

**Interfaces produced:** document-scoped tool registration, list, call, cancellation, and teardown.

**Sequence:** test session/token isolation, schema mapping, read tool execution, mutation approval, revision mismatch, snapshot ordering, denial, cancellation, and closed renderer; implement pure router first, then IPC adapter.

**Acceptance:** one session cannot list/call another session's tools; mutation executes only after approval and snapshot; stale revisions do not mutate.

**Commit:** `feat(codex): expose scoped wiswork tools over mcp`

## Task 5: Introduce the runtime abstraction and Shell orchestration

**Files:**

- Create `packages/agent-runtime/` with runtime/event contracts and legacy adapter.
- Create `apps/shell/src/main/codex-runtime.ts`.
- Modify `apps/shell/src/main/index.ts`, `package.json`, and relevant shared IPC declarations.
- Add unit tests for feature-flag selection, lifecycle ordering, logout, document close, crash, and failover behavior.

**Interfaces produced:** `AgentRuntime`, `AgentSession`, and normalized `AgentEvent` stream.

**Sequence:** characterize current legacy behavior; add runtime interfaces and legacy adapter without UI behavior change; add Codex orchestrator; wire authoritative feature flag; test no silent mid-turn fallback.

**Acceptance:** legacy remains default; Codex cannot start without a logged-in WisPaper session and healthy bridge; logout closes threads, MCP sessions, bridge, and child process.

**Commit:** `feat(agent): add codex runtime orchestration`

## Task 6: Integrate the LaTeX pilot

**Files:**

- Modify `apps/latex/src/renderer/ai/AiPanel.tsx`, `latex-skill.ts`, and proposal review integration.
- Add renderer tool-session adapter and tests.
- Add Shell/LaTeX end-to-end driver coverage.

**Sequence:** preserve legacy UI via characterization tests; register LaTeX skill through the runtime abstraction; map Codex progress/tool events; retain proposal confirmation; add compile/fix/undo workflow.

**Acceptance:** a real Codex turn can read a multi-file project, propose an edit, wait for confirmation, apply through the LaTeX engine, compile, and undo; denial and concurrent edit are safe.

**Commit:** `feat(latex): add codex runtime pilot`

## Task 7: Package, secure, evaluate, and roll out

**Files:**

- Modify `apps/shell/electron-builder.cjs` and third-party notice generation.
- Add binary manifest/hash verification and platform-resolution tests.
- Add protocol compatibility/eval fixtures and release documentation.

**Sequence:** package pinned binaries; verify hashes before spawn; include license notices; run security matrix; run LaTeX eval set against legacy and Codex; document runtime flag and rollback.

**Acceptance:** packaged supported platforms start the pinned binary; tampered/missing binary fails closed; no secret appears in package/log/IPC; Codex meets the approved eval threshold; rollback requires only the runtime flag.

**Commit:** `build(codex): package and gate the runtime pilot`

## Final verification and review

1. Run targeted package tests after each task.
2. Run root typecheck and all applicable tests using a writable Cargo target directory.
3. Record the approved pre-existing Sheets test failure separately; no new failures are permitted.
4. Run formatting and inspect the complete diff for secrets and generated binary artifacts.
5. Dispatch a broad independent security/architecture review.
6. Use `finishing-a-development-branch` and let the user choose merge, PR, or keep-as-is.
