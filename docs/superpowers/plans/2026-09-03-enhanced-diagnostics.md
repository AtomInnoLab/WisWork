# Enhanced Runtime Diagnostics Plan

## Goal

Deliver a complete, privacy-preserving diagnostic loop for Enhanced mode: correlate each task,
retain the ten most recent outcomes, expose a bounded self-check, stop deterministic protocol
retry amplification, and let a user copy a diagnostic ID or export a redacted report.

## Non-goals

- Never persist prompts, document contents, model text, tool arguments, credentials, response
  bodies, usernames, or absolute paths.
- Do not upload diagnostics automatically.
- Do not expose main-process logs or arbitrary filesystem access to renderers.

## Architecture and constraints

The Shell main process owns an `EnhancedDiagnosticsStore`. It accepts only closed-enum structured
events, stores a bounded in-memory ring plus an atomic bounded JSON snapshot under userData, and
creates opaque task trace IDs. Codex bridge and runtime codes are normalized at this boundary;
unknown values collapse to `unknown_failure`. A trusted Shell-only IPC exposes summaries, a
read-only self-check, clipboard copy, and save-dialog export. Editor renderers receive only the
current diagnostic ID on failure through the existing runtime error contract.

Global limits: ten tasks, 256 events per task, 64 system events, 512 KiB exported JSON, 30-minute
detailed-diagnostic lease, no raw values. Writes are atomic and fail-open. A deterministic stream
protocol error is terminal and must not be replayed as a timeout.

## Deliverable 1: structured store and safe error correlation

Files:

- Create `apps/shell/src/main/enhanced-diagnostics.ts`: strict schemas, ring buffer, atomic
  persistence, export serialization, task lifecycle and self-check result model.
- Modify `apps/shell/src/main/codex-runtime.ts`: start/end task traces and attach the opaque ID to
  public errors without changing authorization semantics.
- Modify `apps/shell/src/main/codex-engine.ts` and bridge diagnostics wiring only as needed to mark
  deterministic protocol failures terminal.
- Add `apps/shell/tests/enhanced-diagnostics.test.ts` and runtime regressions.

Acceptance: hostile strings and extra fields are rejected/collapsed; secrets never appear; bounds,
restart readback, concurrent task attribution, first-error retention, and protocol fail-fast are
covered RED then GREEN.

## Deliverable 2: trusted IPC, self-check, and export

Files:

- Extend `apps/shell/src/shared/enhanced-mode-api.ts` with exact diagnostic summary/self-check APIs.
- Extend `apps/shell/src/main/enhanced-mode-component.ts` with trusted handlers and strict zero-arg
  validation.
- Extend `apps/shell/src/preload/index.ts` with the narrow API.
- Wire the store, component/auth/runtime probes, clipboard, save dialog, and bounded atomic export
  in `apps/shell/src/main/index.ts`.
- Add focused IPC and self-check tests.

Acceptance: untrusted callers fail closed; self-check reports component/auth/runtime/MCP/WisUsage
as closed states without reading a document; export requires an explicit save destination and
contains only schema-approved data; copying exposes only the opaque ID.

## Deliverable 3: concise diagnostic center UI

Files:

- Extend `apps/shell/src/renderer/src/Home.tsx` with a secondary diagnostic panel, recent tasks,
  self-check, copy-ID and export actions.
- Extend `apps/shell/src/renderer/src/home.css` using semantic design tokens only.
- Add UI/view tests and localization-safe copy (Chinese/English fallback initially follows current
  Shell account-menu convention).

Acceptance: implementation details remain outside the primary account menu; no raw events are
rendered; keyboard and screen-reader semantics are present; success/failure/loading/empty states
are deterministic.

## Verification, migration, rollback, and release

- No migration of old console logs. Missing/corrupt persisted diagnostics starts empty.
- Removing the feature requires deleting the new IPC/UI while leaving existing console diagnostics
  intact; runtime behavior remains fail-open if persistence/export fails.
- Run focused RED/GREEN tests, full Shell and Codex bridge suites, workspace typechecks, Shell
  production build, theme-color gate, lint, Prettier, and diff check.
- Run the real Codex 0.147 integration with data-only WisUsage SSE fixture and verify a protocol
  failure returns promptly with a diagnostic ID rather than timing out.
