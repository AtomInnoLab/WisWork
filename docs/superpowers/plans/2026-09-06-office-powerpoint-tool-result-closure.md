# Office PowerPoint Tool Result Closure

## Goal

Restore reliable model-driven PowerPoint creation in the Office Taskpane by making the real Office.js tool outcome observable end to end, giving the model a deterministic presentation-state probe, and proving the read/create/edit/screenshot/verify loop without relying on a packaged desktop build.

## Non-goals

- Do not replace WisWork's Agent Runtime, UI, pairing, or authorization model with `office-agents`.
- Do not hard-code questionnaire, planning, research, or authoring decisions in the host.
- Do not expose arbitrary JavaScript beyond the existing bounded declarative tool contract.
- Do not weaken mutation verification or recovery boundaries.

## Architecture

The Taskpane remains the semantic PowerPoint tool executor and returns one bounded result envelope for every remote call. The PC bridge and dynamic MCP gateway wait for and classify that semantic result instead of treating dispatch as success. A small `get_presentation_state` read tool provides slide count, selection, dimensions, and supported API sets so the model can choose valid zero-based targets before editing.

## Global constraints

- Preserve stable public error codes and redact document content from diagnostics.
- All slide indices are zero-based; schemas must say that explicitly.
- A failed semantic tool result must never be aggregated as a successful MCP tool execution.
- Existing Office proposal/verification and cancellation behavior remains authoritative.
- UI colors continue to use shared semantic theme tokens.

## Files and responsibilities

- `apps/office-addin/src/skills/powerpoint/browser-powerpoint-adapter.ts`: bounded presentation-state read and compatible collection loading.
- `apps/office-addin/src/skills/powerpoint/powerpoint-skill.ts`: state tool, explicit index contracts, structured safe failures.
- `apps/office-addin/src/agent/use-office-agent.ts`: record remote semantic tool success/failure and preserve safe error metadata.
- `apps/office-addin/src/pc-bridge/session.ts`: carry the semantic result envelope without collapsing it to transport completion.
- `apps/shell/src/main/dynamic-mcp-gateway.ts` and diagnostics integration: classify semantic completion separately from dispatch.
- Office and shell tests: regression coverage and an in-process PowerPoint creation workflow.
- `tools/` and package scripts if needed: a developer command that directly invokes a registered Taskpane tool through the existing bridge.

## Deliverable 1: truthful remote tool outcomes

Acceptance: a remote Taskpane tool returning `office_read_failed` produces a failed diagnostic event with its tool name and safe error code; successful output is marked successful only after the result response is accepted.

Sequence: add failing bridge/session and diagnostics tests; extend the bounded result/event contract; implement propagation; run targeted Office and shell tests; commit independently.

## Deliverable 2: deterministic PowerPoint state and tool contracts

Acceptance: `get_presentation_state` returns slide count, selected zero-based indices, dimensions, and supported API flags; every slide-index schema explains zero-based numbering; a one-slide blank deck can be read using index 0.

Sequence: add failing adapter/skill tests; implement the bounded state read using loaded collections; update prompt and schemas; run PowerPoint suites; commit independently.

## Deliverable 3: executable creation-loop regression

Acceptance: an in-process enhanced-mode test performs state read, slide mutation, screenshot, and verification and asserts semantic failures cannot be reported as success. A developer command can inspect and invoke a live registered PowerPoint Taskpane session without invoking the model.

Sequence: add failing workflow and command-contract tests; implement the smallest bridge-facing harness; run Office, shell, and agent-core suites; commit independently.

## Verification

- Targeted RED/GREEN evidence for each changed behavior.
- Office Add-in unit and integration suites.
- Shell Codex/MCP/diagnostics integration suites.
- Agent-core loop suite.
- Typecheck, lint, theme-color check, production Taskpane build, and manifest validation.
- Confirm generated diagnostics contain no prompt, document content, secrets, or local paths.

## Rollback and release

Each deliverable is a scoped commit and can be reverted independently. The new result fields are additive and bounded; old clients continue to use `output` and `is_error`. After verification, bump Taskpane cache/version only if deployment is requested; desktop version changes remain a separate release decision.

## Security and migration

Only allowlisted error codes, tool identifiers, timestamps, and durations cross the diagnostic boundary. No persistent data migration is required. Unknown or malformed result envelopes fail closed as `tool_execution_failed`.
