# Office Agent confirmation-loop design

## Goal

Make Word, Excel, and PowerPoint write tools reliably execute through one continuous Agent turn: the tool proposes a bounded change, the Agent pauses, the user can immediately approve or reject, and the same turn resumes with the actual confirmed outcome. Keep the composer visible at the bottom and replace raw JSON approval UI with a host-aware human-readable review.

## Non-goals

- Do not change the WisWork PC Agent harness, identity, Relay protocol, manifest, or deployed URLs.
- Do not allow writes without explicit confirmation or weaken stale-state, cancellation, semantic verification, or recovery checks.
- Do not expose arbitrary Office.js or raw tool JSON in the default UI.

## Architecture

`@wiswork/agent-core` gains an explicit confirmation suspension contract. A tool execution may pause after a proposal is published; the loop keeps its generation and tool-call position but performs no additional provider request until the Office session resolves the confirmation. Approval executes the existing proposal controller, then resumes the same loop with a success tool result; rejection resumes it with a bounded rejection result. Cancellation, logout, new task, or auth loss abort both the pending confirmation and the run.

The Office presentation layer subscribes to proposal lifecycle changes so the confirmation card appears before the tool promise resolves. It renders a host-aware summary and compact before/after or target list; structured objects are converted to bounded labels and rows rather than raw JSON. The workspace remains a fixed-height grid with only the timeline scrolling and the composer pinned to the bottom.

## Global constraints

- Exactly one active confirmation per Office Agent session.
- No provider call may occur between proposal creation and its user decision.
- Confirm and Reject are usable while the loop is suspended, but conflicting composer/session/file/skill actions remain disabled.
- A proposal is executed at most once and only by the existing confirmation controller.
- Logout, disconnect, Stop, New task, and generation changes settle the suspended tool without a late write or stale UI.
- Tool results sent back to the model contain stable bounded outcome codes, never raw Office errors or proposal payloads.
- The default approval card contains no JSON/code block; an optional bounded text diff is allowed only for scalar document text.
- The composer always occupies the last grid row and never scrolls with the timeline.
- Rollback uses the existing `VITE_WISWORK_OFFICE_WORKSPACE=0` legacy workspace flag and a new exact confirmation-loop flag if core protocol rollback is required.

## Components

- `packages/agent-core/src/{types,loop}.ts`: suspension/result-resume contract and lifecycle handling.
- `packages/agent-core/tests/loop.test.ts`: no provider continuation before decision; approve/reject/cancel resume semantics.
- `apps/office-addin/src/agent/proposal-controller.ts`: observable decision lifecycle without expanding write authority.
- `apps/office-addin/src/agent/use-office-agent.ts`: bridge proposals to suspended AgentLoop and expose confirming state.
- `apps/office-addin/src/App.tsx`: host-aware approval summary, actionable buttons during suspension, fixed composer behavior.
- `apps/office-addin/src/styles.css`: independent timeline scrolling and pinned composer layout.
- `apps/office-addin/tests/*`: Word/Excel/PowerPoint write-path integration, readable approval rendering, interaction and layout regressions.

## Failure handling

- Proposal creation failure returns a normal bounded tool error and does not suspend.
- Approval verification/recovery failures resume the tool with their stable code and render safe actionable copy.
- Reject returns `user_rejected_change`; the model may acknowledge but must not retry another write in the same turn.
- Stop/logout/disconnect/new task resolves the suspension as cancelled and prevents any later continuation from mutating UI or Office.
- A malformed or duplicate suspension fails closed as `invalid_tool_output`.

## Verification and release

- RED/GREEN tests at AgentLoop, Office session, UI, and all three host write-skill boundaries.
- Full agent-core and Office test suites, typechecks, configured production build, lint, format, theme, and diff checks.
- Independent review of core suspension semantics and Office integration, followed by broad final review.
- Manual acceptance in desktop Word, Excel, and PowerPoint: generate a write, approve before assistant completion, observe the document mutation and resumed response; repeat reject, Stop, logout, narrow pane, and long preview cases.

