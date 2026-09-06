# Enhanced Live Failure Diagnostics Plan

## Goal and non-goals

Goal: identify the exact local runtime boundary responsible for the live Enhanced-mode failure, while keeping exported diagnostics free of prompts, document content, paths, credentials, and model output. Once one reproducible failing boundary is proven, fix that root cause with a regression test.

Non-goals: do not change Slides document generation, proposal semantics, confirmation UX, Relay, Taskpane, or retry policy until evidence identifies one of those components as the cause.

## Architecture

The shell owns task correlation and persists a closed set of safe lifecycle codes. The Codex engine will emit bounded milestones around thread creation and turn creation; existing app-server error notifications will map to distinct safe categories instead of collapsing into `unknown_failure`. These events allow a live report to show whether failure precedes upstream streaming, proposal creation, or document execution.

## Global constraints

- Never persist prompts, model output, tool arguments/results, document content, filesystem paths, tokens, raw error messages, or arbitrary runtime strings.
- Diagnostic codes must come from a closed enum and pass persisted-state validation.
- Instrumentation must not change control flow, retry behavior, or user-visible completion semantics.
- Preserve existing untracked release artifacts.

## Files and responsibilities

- `apps/shell/src/main/codex-engine.ts`: emit safe start-thread/start-turn and proposal lifecycle milestones.
- `apps/shell/src/main/enhanced-diagnostics.ts`: map raw internal milestones to closed persisted codes.
- `apps/shell/tests/enhanced-diagnostics.test.ts`: prove category separation and redaction.
- `apps/shell/tests/codex-engine.integration.test.ts` or a focused engine test: prove event ordering at each boundary.

## Deliverable 1: boundary diagnostics

Acceptance criteria:

- A thread-start rejection records `thread_start_failed` and no upstream milestone.
- A turn-start rejection records `turn_start_failed` after successful thread creation.
- App-server error notification, Codex error notification, and system-error thread status remain distinguishable.
- A proposal records created, settled-success, settled-cancelled, expired, or execution-failed without proposal contents.
- Existing reports remain parseable and bounded.

Implementation sequence:

1. Add failing tests for exact safe mappings and engine boundary ordering.
2. Run the focused tests and capture the intended RED failures.
3. Add only closed-enum mappings and lifecycle emissions.
4. Run focused shell tests and then the full shell suite and typecheck.
5. Commit the diagnostic-only change.

## Deliverable 2: live evidence and root-cause fix

Acceptance criteria:

- A diagnostic build reproduces the minimal one-slide request and identifies one exact failure stage.
- The identified failure has a deterministic regression test that fails before the behavioral fix.
- The smallest root-cause fix passes the regression and surrounding suites.
- A rebuilt local package passes the same live scenario through proposal application and canvas verification.

Implementation sequence:

1. Package the diagnostic build for local dogfood; the user installs it because `/Applications` is outside the workspace.
2. Reproduce once and inspect the exported report.
3. Trace the failing stage backward to its source and compare with the known-working implementation/configuration.
4. Add the failing regression test, implement one fix, and verify.
5. Package the fixed build and repeat the live one-slide scenario.

## Rollback, migration, security, release

There is no data migration. Reverting the diagnostic commit restores the previous report vocabulary. The package remains unsigned/unnotarized dogfood and does not update Relay or Taskpane. Before release, verify diagnostics contain only closed codes and bounded metadata, then run the standard signed/notarized release pipeline separately.
