# Slides Capability Truth Implementation Plan

## Goal and non-goals

Make supported Slides editing capabilities executable in practice by preventing one false, tool-free capability denial per run. Do not add new mutation powers, infer targets, or change transaction and permission boundaries.

## Architecture

Agent Core provides an optional one-shot terminal-response review hook. Slides supplies a pure denial policy and an explicit capability map. Rejected prose is retained only in model history for correction; the current UI bubble continues into the corrective turn.

## Global constraints

- Maximum one completion review retry per run.
- No review on cancellation or turn-limit finalization.
- Static correction text only; no document content or raw errors.
- Existing tools and receipts remain the sole mutation authority.
- RED must precede GREEN for each behavior change.

## Task 1: Optional one-shot Agent Core completion review

Files:

- `packages/agent-core/src/skill.ts`: define the bounded review context and optional hook; compose child hooks.
- `packages/agent-core/src/loop.ts`: invoke the hook before normal tool-free completion and schedule at most one correction turn.
- `packages/agent-core/tests/agent-loop.test.ts` (or closest existing loop suite): behavioral coverage.

Acceptance:

- A rejected final response does not call `onDone` and produces one corrective request.
- The correction stays in the same run and the hook cannot trigger twice.
- Cancelled/finalizing runs and hook exceptions finish safely.
- Skills without the hook are byte-for-behavior unchanged.

Verification:

- Confirm focused tests fail because no hook exists.
- Implement the smallest interface and loop change.
- Run Agent Core tests and typecheck.
- Scoped commit.

## Task 2: Slides capability map and denial policy

Files:

- `apps/slides/src/renderer/ai/slides-skill.ts`: explicit editing contract, pure denial classifier, hook wiring.
- `apps/slides/tests/slides-capability-truth.test.ts`: prompt and policy regressions using the screenshot wording and safe negative cases.

Acceptance:

- Chinese and English false denials about text color or title/element position request one correction when no mutation occurred.
- Normal questions, real unsupported capabilities, and responses after a mutation are not rejected.
- Prompt gives an actionable multi-page recipe and forbids advice-only completion for edit requests.
- No tool schema, authority, or transaction code changes.

Verification:

- Confirm tests fail on the current prompt/skill.
- Implement policy and wire it through `createSlidesSkill`.
- Run focused tests, full Slides tests, Slides typecheck, and production build.
- Scoped commit.

## Task 3: Independent review and release verification

- Independently review both commits for loop safety, history validity, false positives, privacy, and transaction-boundary preservation.
- Resolve all Critical and Important findings within two fix rounds.
- Run format, lint, Agent Core and Slides typechecks/tests, Slides production build, and `git diff --check`.
- Push the additional commits to PR #74; no force push.
