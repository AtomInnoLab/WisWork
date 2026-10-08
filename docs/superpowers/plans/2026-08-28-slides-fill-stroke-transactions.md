# Slides fill/stroke canonical transaction migration

## Goal and non-goals

Migrate the Slides Agent fill/stroke family to one durable canonical atomic transaction while preserving the existing schemas, confirmation flow, and legacy semantics. Keep previously migrated text/font and geometry behavior intact. Do not migrate add/delete/table/chart/background operations, and never partially execute a script containing an unenrolled operation family.

## Architecture and constraints

The renderer preflights every requested edit, prepares authoritative durable targets/revision/fingerprint, compiles ordered `set_fill`/`set_stroke` operations (and already-enrolled text/geometry operations when safely mixed), then submits exactly one atomic transaction. Unsupported or non-serializable paint variants fail closed before proposal or mutation. The existing transaction executor owns stale/noop/uncertain/abort/retry/reload behavior and history semantics.

Global constraints: preserve tool schemas/UI/confirmation; preserve none/solid, transparency, color normalization, stroke width/dash, supported element-type and group-child semantics; unsupported image/gradient/pattern/theme paint yields zero mutation; mixed scripts with any legacy-only family remain wholly legacy/fail-closed; one invocation creates at most one history entry.

## Files and responsibilities

- `apps/slides/src/renderer/ai/slides-skill.ts`: tool-family classification, preflight, canonical operation compilation, and transaction dispatch.
- `apps/slides/src/renderer/ai/presentation-style-transactions.ts`: focused fill/stroke preparation and canonical transaction builder, if separation is warranted by existing text/geometry patterns.
- `apps/slides/tests/presentation-text-tool-transaction.test.ts` and/or a focused style transaction test: direct tool and transaction-state coverage.
- `apps/slides/tests/slide-script.test.ts`: pure and safely mixed script routing, ordering, fail-closed behavior, abort/retry/history guarantees.
- Existing desktop-host tests: serializer/persistence and group-child reopen coverage where needed.

## Deliverable

1. Add characterization and missing-behavior tests; demonstrate RED for canonical routing, safe mixing, unsupported zero-mutation, sequential final state, and transaction lifecycle states.
2. Implement the smallest renderer-side style-family compiler and route eligible tools/scripts through one canonical transaction; preserve legacy routing for wholly unenrolled scripts and reject unsafe paint before proposal/mutation.
3. Run targeted tests, surrounding Slides tests, type/lint/build gates declared by the repository, inspect the complete diff, obtain independent review, fix Critical/Important findings within two rounds, and create one scoped commit.

Acceptance: all requested semantics are covered by deterministic tests; fresh full gates pass; diff contains no add/delete migration or schema/confirmation regression.

## Rollback and release

Rollback is the single scoped commit. No persisted-data migration or staged rollout is required: canonical operations use the already-supported desktop host contract. Compatibility is protected by leaving non-enrolled scripts on their existing whole-script path and failing closed when safe persistence cannot be proven.
