# Remove presentation asset planning quotas

Goal: remove only the approved 120/20/4 planning count quotas, preserving safety checks and existing layout behavior. Approved design: `../specs/2026-09-11-presentation-asset-counts-design.md`. Worktree: existing `codex/fix-office-production-recovery`, base `5ce0950`; preserve the three earlier local repair commits. No version bump, PR update or deployment.

Architecture: PC and Office share the modern contract schema and runtime parser from agent-core, but expose separate legacy plan schemas. Relax item counts at the shared source and align the legacy tool definitions; retain per-item validation and all existing request/resource envelopes. Inspection confirmed the PC legacy schema was already uncapped, so only the Office legacy schema needs a production edit.

## Deliverables

1. Shared planning contract, parser and tests (`packages/agent-core/src/presentation-design-workflow.ts`, `packages/agent-core/tests/presentation-design-workflow.test.ts`). RED for schema quotas and parser rejection beyond20 references/fourqueries; GREEN preserving all entries and modern/legacy roundtrips. Include more than120 assets and negative malformed/oversized-string/unknown-reference cases. Use existing helpers with an explicit uncapped-count option; no new generic validator framework. Proposed scoped commit: `fix(slides): remove fixed asset planning quotas`.
2. Consumer schema alignment and regression tests (`apps/office-addin/src/skills/powerpoint/powerpoint-skill.ts`, `apps/office-addin/tests/powerpoint-skill.test.ts`, `apps/slides/tests/presentation-design-plan.test.ts`). RED Office's exposed legacy schema still caps image_queries; remove only that maxItems, verify modern schemas inherit shared relaxation and both real planning paths accept a larger query list. PC's production tool definition remains unchanged. Preserve build_deck scalar imageUrl and all unrelated limits. Review with the shared unit before the same scoped integration commit.
3. Final verification and independent review. Run full `npm test`, `npm run typecheck`, lint, formatting, PC Slides/Taskpane/Shell builds and whitespace checks. Review the entire diff against the approved design. Keep local commit/worktree; no external changes. Record exact outcomes and any platform-verification limits here.

## Rollback and release

No migration. Revert the scoped commit to restore former model-facing and runtime quotas together. Future activation requires rebuilt PC and Taskpane; Relay is unaffected by this change. Larger plans remain constrained by unchanged byte/memory/verification envelopes and should not be advertised as unlimited deck capacity.

## Execution evidence

- Shared workflow baseline: 13 tests passed. New tests first produced three intended failures: the inventory schema still contained maxItems:120, 25 legacy image queries were rejected, and 21 modern asset references were rejected. After the minimal shared change, all 178 agent-core tests passed.
- Office schema regression first failed specifically on maxItems:4. Both hosts now preserve five legacy queries through real plan_deck execution. A PC regression also preserves 21 modern assets through saving and restoring DESIGN.md.
- Final independent review found no actionable issues. Reviewer independently ran shared workflow 14/14, PC planning/restoration 21/21, and PowerPoint skill 146/146; formatting and whitespace checks passed.
- Consumer implementer verified full Slides suite (769 passed, six skipped), full Office suite (1,027 passed), and both host typechecks.
- Full `npm test` exited 0: 7,735 Vitest tests passed and 21 skipped; native Rust tests also passed (52 unit tests and one integration test). Office Add-in, Slides and Shell production builds completed successfully. Taskpane retains the existing large-chunk warning.
- Full lint completed with zero errors and 13 existing warnings in unchanged files. Full format check, final seven-file Prettier check and `git diff --check` passed.
- Full `npm run typecheck` exited 0 across all workspaces.
- No actual Mac PowerPoint session was exercised: tests cover the changed planning schemas, parsers, persistence and host tool wiring, not host rendering or unrestricted deck capacity. No release, deployment or remote PR mutation is included.
