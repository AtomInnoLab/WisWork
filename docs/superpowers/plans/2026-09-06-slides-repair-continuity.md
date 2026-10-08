# Implementation plan

Base: 5b84b367, linked worktree /private/tmp/wiswork-latest-package, branch codex/fix-enhanced-turn-isolation. Preserve pre-existing image-panel and diagnostic edits and all release artifacts.

1. Style contract: slides-skill.ts merge/canonical fields and presentation-text-tool-transaction.test.ts. Reproduce setStyle(fontSize) through the real geometry adapter, fix the producer, check false/omitted fields. Scoped change: normalize script style contract.
2. Lifecycle: codex-engine.ts and codex-engine-proposal-wait.test.ts. First reproduce partial build and completed-before-settled ordering; separate questionnaire continuation evidence from successful writes and funnel deferred settlement through requestSettle. Include pending questionnaire, failures, expiry and cancellation. Scoped change: unify questionnaire continuation settlement.
3. Images: slides-skill.ts build_deck and build-deck-tool.test.ts. Test missing-then-success image, rejection/cancellation cleanup and actual final-page statement geometry. Preserve partial-write result; keep page number clear of image. Scoped change: preserve repairability after image failure.
4. Integration: retain enhanced-diagnostics.ts terminal-cause fix; independently review all deliverables, resolve Important findings, run complete affected suites/typechecks and builds. Record fresh results. Commit/package only local artifacts; no automatic app replacement or remote integration.

No migration, auth weakening, or new user-visible synthetic prompts. Existing package remains available for rollback. Full native generation/screenshot acceptance requires a real application run; unit/contract tests must not be described as that acceptance.

## Implementation and review evidence

- Style regression reproduced `write_not_applied` through the real transaction adapter/schema before native prepare; omitting undefined optional style properties fixes it while preserving explicit false.
- Lifecycle regressions cover partial writes, native/proposal completion ordering, failed questionnaire followed by successful replacement, pending work, expiry and cancellation. A valid replacement clears the previous questionnaire failure.
- Image regressions cover statement ending pages, footer/image separation, missing/throwing image IPC, cancellation, and native-write-before-null/throw. Independent review found stale renderer reads after an uncertain native write; this is now guarded by document/session/run-checked authoritative refresh before further reads or repairs. A failed refresh remains explicit and blocks stale edits.
- Diagnostic final causes and bounded per-call denial reasons are retained. Independent review found denial phases missing from persisted-state validation; eight failing restart tests now pass and preserve unrelated tasks as well.
- Both scoped independent reviews report no remaining Critical/Important findings after re-review. AiPanel callback token/cancel races were inspected; the new native-refresh tests exercise the skill boundary with host APIs mocked.

## Verification (2026-09-06)

- Fresh complete affected suites: Shell 462 passed / 8 skipped; Slides 723 passed / 1 skipped; codex-bridge 276 passed / 2 skipped. Loopback HTTP tests require execution outside the default network sandbox. `TMPDIR=/private/tmp` avoids the macOS `/var` symlink discrepancy in an existing path assertion.
- Slides, Shell and codex-bridge typechecks pass; changed TypeScript files pass Prettier and `git diff --check`.
- Whole-repository run with `TMPDIR=/private/tmp RUSTUP_TOOLCHAIN=1.88.0 npm test` reaches Office Addin and has one existing failure: `apps/office-addin/tests/workspace-mode.test.ts` expects `autoCorrection: false`, while the unmodified production default is true. That unrelated rollout decision was not changed. The repository is not claimed fully green or merge-ready.
- Local Slides and Shell builds pass. Final package is a local unsigned macOS arm64 regression artifact, not a signed production release; no automatic installation or remote integration.
