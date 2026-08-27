# Remove Codex and Enhanced mode implementation plan

## Goal and non-goals

Implement the approved removal design in
`docs/superpowers/specs/2026-08-27-remove-codex-runtime-design.md`. Preserve all Standard agent,
generic LaTeX transaction, Office, Relay, auth, and production configuration behavior.

## Task 1: Add the removal policy and delete standalone runtime packages

**Files:** create `tools/removed-codex.test.mjs`; delete `packages/codex-bridge/**`,
`packages/agent-runtime/**`, `tools/codex/**`, `tools/install-enhanced-component.ts`,
`tools/run-enhanced-mode-evals.mjs`, `tools/codex-release.test.mjs`, and the Enhanced mode release
documentation.

**Acceptance:** the new test is RED against the current tree and GREEN only when removed active
paths are absent. The test permits unrelated historical branch-name text but rejects executable,
package, workflow, UI, IPC, and legal references.

**Sequence:** add and run RED policy test; remove standalone files; update the root test script to
run the policy; run GREEN policy test; scoped commit.

## Task 2: Remove Shell product and process integration

**Files:** `apps/shell/src/main/index.ts`, `tab-manager.ts`, preload, Home UI/CSS, package manifest,
and corresponding tests; delete Codex/Enhanced-specific Shell source and tests.

**Acceptance:** account UI has no Enhanced mode row; preload exposes no related authority; main
does not download/spawn/bridge Codex; ordinary close/quit and Standard editor tabs retain existing
behavior; Shell tests/typecheck/build pass.

**Sequence:** add/adjust UI and lifecycle assertions to fail on current behavior; remove imports,
state, IPC and lifecycle consumers; preserve generic close guards; run focused then full Shell
verification; scoped commit.

## Task 3: Remove LaTeX Codex integration while preserving generic transactions

**Files:** LaTeX main/preload/shared IPC, renderer AiPanel/proposal workflow, package manifest, and
tests; delete Codex tool-session and integration tests. Retain generic proposal/snapshot safety in
`packages/latex-project` unless it becomes provably dead and has equivalent Standard coverage.

**Acceptance:** LaTeX uses only Standard AgentLoop transport; no Codex renderer/window API remains;
proposal review/apply/undo and dirty-state close tests pass; LaTeX and latex-project tests,
typechecks, and builds pass.

**Sequence:** create/adjust Standard-only behavioral assertions; remove runtime branches and IPC;
remove only dead Codex-specific proposal adapters; run focused and full verification; scoped
commit.

## Task 4: Remove release, dependency, and legal integration

**Files:** root/app package manifests and lockfile, `.github/workflows/package-macos.yml`,
`tools/check-licenses.mjs`, `tools/gen-third-party-notices.mjs`, and generated notices if tracked.

**Acceptance:** package workflow retains base unsigned arm64 packaging but has no Enhanced job;
`npm ci` resolves without removed workspaces; license/notices output contains no optional Codex
section; root test/typecheck scripts have no removed workspace; policy test passes.

**Sequence:** edit manifests/workflow/legal tools; regenerate lockfile using npm; run deterministic
YAML/text assertions, licenses and notices; scoped commit.

## Task 5: Full verification, independent review, and release handoff

Run formatting, lint, typecheck, full tests, build, theme/branding/licenses/notices, diff checks, and
the removal policy. Inspect final dependency graph and package/workflow diff. Have an independent
reviewer check for remaining executable authority, Standard-mode regressions, lockfile/legal gaps,
and unsafe user-setting migration. Fix findings with tests, commit, then prepare the branch/PR.

## Security, migration, and rollback

No removed binary is deleted from user machines in this release; it becomes unreachable inert
cache data. Old `agentRuntime` settings are ignored. Roll back only by reverting the removal commit,
and do not reactivate Codex until the original platform-trust gates are satisfiable.
