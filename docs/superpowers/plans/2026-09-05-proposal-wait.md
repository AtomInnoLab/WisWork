# Proposal wait repair

Approved design: ../specs/2026-09-05-proposal-wait-design.md

Goal: bounded human confirmation separate from execution; no automatic consent or extended execution authority. Existing isolated branch baseline f3f1271e. No migration or deployment.

## Unit 1: trusted lifecycle
Modify packages/codex-bridge/src/tool-router.ts and dynamic-mcp-gateway.ts, apps/shell/src/main/codex-engine.ts and relevant tests. Queued mutations wait up to five minutes (proposal capped by grant), claimed execution remains bounded at 30 seconds. Completed models wait for proposals without idle timeout. Preserve cancellation/close/identity/one-shot safeguards. Add failing clock-based tests, implement, run bridge/Shell suites and typechecks. Scoped commit: fix proposal wait lifecycle.

## Unit 2: confirmation feedback
Inspect packaged confirmation delivery and packages/ui/src/EnhancedMutationConfirmation.tsx plus i18n and tests. Expiry must retain explicit not-applied feedback; no consumed token retry. Test expiry/confirmation/close. Scoped commit: show proposal expiration.

## Review and delivery
Independent complete-diff review, remedy important findings, rerun affected suites and builds. Package PC only if validation passes; user installs manually. Attempt actual UI validation or explicitly report inability. Do not claim document delivery from model prose. Roll back by reverting scoped commits; preserve user data and existing release output.

## Verification record

- Router/gateway/engine regressions failed before implementation, then passed: consent beyond 30 seconds, 5-minute expiry, 30-second claimed execution, grant cap, exact-call claim, deferred completion, cancellation and close.
- UI expiry test failed because prompt vanished; now passes with explicit not-applied notice. Waiting 61 seconds retains confirmation controls.
- Diagnostic expiry test failed with runtime_event; now records proposal_expired. Slides error copy likewise has RED/GREEN evidence.
- Full Shell: 421 passed / 8 opt-in skipped; Bridge: 271 passed / 2 skipped. Shell tests require canonical TMPDIR=/private/tmp on this Mac.
- Slides: 703 passed / 1 skipped. Docs: 799 passed before adding the additional 61-second component test; latest confirmation component suite 8 passed. Agent runtime 21 passed, i18n 23 passed.
- Opt-in installed codex-app-server 0.147.0 with local simulated provider: 10 passed / 1 production-credential test skipped. This is not a real production-model/UI acceptance run.
- Shell, Bridge, Slides and Docs typecheck passed. Shell, Slides, Docs, Sheets and PDF production builds passed (Sheets requires installed Rust 1.88 toolchain).
- Current installed UI inspected read-only; expired proposal absent. New package UI acceptance remains pending user installation; no deployment was performed.
