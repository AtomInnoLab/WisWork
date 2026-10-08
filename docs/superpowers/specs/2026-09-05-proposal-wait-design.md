# Separate proposal waiting from execution deadlines

## Goal

Repair desktop generation ending in failure while a mutation proposal awaits user confirmation. A generated plan is not an applied document. Preserve explicit consent and fail-closed document identity checks.

## Evidence

Latest diagnostic `diag_BBaca72gf7fg-igQL39vb-qJ` records tool completion at 1788570709455 and host failure at 1788570739456, exactly 30 seconds later. The runtime reports successful model completion in between. This strongly matches expiry but does not prove whether a confirmation dialog was visible.

Two independent deadlines exist: `dynamic-mcp-gateway.ts` sets proposal expiry to 30 seconds; `tool-router.ts` starts a 30-second mutation timeout while the mutation is still queued. The engine also has an idle deadline. Changing only the dialog expiry cannot fix this chain.

## Proposed design

1. Give queued consent a bounded five-minute budget, capped by the existing grant expiry. Keep the actual mutation execution budget at 30 seconds, starting when the trusted authority claims it. Cancel, close, and session invalidation still settle queued operations immediately. Do not increase read-tool budgets.
2. While the model has completed and proposals remain pending, suspend the model idle deadline; proposal expiry still terminates the wait. Never suspend deadlines for genuinely active model work without a separate justification.
3. Keep the confirmation visible until action or expiry. Show expired/not-applied state explicitly instead of silently disappearing. Verify event subscription and rendering in the packaged desktop path before attributing invisibility to a particular cause.
4. Preserve one-shot confirmation, owner/document/generation checks and explicit consent. Do not auto-approve, retry consumed tokens, or declare a pending proposal applied.
5. Distinguish expiry from execution failure in safe diagnostic/UI codes. Successful completion requires the host execution receipt; existing canvas quality checks must still run.

## Verification

- Fake-clock regression: proposal remains pending beyond 30 seconds, can be confirmed before five minutes, applies once.
- Actual execution exceeding 30 seconds still times out.
- Expiry, cancellation, close and stale identity never apply changes.
- Completed model plus pending consent does not fail at the model idle deadline.
- Confirmation component renders, confirms once and shows expiry clearly.
- Run bridge, Shell and relevant renderer suites/typechecks; independent review of timer and consent changes.
- Packaged UI acceptance: submit, wait over 30 seconds, confirm, verify nonempty slides and final receipt. No manual deployment without the user.

## Risk and rollback

This crosses the trusted mutation lifecycle and requires high-assurance review. Longer queued lifetime is bounded and does not grant extra capabilities. Keep modifications in the existing isolated worktree; revert the scoped patch to restore prior behavior. No data migration, Relay or Taskpane deployment is included. The user installs any resulting PC build manually.

## Non-goals

No removal of consent, unlimited capability lifetime, automatic retry of writes, model/provider change, or unrelated generation redesign.
