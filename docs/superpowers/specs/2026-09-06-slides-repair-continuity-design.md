# Slides repair continuity

Approved direction: continue the root-cause fixes reviewed with the user on 2026-09-06.

## Goal and boundaries

Fix ordinary style scripts rejected before writes, questionnaire completion misclassification and event-order races, and generated-deck image failure cleanup/covered page numbers. Preserve existing document identity, confirmation, scope, cancellation, and strict payload validation. Do not implement unrestricted writes, a new automatic user turn, full-page image backgrounds, or an independent visual scoring model.

## Architecture

The remote agent owns one native tool loop. Host lifecycle tracking must distinguish questionnaire continuation from document write success, and all deferred completion paths use the same admission checks. Renderer tool producers create valid canonical transactions; build_deck reports partial mutation and permits scoped repair when image operations fail.

## Acceptance

- A font-size-only script on ordinary text passes the real canonical transaction adapter; absent optional fields are omitted and explicit false is retained.
- Questionnaire answer followed by actual tool activity cannot later be called questionnaire-incomplete merely because a write failed. A genuine unanswered/uncontinued questionnaire remains distinguishable. Early native completion cannot bypass pending work or cancellation/expiry checks.
- A missing or throwing image insertion preserves partial-write truth, cleans the build gate, and permits later scoped repair; user cancellation still stops. Later successful images are applied after an ordinary missing-image result.
- Cover and image-bearing statement page numbers do not sit underneath their image panel. Statement images are documented as panels, not full-page backgrounds.
- Existing pending-proposal consent, session authorization and strict schema validation remain intact.

## Risk, verification, release

Concurrency behavior is high-risk; add event-order tests and independent review. Mock external/process boundaries only; style regression must traverse the real canonical schema. Run complete Slides, Shell and bridge suites, typechecks, build and inspect changes. No data migration. Keep the isolated branch and do not install over the user's running application. Local artifacts may be packaged after checks; no merge/push without user direction. Rollback is to the previous package/commit without altering document data.
