# Word Write Verification Classification Implementation Plan

## Goal and Non-goals

Implement the approved design in
`docs/superpowers/specs/2026-08-28-word-write-verification-design.md`: correct
native Word `write_document` commit-time mismatch classification and add a
content-free local verification stage. Do not relax semantic verification,
change automatic recovery, transmit the new field over relay v2, or alter Excel,
PowerPoint, or declarative Word operation behavior.

## Architecture and Global Constraints

The Word adapter remains the authority for semantic write verification. It will
surface the existing detailed verification stage through a bounded error cause;
the diagnostics layer will independently sanitize that optional property for
local snapshots and exports. Relay v2 must continue constructing an explicit
wire allowlist that omits the new field.

Every unknown document state remains fail-closed. No test or implementation may
store document content, OOXML, fingerprints, prompts, stacks, or raw messages in
diagnostic events. All observable behavior changes must be introduced RED-first.

## Files

- `apps/office-addin/src/diagnostics/office-diagnostics.ts`: define and sanitize
  optional local `verification_stage` metadata.
- `apps/office-addin/src/skills/word/browser-word-adapter.ts`: attach the detailed
  stage and classify native commit readback accurately.
- `apps/office-addin/tests/diagnostics.test.ts`: privacy and allowlist coverage.
- `apps/office-addin/tests/relay-session.test.ts`: prove relay v2 omission.
- `apps/office-addin/tests/word-skill.test.ts`: native write state-machine coverage.
- `docs/superpowers/specs/2026-08-28-word-write-verification-design.md`: approved
  design record.
- `docs/superpowers/plans/2026-08-28-word-write-verification.md`: execution record.

## Task 1: Add content-free local verification-stage diagnostics

**Interfaces produced:** optional `OfficeDiagnosticEvent.verification_stage` with
the closed vocabulary `text | body_shape | content | boundary`. Diagnostics may
discover a `verificationStage` property from an error or one of at most two
causes; relay v2 does not transmit it.

**Acceptance criteria:**

- A valid staged error produces `verification_stage` in the local snapshot and
  exported JSON.
- Invalid strings, raw messages, document content, and throwing accessors are
  discarded without changing document behavior.
- The relay v2 frame omits `verification_stage` even when the local event has it.
- Existing Office error code/name/location extraction is unchanged.

**Sequence:**

1. Add diagnostics tests for valid nested stage, invalid/hostile stage, privacy,
   and relay omission.
2. Run
   `npm exec -- vitest run apps/office-addin/tests/diagnostics.test.ts apps/office-addin/tests/relay-session.test.ts`
   and record the expected RED type/assertion failures.
3. Add the optional event field and allowlisted shallow extraction. Keep relay's
   explicit safe-event copy unchanged so the new field is local-only.
4. Re-run the targeted command and obtain GREEN.
5. Run `npm run typecheck -w @wiswork/office-addin`.
6. Commit as `feat(office): add staged Word write diagnostics`.

## Task 2: Correct native Word commit classification

**Interfaces consumed:** the diagnostics convention reads a bounded
`verificationStage` property from the adapter's error cause. No diagnostics
module import is required in the Word adapter.

**Acceptance criteria:**

- A commit readback equal to the original remains `office_write_failed`.
- A changed commit readback failing semantic verification throws
  `office_verify_failed` with the exact content-free stage.
- A successful fingerprinted write that changes before the separate verification
  pass remains `office_concurrent_change`.
- Readback failure remains `office_state_uncertain`.
- No path restores over an unknown third state and no semantic rule is relaxed.

**Sequence:**

1. Change/add Word adapter tests for text mismatch, structure/format mismatch,
   unchanged state, and post-success concurrent change. Inspect the error cause
   only for the bounded stage, never document data.
2. Run `npm exec -- vitest run apps/office-addin/tests/word-skill.test.ts` and
   record the expected RED mismatch (`office_concurrent_change` instead of
   `office_verify_failed`).
3. Reuse `verifyNativeDocumentWriteDetailed` once per final readback, create a
   small staged error cause, and apply the approved classification table.
4. Update existing tests that intentionally encoded the old false classification,
   including inherited-format cases, without weakening their fail-closed result.
5. Re-run the targeted Word test and Task 1 diagnostics tests for GREEN.
6. Commit as `fix(office): classify Word write verification failures`.

## Final Verification and Review

1. Run `npm run test -w @wiswork/office-addin`.
2. Run `npm run typecheck -w @wiswork/office-addin`.
3. Run `npm run build -w @wiswork/office-addin`.
4. Run repository lint scoped by the root command: `npm run lint`.
5. Inspect `git diff origin/main...HEAD`, status, and commit scope.
6. Obtain independent concurrency/privacy review of the complete diff. Fix all
   Critical and Important findings and re-run affected verification.

## Rollback, Migration, Security, and Release

There is no migration or persisted-state change. Roll back by reverting the final
merge or squash commit, or the complete commit set for this change. The optional
local JSON field is backward compatible; relay v2 wire output is byte-shape
compatible apart from unrelated generated identifiers.
Release through the normal Office add-in build after CI. Review must explicitly
check fail-closed behavior, unknown-state rollback avoidance, cause-chain bounds,
and absence of document-derived diagnostic data.
