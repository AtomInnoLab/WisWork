# Word Write Verification Classification Design

## Purpose

Correct the Word `write_document` failure classification exposed by trace
`b9707256-d33a-454a-af74-1e61eddbbfa4`. A Word commit whose readback differs
from both the original document and the requested semantic result must not be
reported as proven concurrent editing when the adapter cannot distinguish a
coauthor edit from Word's own OOXML or formatting transformation.

## Goals

- Preserve fail-closed document writes and never accept an unverified result.
- Report an unchanged readback as `office_write_failed`.
- Report a changed but semantically invalid readback from the commit itself as
  `office_verify_failed`.
- Reserve `office_concurrent_change` for a document that changes after a
  successful, fingerprinted write, when the adapter has an authoritative known
  post-write state.
- Attach a bounded, content-free verification stage (`text`, `body_shape`,
  `content`, or `boundary`) to local diagnostic events and exported diagnostics.
- Preserve underlying Office error identifiers through the error cause chain.

## Non-goals

- Do not relax the semantic checks for text, headings, tables, inline formatting,
  or append/prepend boundaries.
- Do not automatically retry or roll back a third document state.
- Do not claim that a same-commit third state is definitely caused by Word or by
  a coauthor; that distinction is not observable for whole-document replacement.
- Do not add the new verification-stage field to relay protocol v2. The relay
  continues receiving its existing allowlisted diagnostic shape until a protocol
  version explicitly advertises the field.
- Do not change Excel or PowerPoint transaction behavior.

## Architecture

`BrowserWordAdapter` already computes a detailed semantic verification result.
The adapter will retain that result, wrap failures in a bounded error carrying
only the allowlisted stage, and classify the commit readback from the states it
can prove: unchanged, verified, or changed-but-unverified. The diagnostics layer
will walk the existing shallow cause chain, extract only an allowlisted stage,
and include it in local snapshots/exports while the relay serializer omits it.

## Classification Rules

| Observation | Error | Rationale |
| --- | --- | --- |
| Commit readback equals the original fingerprint | `office_write_failed` | The requested mutation did not commit. |
| Commit readback passes semantic verification | success | Word may normalize irrelevant OOXML, but requested content and structure are present. |
| Commit readback changed but fails semantic verification | `office_verify_failed` | The result is not acceptable, but the actor that produced the third state is not provable. |
| Readback itself fails after a possibly applied commit | `office_state_uncertain` | The document may be partially changed and cannot be inspected safely. |
| A fingerprinted, successfully verified post-write state changes before the separate verification pass | `office_concurrent_change` | A known committed state existed and was subsequently replaced or edited. |

`executeOperations` remains unchanged because its exact text operation model and
existing tests intentionally classify a third full-text state as concurrent.
This change is scoped to native `write_document` writes.

## Diagnostic Data

Add optional `verification_stage` to `OfficeDiagnosticEvent`, restricted to:

- `text`
- `body_shape`
- `content`
- `boundary`

The field contains no document text, paths, fingerprints, prompts, OOXML, or raw
error messages. Unknown, hostile, or invalid values are discarded. Local ring
buffers and exported JSON retain it. `RelaySession.sendDiagnostic` deliberately
does not copy it into the relay v2 frame.

## Files and Ownership

- `apps/office-addin/src/skills/word/browser-word-adapter.ts`
  - Produce staged Word verification errors and correct write classification.
- `apps/office-addin/src/diagnostics/office-diagnostics.ts`
  - Define, sanitize, retain, and export the optional stage.
- `apps/office-addin/tests/word-skill.test.ts`
  - Lock commit-time mismatch and post-success concurrency behavior.
- `apps/office-addin/tests/diagnostics.test.ts`
  - Lock allowlisting, privacy, and hostile-value handling.
- `apps/office-addin/tests/relay-session.test.ts`
  - Prove relay v2 continues omitting the unadvertised field.

The Office add-in owner owns classification behavior and local diagnostics. A
future relay protocol owner must explicitly version and review transmission of
`verification_stage` before it leaves the add-in.

## Failure Handling and Safety

- Every unverified changed state remains an error and prompts document inspection.
- No failure path restores the original whole document over an unknown third state.
- Existing cancellation reconciliation remains unchanged.
- Diagnostic extraction remains exception-safe and bounded to a shallow cause
  chain; diagnostics must never affect document behavior.

## Verification

1. RED tests demonstrate that commit-time semantic mismatch currently reports
   `office_concurrent_change` and lacks staged diagnostics.
2. GREEN tests demonstrate the classification table and allowlisted local field.
3. Word-specific tests and the complete Office add-in suite pass.
4. Office add-in typecheck and lint pass.
5. Final diff receives independent security/concurrency review.

## Rollback

The change is code-only with no migration or persisted state. Rollback is a
single commit revert. Older diagnostic exports remain valid because the field is
optional, and relay v2 frames remain unchanged.
