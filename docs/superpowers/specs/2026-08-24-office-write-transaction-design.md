# Office Write Transaction Design

## Goal

Make Word, Excel, and PowerPoint Agent writes fail safely and report truthfully under Office.js partial commits, delayed read-after-write visibility, cancellation, and concurrent user edits.

## Non-goals

- No new Office.js permissions or ambient code execution.
- No Manifest, Relay, WisWork PC, provider, or model protocol changes.
- No relaxation of proposal confirmation, input limits, or host capability checks.
- No attempt to provide cross-process ACID transactions that Office.js cannot guarantee.

## Contract

Every confirmed write must finish in exactly one provable state:

1. **Applied**: bounded semantic readback proves the approved mutation and its target identity.
2. **Restored**: bounded readback proves the affected state equals the captured pre-write state.
3. **Uncertain**: neither applied nor restored can be proved; report `office_state_uncertain` and explicitly prevent blind automatic retry.

Recovery may run only when the current state matches a state attributable to the current write. A third state is treated as a concurrent edit and must never be overwritten by a whole-document, whole-range, slide, or package restore.

## Shared invariants

- Use stable Office object identity, not collection ordinal alone, whenever the host exposes it.
- Fingerprints include all state needed to interpret the approved operation.
- Volatile package bytes are excluded only by narrow, operation-specific normalization.
- Generic XML/package edits remain byte/path strict outside explicitly normalized regions.
- Read-after-write verification uses a bounded convergence policy: immediate read plus a small fixed number of abortable retries.
- Partial commit handling captures the smallest recoverable target state before dispatch.
- Recovery is compare-and-restore: re-read, classify as pre-state / known write-state / applied / third-state, then restore only a known write-state.
- Diagnostics distinguish validation, write, verification, recovery, concurrency, and uncertain-state failures without serializing document content.

## Host boundaries

### Word

- `write_document` retains semantic Flat OPC verification, adds bounded convergence, includes styles relevant to heading semantics, and never restores over a third state.
- Declarative `execute_office_js` captures affected ranges or a bounded body state, reconciles partial commits, and never reports a failed write as unchanged without proof.
- Verification rejects unexpected formatting on explicitly plain spans while tolerating only documented Word normalization.

### Excel

- Resolve worksheet ordinal inputs to stable Office worksheet IDs at proposal time and revalidate identity at confirmation.
- Capture recoverable state for every direct mutation family; verification failure triggers compare-and-restore only for attributable states.
- Structural and object operations use operation-specific receipts rather than a generic post-write fingerprint alone.
- Cell, style, formula, note, structure, workbook, resize, and object verification use bounded convergence.

### PowerPoint

- Preserve the strict package replacement path, including background-only normalization.
- Direct text, shape, geometry, delete, duplicate, and declarative batches capture semantic slide state and reconcile partial commits.
- Apply the same bounded convergence policy to every direct operation, not only `add_text_box`.
- A post-write deck quality check cannot convert a locally proved write into an unqualified write failure; it reports a distinct quality result.

## Failure codes and UX

- Existing `proposal_stale`, `office_write_failed`, `office_verify_failed`, and `office_recovery_failed` remain stable.
- Add `office_concurrent_change` when a third state is observed.
- Add `office_state_uncertain` when the implementation cannot prove applied or restored.
- Both new outcomes are non-retryable in the current Agent turn. UI copy warns that the document may already contain changes and asks the user to inspect before retrying.
- Relay upload maps new local detail to the existing safe compatible diagnostic vocabulary unless the Relay protocol is deliberately versioned later.

## Verification

- Fault injection for sync rejection after a committed prefix.
- Abort during mutation sync and during delayed readback.
- First and second readbacks stale, later readback converged.
- User edit between execute and verify, and between recovery read and restore.
- Worksheet reorder/insert between proposal and confirmation.
- Unexpected Word formatting, style-definition drift, PowerPoint geometry normalization, and package-only normalization boundaries.
- Full Office Add-in tests, typecheck, lint, format, theme, production build, and independent per-host plus broad review.

## Rollback and release

The implementation is Taskpane-only and can be rolled back by redeploying the previous `apps/office-addin/dist`. Manifest, Relay, and WisWork PC remain compatible. Deploy the Taskpane once all three host units and the broad review pass; then perform real Mac and Windows Word/Excel/PowerPoint partial-failure and delayed-readback acceptance tests.
