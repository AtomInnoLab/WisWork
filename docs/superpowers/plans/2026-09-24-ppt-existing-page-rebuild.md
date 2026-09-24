# Existing-deck single-page rebuild

## Goal and boundaries

Implement an existing PowerPoint page rebuild from an explicitly prepared one-page PPTX package, with a fresh native baseline, before/after difference, durable original-page backup, explicit confirmation, readback, crash recovery and undo. Preserve unrelated pages. This follows the approved product plan §6.1–6.4 and O5. The real PowerPoint task acceptance remains deferred by the user.

Do not reuse generated-project identity or imported-page receipts as persisted existing-deck identity. Do not accept a raster image as the editable replacement. Do not silently replay an uncertain insertion. Continue on the isolated `codex/ppt-agent-implementation` branch; no merge, push or deployment.

## Design

Reuse the browser native page replacement adapter's order/content checks and stage/commit/undo protocol. Bridge the existing-page record to the adapter's structural input in memory only. A new PC backup route saves the exported original PPTX package under document identity without requiring production lineage. A new Office settings record owns the exact native scope, package digests and replacement state; its history entry and unresolved-change guard prevent conflicting writes. The tool validates a bounded one-page VFS PPTX, presents a difference proposal, and performs backup → savepoint → stage → commit. Resume/undo are only allowed when host evidence matches the saved IDs and digests.

## Deliverables

### 1. Existing-page immutable backup

- Files: `apps/shell/src/main` backup service/router and focused tests.
- Interface: document-bound `existing_page_backup_begin/chunk/finish/status/read` operations carrying backup ID, native slide ID/order, byte digest and size. Ready bytes are immutable and reverified on read.
- RED/GREEN: reject another document, wrong bytes, duplicate mutable metadata, excess size and invalid PPTX; persist and reopen a valid original one-page package. Commit scoped.

### 2. Durable replacement record

- Files: `presentation-existing-page.ts`, `presentation-document.ts`, `presentation-change-history.ts`, focused tests.
- Interface: `readExistingPageChange(changeId)` and CAS `writeExistingPageChange(record, expected)` with pending/staged/applied/undo states, original/replacement package digests, native slide IDs/order and backup identity.
- RED/GREEN: reject illegal transitions, stale writes, overlapping unresolved changes and capacity overflow; persist/read/recover after reopen with rollback on failed save. Commit scoped.

### 3. Confirmed host tool and workbench

- Files: new existing-page editing skill, native runtime/tool dispatch, history/workbench and focused tests.
- Tool reads fresh baseline and a bounded one-page VFS PPTX, verifies target and source page package, surfaces a concise before/after scope and high-risk full-page replacement proposal. Confirmed execution first backs up original page and persists `pending`; adapter stages a known replacement ID, then requires a separate confirmation to commit deletion of the original. Recovery never repeats an unknown insertion. Undo restores backed-up original before deleting the revision. Expose inspect/resume/undo in history. Invalidate affected-page QA; no claim of visual QA pass.
- RED/GREEN: no host write before confirmation; failed backup/savepoint blocks insertion; interrupted stage never repeats; fresh host proof gates commit/undo; unrelated pages remain intact; reopened workbench exposes recovery. Commit scoped.

## Verification and release

Run focused tests, Office and Shell full relevant suites, TypeScript checks, ESLint, production builds and `git diff --check`. Independently review backup/storage and final host flow for data loss. Record exact progress and gaps. Preserve the original plan and deferred 0/20 acceptance. No migration of older generated replacement records; old history remains readable.
