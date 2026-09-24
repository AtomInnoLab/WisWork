# Existing-deck native picture replacement

## Goal and boundaries

Implement one ordinary native-picture replacement on an existing PowerPoint document, using the approved baseline, durable original-image backup, readback, recovery, and undo. Preserve unrelated objects and existing page identity. Do not route through generated-project identities. This unit excludes grouped/complex pictures, image crops, chart/table internals, whole-page rebuild, and real PowerPoint acceptance (deferred by user).

The existing `BrowserPresentationImageAdapter` and `PresentationImageBackup` perform native picture proof, insertion, recovery and backup. A new existing-document record owns the baseline and state machine; its host tool uses the adapter after a confirmed proposal. A pending or ambiguous insertion is never automatically repeated.

Constraints: work only in `codex/ppt-agent-implementation`, preserve the original plan and root checkout, do not merge/push/deploy, and keep bearer credentials out of the repo. Reuse existing VFS PNG/JPEG validation. Keep existing history and unresolved-write guards, 64-record and 1 MiB limits. Version rollback must not delete history.

## Deliverables

### 1. Durable existing-image record

- Files: new `presentation-existing-image.ts`; `presentation-change-history.ts`; `presentation-document.ts`; storage tests.
- Interface: `PresentationExistingImageChange` with `changeId`, `documentId`, `baselineId/digest`, native scope/slide/shape IDs, asset digest, original `PictureSnapshot`, required `ImageBackupMetadata`, optional inserted/restored IDs and after/undo snapshots, and `pending | complete | undo_pending | undone` state. `readExistingImageChange(changeId)` and CAS `writeExistingImageChange(record, expected)`.
- RED: reject invalid snapshot/scope, impossible transitions, duplicate/unresolved entries, and failed dual-save. GREEN: reopened record, state progress, and rollback work. Reserve space before first native write for recovery evidence.

### 2. Confirmed native tool and recovery

- Files: new `presentation-existing-image-editing.ts`; `host-runtime.ts`; tool tests.
- Tools: replace, list/read/inspect, resume and undo with exact native IDs and current-document checks. Use a fresh baseline and validated VFS image, back up original bytes before savepoint, then native replace. Record inserted ID inside adapter callback; finish by verifying the replacement and saving its final snapshot. Resume only an already identified candidate after adapter proof; never insert twice. Undo loads and checks the original backup, then uses the same adapter with a reverse replacement and records the restored ID. Each write requires proposal confirmation.
- RED: no host write before confirmation, savepoint failure stops insertion, interrupted insert cannot replay, recovery only with matching host proof, undo restores original bytes. GREEN: all above plus document/selection drift rejection.

### 3. History and UI integration

- Files: `presentation-existing-editing.ts`, `presentation-changes.ts`, `presentation-changes-card.tsx`, relevant tests.
- Existing-deck history and offline workbench show picture identity/digests and state, with exact-ID inspect/resume/undo actions. Invalidate QA for the native affected page. A screenshot is available for manual review; no automatic QA-pass claim.

## Verification and release

Run focused model/tool/UI tests first; then Office Add-in full tests, typecheck, lint, format/diff check and production build. Review the complete diff independently for data-loss and recovery risks. Commit scoped, verified units on the isolated branch. No migration is required for old history; new kind is unknown to old binaries and must remain recoverable after upgrade. Do not perform the user-deferred real PowerPoint task acceptance or integrate this branch without a later instruction.
