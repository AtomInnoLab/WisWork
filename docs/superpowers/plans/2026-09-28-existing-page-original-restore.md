# Existing page original-package restore

## Goal and boundaries

Enable confirmed, recoverable full-page restoration from single or batch existing-edit savepoints. Reuse the existing staged page replacement transaction so the current edited page is itself backed up before deletion. The restored page receives a new PowerPoint slide ID; original edit records remain historical. Do not claim host visual or save/reopen acceptance from structural checks.

The user's approved PPT Agent plan §6.4 and approval for paired-PC original-page backups govern this work. Keep the backup bytes within the existing PC and Office Taskpane data path, bounded by the existing 8 MiB package limit. Never auto-release an original backup to make room.

## Architecture

The existing-page skill reads an exact source change and page backup from the PC, verifies byte SHA-256 and normalized PPTX content, and writes the package to the Taskpane VFS. Its existing stage and commit tools then create a separate page replacement record: stage preserves the edited page, commit removes it, and undo can restore it. Stage must validate the source link again and persist provenance, so a VFS path alone cannot be presented as a verified original restore.

The PC service raises the per-document active backup count from 8 to 16. This permits eight original-page savepoints plus eight restore-transaction backups in the worst supported batch. The Office backup inventory accepts the same bounded count. Existing metadata and API operations remain unchanged.

## Deliverables

1. **Capacity and ownership.** Update PC quota and Office inventory bound; verify ninth through sixteenth backups and seventeenth refusal, exact release and quota reuse. Files: `apps/shell/src/main/presentation-existing-page-backups.ts`, shell backup tests, Office changes controller and tests.
2. **Restore source preparation.** Add a read-only-host tool that accepts exact source kind/change/page, rejects wrong document, missing/released/in-progress records, reads and verifies PC backup, checks one-slide PPTX, and writes a bounded VFS file. Its output names the existing stage tool and required provenance fields. Files: existing page editing skill and tests; host runtime wiring.
3. **Confirmed transaction provenance.** Add optional immutable restore-source fields to page replacement records. Stage re-verifies the source backup against the VFS package and current source record before proposal and execution; the existing stage/commit/undo state machine remains authoritative for host writes and recovery. Files: existing page record, skill, document binding tests, workbench display, host tool inventory.

## Verification, recovery, and release

For behavior changes, first run tests that fail on the missing behavior, then implement and run focused tests. Run Office and shell full relevant suites, typecheck, format, build, and changed-file lint. An insertion without a durable new ID remains manual review; do not replay it. An interrupted staged transaction resumes through the existing inspection path. Failed PC reads or capacity checks leave the host unchanged. Do not release either original or current-page backup automatically. Real PowerPoint tests remain a separate acceptance gate.

Rollback is reverting the new tool and quota while retaining already-written backup files and page transactions for manual inspection; do not delete user backups during rollback.
