# Durable savepoint for existing chart value edits

Goal: make the bounded existing-chart value update recoverable across taskpane restart and undoable after a confirmed host write. Reuse the existing one-slide package backup service and document change history. Do not expand chart edit scope or claim real PowerPoint acceptance.

The Office skill stores a verified original one-slide package in the local backup service, then writes a document-bound chart change record before replacing the host slide. It records the replacement slide ID and verifies the entire resulting package. Undo requires the exact recorded after-package and restores the backed-up original only through a confirmed proposal. Interrupted writes are inspected and classified; never replay a package insertion without evidence.

Files: `presentation-existing-chart.ts` and document/history binding validate and persist records; `powerpoint-skill.ts` integrates backup, proposal, inspect, resume and undo; `host-runtime.ts` passes the durable binding and backup transport; focused tests cover persisted journal and host outcomes.

1. Add chart record validation and CAS persistence into document history, including write-pending and undo-pending states. Test reopen, invalid transitions, and document identity. Scoped commit.
2. Gate chart edit on durable storage and backup availability; back up original before host write, persist write intent, capture new slide ID, verify post-write. Test no host write on backup/persistence failure. Scoped commit.
3. Add inspect/resume/undo with exact package classification and backup digest checks; interrupt states must not replay unknown writes. Test restart, concurrent drift, and undo. Scoped commit.
4. Independent review and fresh Office tests/typecheck/lint/build. Record the remaining real-host acceptance work and progress.
