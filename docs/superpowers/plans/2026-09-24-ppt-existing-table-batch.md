# Native table cell batch for existing PowerPoint

Goal: implement a 2–8 cell ordered ChangeSet in one existing native table, with the already established durable cursor, recovery, undo, screenshot review and exact host identity checks. Follows original plan §§6.2–6.4.

1. Produce one table structure digest that excludes the text of every selected target cell and includes all other table content and topology. Reject merged and complex cells, duplicate coordinates and oversized text.
2. Extend the existing batch record and state machine with `table_cell` operations; keep text/geometry batch behavior unchanged. Add a separate table-batch proposal tool that only accepts cells in one table.
3. On each forward, resume or reverse step, recheck the native cell and table package against the persisted batch digest. Save cursor after each verified step; never replay an ambiguous write.
4. Cover partial failure, reopen/resume/undo, structure drift and workbench history with mock-host tests. Independent review and full Office checks. Real PowerPoint acceptance remains deferred.
