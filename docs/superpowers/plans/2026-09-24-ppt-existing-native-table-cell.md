# Native existing PowerPoint table cell edit

Goal: implement one confirmed, durable, reversible native table cell text edit under the saved existing-deck baseline. Follow original solution plan §§6.1–6.4. No chart workbook edit and no live-host acceptance in this phase.

1. Add Office.js 1.8 exact-slide/shape table cell read and compare-before-write methods. Reject missing/merged cells, wrong shape type, out-of-range coordinates, and host drift. Verify the target after write.
2. Extend the existing single-change record with `table_cell` and exact row/column, retaining current persisted transition and conflict rules.
3. Add one proposal tool to existing-deck editing. Reuse durable savepoint, undo/resume, post-write screenshot, review, history and workbench actions. Reject truncated/ambiguous baseline or unsupported table shape.
4. Test validator, proposal/confirmation, drift, persisted recovery and undo with Office mocks; update inventory. Review independently and run Office full tests, typecheck, lint and build.

No new dependency or OOXML write path. Preserve plain cell text only; rich cell formatting requires separate evidence and is outside this operation.
