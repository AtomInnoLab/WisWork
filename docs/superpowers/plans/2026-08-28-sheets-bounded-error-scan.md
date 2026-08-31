# Sheets bounded workbook error scan

## Goal and non-goals

Implement a complete, bounded, cancellable error-value scan for streamed workbooks and safe navigation to findings. The scan covers visible and hidden worksheets and both file-side and journal cells; filtered/manual-hidden rows remain in scope because error checking diagnoses workbook data, not merely rendered rows. It does not expose formula text, mutate cells, or claim completeness after any budget, timeout, read, identity, or session failure.

## Architecture and constraints

A pure scanner pages worksheets in stable workbook order through `readSheetRangeMapped`, merges journal entries as authoritative shadows, and emits only `{sheetId,address,errorCode}` findings plus progress and a terminal status. Every batch validates the captured lazy-state object, session, generation, workbook id, worksheet identity, sheet/filter/journal snapshots, and `AbortSignal`; cells, findings, batches, accepted bytes, and deadline are independently capped. The ribbon adapter owns one scan session, exposes progress/cancel through repeated clicks, renders localized safe summaries, and navigates only after revalidating the current workbook and target worksheet.

Global constraints: exact spreadsheet error tokens only; no formula or raw exception text; deterministic ordering; fail closed on drift; zero findings is “clean” only for a complete scan; late async results are inert.

## Files

- `apps/sheets/src/renderer/error-checking.ts`: scanner, budgets/statuses, ordering, cancellation, identity/drift validation, safe runner/navigation.
- `apps/sheets/tests/error-checking.test.ts`: RED/GREEN behavior coverage for unloaded/large/multi-sheet/hidden/error tokens/journal/budgets/timeout/identity/drift/clean completeness.
- `apps/sheets/src/renderer/ribbon-actions.ts`: dispatch error-checking to the streamed scanner.
- `apps/sheets/src/renderer/ExcelShell.tsx`: make Error Checking an accessible actionable ribbon button.
- `apps/sheets/src/renderer/i18n/strings-app.ts`: localized progress/cancel/complete/incomplete messages without raw errors.

## Deliverable

1. Add scanner tests and confirm they fail because the module/API is absent.
2. Implement bounded scan and identity-safe runner until targeted tests pass.
3. Wire ribbon/UI/i18n, then run surrounding Sheets tests and typecheck.
4. Run Sheets non-LO full tests, build, lint, format check, inspect scoped diff, and create one scoped commit.

Rollback is the single scoped commit; there is no persistence migration or file-format change. Release risk is limited to the new read-only command and is controlled by the explicit incomplete/unavailable states and identity fail-closed behavior.
