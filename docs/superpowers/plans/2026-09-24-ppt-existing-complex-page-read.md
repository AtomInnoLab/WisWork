# Existing PowerPoint complex-page read

Goal: add a read-only, baseline-scoped view of native table cells and chart cached series before choosing an edit path. The view is bounded, uses exact host slide IDs, and never treats chart caches as verified source data.

Non-goals: edit chart/table data, validate linked workbooks, or change the original solution plan. Real PowerPoint acceptance remains deferred by the user.

Architecture: reuse the existing bounded PPTX ZIP loader and XML parser in the Office add-in. A package parser returns conservative table/chart summaries; the baseline skill exports the exact slide package only after checking the saved baseline and repeats the check afterward. Any unavailable or malformed package fails closed without host writes.

Constraints: no new dependency; 8 MiB package ceiling and 256-entry ZIP ceiling; bounded XML and tool output; no external relationships or path traversal; model-visible content is untrusted data; `qaPassed: false` always.

Files:
- `apps/office-addin/src/skills/powerpoint/powerpoint-package.ts`: expose the existing bounded ZIP reader.
- `apps/office-addin/src/skills/powerpoint/presentation-complex-page-package.ts`: parse bounded table cells and chart cache data.
- `apps/office-addin/tests/presentation-complex-page-package.test.ts`: malformed, external relation, and bounded data checks.
- `apps/office-addin/src/skills/powerpoint/presentation-baseline.ts`: add exact-ID read tool and before/after baseline checks.
- `apps/office-addin/src/agent/host-runtime.ts`: wire existing page export adapter.
- `apps/office-addin/tests/presentation-baseline.test.ts`: tool scope, drift, and no-write behavior.
- Product progress note: report capability and limits.

Task 1: package parser. Write failing tests for one native table and one chart cache, then bounded parser; test unsafe relationship and oversized data. Commit parser unit.

Task 2: baseline tool. Write failing skill tests, add tool and host wiring, test exact ID and baseline drift before/after export. Commit integration unit.

Finish: independent diff review, full Office tests, typecheck, lint, production build; update progress note and retain isolated branch.
