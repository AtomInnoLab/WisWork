# Existing PowerPoint chart source evidence

Goal: extend the baseline-scoped read-only chart view with bounded evidence about the chart's data source and cache consistency. This follows the original solution plan's §§6.1, 6.3 and factual-chart requirements. Do not expose a chart write operation until source and cache can both be updated and read back.

1. In a separately testable package parser, identify one chart by exact native shape ID, resolve the chart relationship, classify external links, embedded XLSX, or cache-only/unsupported state. Never fetch remote links.
2. For a simple embedded XLSX and one-sheet A1 range formulas, resolve workbook relationships and read bounded string/number cells. Compare exact chart cache points to referenced cells; report `matches`, `mismatch`, or `not_verified` with reasons. Preserve a digest of source bytes, not unbounded workbook content.
3. Add an exact-slide, baseline-scoped tool with pre/post document and package drift checks, using the existing export method. Surface source status as read-only evidence, never as truth/QA/write authorization.
4. Test unsafe relationships, malformed archives, external references, cache mismatch, and drift. Run independent review and Office full tests/typecheck/lint/build. Real PowerPoint acceptance remains deferred.
