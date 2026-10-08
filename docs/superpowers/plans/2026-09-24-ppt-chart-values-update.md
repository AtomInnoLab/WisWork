# Bounded embedded chart values update

Goal: update numeric values of one existing native PowerPoint chart while keeping its embedded XLSX and chart cache identical. Non-goals: external data links, formulas with calculated cells, category edits, adding/removing series or points, arbitrary OOXML scripts.

Architecture: a package helper resolves the exact native chart shape and supported `Sheet1` cell ranges, requires a verified source/cache baseline, then writes both the chart XML cache and embedded workbook. The Office skill exposes a confirmed proposal with package fingerprint validation, host replacement, package readback and existing rollback path.

Constraints: follow original product plan §§6.3–6.4 and O5; keep source evidence separate from factual QA; never fetch external links; no real PowerPoint acceptance this phase; fail closed on unsupported package layouts or ambiguous mappings.

Files: `presentation-chart-source-package.ts` implements bounded writer; `powerpoint-package.ts` produces binary-aware package edit metadata; `powerpoint-skill.ts` exposes and confirms the operation; corresponding tests cover successful synchronized write, stale source, unsupported chart and host proposal/readback. A progress note records remaining limits.

1. Add a package-level value writer and direct tests. Red: changing values cannot currently be represented; Green: source and cache update together, exact source/cache mismatch and unsupported structures reject. Produce a scoped commit.
2. Add package diff metadata for chart XML and embedded XLSX, and integrate the confirmed host operation. Red: proposal cannot be made/executed; Green: proposal validates the package fingerprint, replacement calls existing transactional adapter, and readback proves both changed parts. Produce a scoped commit.
3. Independently review the complete diff, fix important findings, then run full Office tests, typecheck, lint, build and diff check. Document the boundary and retain rollback via the existing package import recovery path. No deployment or release.
