# Office PowerPoint Enhanced tool contract repair

## Goal and non-goals

Restore the model-driven PowerPoint flow in Enhanced Office sessions so state inspection and interactive clarification are real callable tools, and make state inspection degrade safely on PowerPointApi 1.2/1.4 hosts. Do not hard-code the questionnaire or presentation workflow, expand Office authority, change Relay protocol, or deploy production in this task.

## Architecture and constraints

The Taskpane remains the semantic Office tool executor; WisWork PC validates and exposes those tools through the document-scoped Codex gateway. Both sides must share an exact allowlisted tool contract. PowerPoint state reads require only slide enumeration at API 1.2, while selected-slide enumeration is conditional on API 1.5.

Security constraints remain unchanged: bounded schemas/results, semantic reads only, proposal-gated mutations, no arbitrary JavaScript, and no prompt/document data in diagnostics.

## Files and responsibilities

- packages/codex-bridge/src/tool-router.ts: compile the two PowerPoint workflow reads for the Office host.
- apps/shell/src/main/office-codex-proxy.ts: retain those tools when parsing Taskpane requests.
- apps/office-addin/src/skills/powerpoint/browser-powerpoint-adapter.ts: make selected-slide state conditional on PowerPointApi 1.5.
- Corresponding tests: reproduce catalog omission, proxy filtering, and API 1.2 fallback.

## Deliverable 1: exact Enhanced tool exposure

Add failing tests proving get_presentation_state and ask_clarification survive both the Office proxy and document tool compiler. Add both names as read-only semantic tools in the two exact catalogs. Acceptance: the Enhanced PowerPoint tool list includes both tools and unknown tools remain filtered. Commit independently.

## Deliverable 2: compatible state inspection

Add a failing adapter test for a PowerPointApi 1.2/1.4 host with no getSelectedSlides method. Load the slide collection and return an empty selectedSlideIndexes array; call getSelectedSlides only when API 1.5 is supported. Acceptance: API 1.5 hosts retain selected indices and API 1.2 hosts return bounded state without failure. Commit independently.

## Verification and rollback

Run targeted RED/GREEN tests, full Office Add-in/Codex Bridge/Shell suites, relevant typechecks, production builds, lint/theme/format/diff checks, and inspect the final PR diff. Rollback is reverting the two scoped commits; no data migration or protocol rollout is required. Production deployment remains a separate explicit action.
