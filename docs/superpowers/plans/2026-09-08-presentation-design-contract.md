# Presentation DESIGN.md Contract Implementation Plan

## Goal

Make whole-deck production in WisWork PC and Office Taskpane execute and verify a shared, versioned DESIGN.md contract while remaining compatible with older clients.

## Work units

1. Shared protocol and normalization
   - Add schema types, compatibility normalization, readiness validation, Markdown rendering, revision helpers, and invalidation rules in `packages/agent-core`.
   - Preserve legacy `plan_deck` parsing.
   - Add focused unit tests before implementation.

2. Slides/PC integration
   - Extend `plan_deck` schema and state with the normalized contract.
   - Enforce readiness before production and attach revision/slide/acceptance references to production and screenshot outputs.
   - Persist full DESIGN.md through the existing sidecar.
   - Add workflow/gate tests.

3. Office Taskpane integration
   - Use the same schema and normalization in the PowerPoint skill.
   - Include the active contract in every turn context and expose revision lifecycle in tool output.
   - Keep legacy fields accepted for mixed PC/Taskpane versions.
   - Add PowerPoint skill tests.

4. Timeline and editor
   - Render DESIGN.md creation, lock, revision, and verification inline in both PC and Taskpane timelines.
   - Open the latest revision as editable only while draft/ready; show producing/verified revisions read-only and revise through a new plan call.
   - Add component tests for labels and activation.

5. Verification
   - Run focused agent-core, Slides, Office add-in, bridge, typecheck, lint, and formatting checks.
   - Perform independent specification and code-quality reviews.

## Compatibility constraints

- Do not require new fields from older clients.
- Ignore unknown future fields.
- Preserve existing sidecar files containing plain Markdown.
- Never turn a missing optional capability into `turn_capability_denied`; use declared fallback or report the specific unavailable operation.
