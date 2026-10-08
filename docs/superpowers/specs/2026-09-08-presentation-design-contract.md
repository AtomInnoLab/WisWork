# Presentation DESIGN.md Production Contract

## Decision

Whole-deck creation and redesign are driven by a versioned DESIGN.md contract. The contract is created before slide mutation, locked for production, referenced by every production and verification action, and revised explicitly when the plan changes.

## Contract

The user-readable Markdown and the structured contract are two views of the same data. The contract contains:

- brief: topic, audience, occasion, desired outcome, language, page count, aspect ratio, and source constraints;
- narrative: hook, opening, development, tension, resolution, and closing action;
- visual system and anti-patterns;
- a director plan for every slide: role, claim, content, evidence, visual route, layout family, focal visual, density, assets, and stable acceptance IDs;
- an asset inventory with provenance, placement intent, validation status, local reference, and fallback;
- deck-level acceptance criteria.

Assets move through `needed -> searching -> downloaded -> validated -> ready`. An asset may use `fallback_ready` only when its fallback is explicit and validated.

## Lifecycle

`draft -> ready -> producing -> verified`.

Production may start only when the contract is complete, all required assets are `ready` or `fallback_ready`, every slide has acceptance criteria, and prototype pages are declared. Starting production locks the current revision.

If a plan becomes infeasible, create a new revision instead of silently changing the implementation. Slide-local changes invalidate that slide's checks; visual-system changes invalidate deck consistency; narrative or page-count changes invalidate all affected slide plans.

## Workflow

1. Understand the request and inspect the current deck and attachments.
2. Ask only material missing questions.
3. Research facts and visual assets.
4. Build the story, visual system, slide director plan, asset plan, and acceptance criteria.
5. Validate readiness and create/lock DESIGN.md revision 1.
6. Produce representative prototype pages.
7. Verify prototypes against their acceptance IDs, repair, and re-screenshot.
8. Produce remaining slides in bounded batches, repeating screenshot and repair.
9. Verify slide and deck acceptance, then mark the revision verified.

## Compatibility

The shared protocol uses `schemaVersion: 1`. New fields are optional at transport boundaries so newer PC and Taskpane versions can read older plans. Unknown fields are ignored. Legacy `core_hook`, `style`, `pages`, and `prototype_pages` inputs are normalized to the structured contract rather than rejected.

PC persists the Markdown sidecar next to the presentation. Taskpane receives and renders the same contract snapshot through the PC bridge. Timeline events show creation, locking, revision, and verification at the point they occur and open the editable/read-only Markdown view.

Execution logs and tool errors remain in the timeline and are not copied into DESIGN.md.
