# Slides Capability Truth Design

## Goal

Prevent the Slides agent from falsely claiming that supported text-style or geometry edits are unavailable. A request such as “make pages 6–8 use the same title color and position” must attempt the existing safe write tools instead of stopping after inspection with a capability denial.

## Current failure

Slides already exposes:

- `set_element_style` for font size, color, bold, italic, underline, alignment, and font family;
- `set_element_transform` for position, size, and rotation;
- `execute_slide_script` for coordinated text, style, paint, and geometry edits on one page.

The current system prompt describes these tools but also permits the model to report an unavailable capability without first reconciling that claim with the advertised tool contract. Agent Core accepts the resulting tool-free prose as a successful terminal response.

## Design

### Explicit Slides execution contract

The Slides system prompt will contain a compact capability map and mandatory editing rules:

- use `set_element_style` or `execute_slide_script.setStyle()` for text color and typography;
- use `set_element_transform` or script geometry primitives for title position and size;
- execute multi-page consistency work once per target page;
- do not stop at inspection or advice when the user requested an edit;
- report a limitation only after a relevant tool attempt returns an unsupported or fail-closed result.

This changes model guidance only; it does not expand tool permissions.

### One-shot terminal-response review

`AgentSkill` gains an optional pure `reviewFinalResponse` hook. Agent Core invokes it only for a normal, tool-free terminal turn. The hook receives bounded run facts: terminal text, whether a mutation occurred, and the original formatted user message.

When the hook returns a static correction:

1. Agent Core records the rejected assistant turn and a synthetic correction turn in model history.
2. It starts one more model turn without publishing `onDone`.
3. The existing UI bubble is reused, so rejected prose is replaced by continued tool activity or the corrected final response.
4. Each run permits at most one review-triggered retry. A second rejection is finalized normally, preventing loops.

`composeSkills` forwards the first child review that returns a correction.

### Slides denial policy

The Slides hook activates only when:

- no mutation has occurred; and
- the terminal response asserts that font/text color or title/element positioning cannot be edited with the available interface.

It recognizes the concrete Chinese and English denial families while requiring both a denial phrase and a supported capability phrase. Questions, real unsupported features, tool errors, and successful edits are not rejected. The correction contains only static capability guidance—never slide text, paths, IDs, or raw errors.

## Boundaries

- No automatic target selection or write is performed by the guard.
- Existing durable targets, transaction receipts, stale checks, cancellation, undo/history, and QC remain authoritative.
- The guard cannot override an `unsupported`, `conflict`, `uncertain`, or cancelled tool result.
- Files, Docs, Sheets, Office, and other skills are unchanged unless they explicitly opt into the hook.

## Failure handling

- Hook exceptions fail open to the normal terminal response.
- Cancellation and turn-limit finalization bypass review.
- The correction retry counts as a normal model turn and remains subject to all existing loop and repeated-tool breakers.
- One correction maximum prevents retry storms.

## Verification

- Agent Core tests: retry once, reuse the same run, no premature `onDone`, no infinite retry, cancellation/finalization bypass, compose forwarding.
- Slides tests: exact Chinese screenshot denial and equivalent English denial are rejected; questions and genuine unsupported statements pass; a prior mutation bypasses rejection.
- Existing Slides text/style/geometry transaction tests continue to prove that the corrected tool calls remain safe and authoritative.

## Rollback

Remove the Slides hook and prompt contract to restore previous behavior. The optional Agent Core hook is inert for all skills that do not implement it, so rollback does not require history, file-format, or protocol migration.
