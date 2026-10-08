# Office production recovery

Baseline: `867cc93224cfa528d133694da2340f508bc6f72c` (PR #201).

The user authorized fixing the defects found in the deployed Office plugin and previously chose rendered DESIGN.md in the plugin with editing in the connected WisWork PC and save synchronization.

## Design

Keep the existing Office mutation and visual-verification authorities. Treat a searched image candidate as download plus decoding/normalization: failure of either stage can select the already-authorized backup, within bounded time, bytes, pixels and cancellation. Serialize remote tool dispatch at the existing single-flight boundary, preserve call identity, and settle queued/in-flight work on cancellation without poisoning the next turn.

Negative visual reviews always invalidate an old pass. An applied-but-unverified repair may be accepted only after a new screenshot and a positive visual review; repeated reviews without a repair cannot clear a failure. Native screenshot failures must remain truthful and provide a concrete same-page recovery route, never synthetic acceptance.

Reuse the shared Markdown renderer and the PC Markdown tab. Add a narrowly negotiated `design-document.v1` capability using the existing bounded Relay request transport for opening and reading a task-scoped DESIGN.md. The PC chooses its own file path; no path or arbitrary filesystem access is accepted from the plugin. Read/save synchronization is scoped to the connected session and base revision, with conflicts surfaced instead of overwriting newer documents. Saving updates the visible draft; applying a changed contract uses the existing validated revision workflow, avoiding silent slide edits.

For image composition, use the existing browser canvas image preparation to preserve aspect ratio in the requested rectangle and supply explicit fit semantics. Do not weaken shape geometry, source provenance, SSRF protection, screenshot inspection, or final completion gates. Add actionable screenshot recovery and preserve honest limitations when the native host cannot render.

## Verification and release

Turn each reproduced defect into a failing regression before implementation; exercise real dispatch/state-machine code with only network/Office boundaries mocked. Validate cross-component capability negotiation, session isolation, content limits, stale revisions, cancellation and late results. Run the full repository test/typecheck/lint/build requirements and Relay tests for the protocol extension, followed by independent review. Real Mac PowerPoint visual acceptance is reported separately from automated tests. Keep the isolated branch and reviewable commits; no production deployment or version bump is implied by this repair turn.

Rollback uses the prior PC and Taskpane builds and disables the optional new capability by negotiation. No existing document migration or deletion is needed.

## Rollout dependencies and remaining acceptance

Release all three components for the complete change: Relay (optional capability negotiation and retired-result handling), WisWork PC (image normalization, serialized tools, bounded local design files), and Taskpane (aspect-preserving insertion, review recovery, DESIGN.md reader). Older pairings keep their original capability scope; reconnect by creating a new approved pairing to enable DESIGN.md editing. Updating Taskpane alone does not enable the PC file channel or repair PC image decoding.

The design reader checks PC saves every five seconds while the current document is open and the agent is idle. Closing it pauses reads without discarding the draft. New tasks and replacement sessions discard only the UI association, not the user's local file. The PC retains duplicate-request detection for the whole session with a bounded request budget covering eight hours of reads plus the original interactive allowance.

Real-host acceptance is still required on Mac PowerPoint: generate an eight-page deck with photographic cover and content pages, inspect actual screenshot transport and readable/cropped images, repair a deliberately failed visual review, complete final geometry/visual verification, cancel an in-flight tool and start another task, then open/edit/save/apply DESIGN.md through the paired PC. Native screenshot failure remains a truthful blocker for that page's final acceptance; no synthetic screenshots or bypassed reviews are introduced.
