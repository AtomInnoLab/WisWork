# Office PowerPoint and WisWork Slides parity

## Product thesis

Office PowerPoint is another host for the WisWork presentation agent, not a separate agent product. A user should recognize the same workflow, state language, trust boundaries, and recovery controls in the Taskpane while the layout adapts to a narrow Office surface and Office.js capability limits.

The agent is appropriate because producing a deck requires iterative planning, research, visual construction, user checkpoints, document mutation, and post-write verification rather than a single deterministic command.

## Goals

- Give PowerPoint Taskpane and desktop Slides one presentation-agent interaction model.
- Present the same lifecycle: understand, clarify when necessary, plan, research, build, review a proposed change, apply, verify, correct, and report completion.
- Use the same visual language for messages, grouped work steps, progress, confirmation, errors, retry, stop, and undo/recovery status.
- Preserve Office's existing proposal, stale-revision check, snapshot, readback, rollback, quarantine, and raw Office permission boundaries.
- Make unsupported Office.js operations explicit without degrading the whole task when a supported alternative exists.

## Non-goals

- Pixel-for-pixel duplication of the desktop panel in a narrow Taskpane.
- Pretending Office.js has every desktop Slides editing primitive.
- Moving document authority into shared UI code or allowing the Taskpane to bypass confirmation.
- Changing Word or Excel agent semantics beyond using shared presentational primitives safely.

## Unified workflow

1. **Intent** — show the user's request immediately and create one assistant run.
2. **Context** — inspect the presentation, selected slide, attachments, and available capabilities.
3. **Clarification** — ask only when audience, purpose, or another material requirement cannot be safely inferred.
4. **Plan** — expose a compact plan with build and verification steps; do not dump protocol details.
5. **Research** — use web and image search when relevant and surface sources/images as bounded step details.
6. **Build** — perform bounded slide operations. Reads are automatic; mutations produce a reviewable proposal.
7. **Checkpoint** — show operation, target, scope, and count without raw arguments, secrets, hashes, or internal IDs.
8. **Verification** — read back affected state, run presentation postconditions and visual review where supported, and correct only inside the approved policy.
9. **Handoff** — report verified completion, partial/unverified state, or failure truthfully and expose retry/stop/undo as applicable.

## Shared interaction contract

The shared UI model is host-neutral and contains only display-safe values:

- user and assistant messages;
- grouped activity steps with `running`, `done`, or `error` state;
- plan/clarification/correction/receipt notices;
- structured confirmation summaries;
- terminal states and recovery actions.

It never carries document handles, Office objects, tool arguments, credentials, fingerprints, or raw snapshots. Desktop Slides and Office continue to own their runtime state and map it into this display contract.

## Responsive UX

- Desktop keeps its resizable side panel.
- Office uses the same hierarchy and components in a single-column layout down to 280 px.
- User messages remain quiet surface bubbles; assistant content remains unboxed.
- Tool activity is one collapsible “Working / Worked · N steps” group instead of a stack of status cards.
- The composer remains pinned, supports Enter to send, Shift+Enter for a newline, Escape/Stop to cancel, and exposes attachments through a compact action.
- Empty-state prompts are PowerPoint-specific and match desktop Slides language and intent.
- Colors come exclusively from shared semantic tokens; document colors never follow chrome theme.

## Trust and autonomy

| Action                            | Policy                                          | Rationale                                                   |
| --------------------------------- | ----------------------------------------------- | ----------------------------------------------------------- |
| Read deck/context                 | Automatic                                       | Bounded and non-mutating.                                   |
| Search web/images                 | Automatic within configured service policy      | Needed to complete the task; results remain untrusted data. |
| Prepare plan/proposal             | Automatic                                       | No document mutation.                                       |
| Apply Office mutation             | Explicit confirmation                           | User-visible document change through Office.js.             |
| Verify/read back                  | Automatic                                       | Establishes truthful completion.                            |
| Correct after failed verification | Only within the approved presentation policy    | Prevents silent scope expansion.                            |
| Continue after uncertain write    | Refused until reconciliation or document reload | Avoids compounding an unknown document state.               |

## Failure handling

- One run owns one visible progress group and one terminal result.
- A stopped or failed run removes indefinite spinners.
- Errors use localized user-facing categories and retain correlation-safe diagnostics separately.
- A tool-heavy run with no prose still ends with a meaningful completion state.
- Capability gaps name the unavailable operation and offer the supported alternative when possible.
- Office write uncertainty retains the existing document-context quarantine.

## Acceptance criteria

- PowerPoint Taskpane renders the same hierarchy, grouped progress, message treatment, confirmation states, and terminal recovery actions as desktop Slides at desktop and 280 px widths.
- A new-deck request produces context, optional clarification, a visible plan, bounded build proposals, explicit confirmation, readback verification, and a truthful receipt.
- Standard and Enhanced runtimes produce the same Taskpane event model and never fall back silently.
- Cancel, retry, stale proposal, verification failure, rollback, and uncertain-write states have deterministic tests.
- Existing Word/Excel behavior, Office security tests, Slides presentation tests, theme checks, typechecks, and production builds remain green.

## Alternatives considered

1. Copy desktop `AiPanel.tsx` into the add-in: rejected because it duplicates state and CSS and would drift again.
2. Embed the desktop renderer in Office: rejected because Office authority, CSP, dimensions, and lifecycle differ.
3. Share only colors: rejected because the main inconsistency is lifecycle semantics and component behavior, not just styling.

## Validation and rollback

- Characterization tests first, followed by shared contract/component tests and Office PowerPoint integration tests.
- Test both Standard and Enhanced event streams against the same presentation.
- Run Office and Slides complete suites, root typecheck, theme-color check, production builds, and the existing presentation acceptance E2E where the environment supports Electron.
- Changes stay behind existing PowerPoint host selection; rollback removes the shared presentation view adapter and restores the prior Office rendering without changing stored documents or protocols.
