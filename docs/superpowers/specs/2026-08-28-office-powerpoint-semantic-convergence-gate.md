# Office PowerPoint semantic convergence release gate

Status: **NO-GO — defer runtime convergence**

Decision date: 2026-08-28

Scope: `@wiswork/presentation-ops`, Desktop Slides, and the Office PowerPoint Taskpane

## Decision

Do not import `@wiswork/presentation-ops` from
`apps/office-addin/src/skills/powerpoint/**` in this release train. The prerequisite in the
selective-sync plan says that the Desktop receipt contract must be stable for one release. It is
not satisfied:

- the contract was introduced by `2f7e0d4` on 2026-08-27;
- transaction, receipt, target, and QC behavior continued changing on 2026-08-28, through at
  least `9018de5` in this branch;
- the newest repository release tag inspected, `v0.8.440`, does not contain `2f7e0d4`; and
- no release tag contains `2f7e0d4` (`git tag --contains 2f7e0d4` is empty).

For this gate, “stable for one release” means that one production release tag contains the
contract and Desktop executor, the packaged macOS and Windows acceptance matrix below completes,
and no incompatible contract or receipt-semantic change is required during that release's
observation period. A branch-only implementation or a tag on another ancestry does not qualify.

This is a documentation-only deferral. It changes no Office tools, Desktop executor, Relay,
manifest, permissions, or advertised capability.

## Authority boundaries that must remain unchanged

Semantic convergence may share pure parsers, types, fingerprints, operation names, and receipt/QC
meanings. It must not share host authority or mutation executors.

The Office implementation remains authoritative for all of these boundaries:

- `StructuredProposalController` suspends every write and requires explicit user confirmation;
- confirmation revalidates the proposal immediately before the write;
- Office.js performs the mutation under the existing bounded adapter and requirement-set checks;
- host-specific conditional readback handles delayed PowerPoint visibility;
- verification distinguishes verified, unchanged/concurrent, and third/uncertain state;
- existing recovery restores recoverable mutations and fails closed when restoration cannot be
  proven;
- Office diagnostics continue emitting only stable safe codes and phase metadata;
- cancellation after dispatch reconciles state instead of claiming that no write occurred; and
- Relay transport, pairing, timeouts, frames, and cancellation protocol remain unchanged.

The Desktop implementation remains authoritative for deck leases, transaction scheduling,
snapshots, durable PPTX identifiers, target enrollment, package persistence, and one-step atomic
recovery. A shared receipt must describe outcomes; it must never grant Office file-system, package,
Electron, or Desktop IPC authority.

## Capability map

“Candidate” means semantic overlap worth testing after the release gate. It does not mean that the
Office tool may advertise the canonical operation today.

| Canonical operation    | Current Desktop semantics                                                                                  | Current Office capability                                                                                                | Gate result                                                                                                                              |
| ---------------------- | ---------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------------------------------------------- |
| `set_text`             | Plain or structured paragraphs/runs; durable/generated target                                              | `edit_slide_text` and `set_shape_text` support plain text by slide index and shape ID                                    | Candidate only for the plain-text subset. Formatted paragraphs, generated targets, and canonical durable slide identity are unsupported. |
| `set_geometry`         | Points plus optional rotation                                                                              | `set_shape_geometry` supports left/top/width/height                                                                      | Candidate only without rotation and only after identity/fingerprint parity. Rotation is unsupported.                                     |
| `set_fill`             | Shape fill: none or solid                                                                                  | No native declarative shape-fill operation                                                                               | Unsupported. XML editing is not a semantic substitute and must not be advertised as this operation.                                      |
| `set_stroke`           | Shape stroke or no stroke                                                                                  | No native declarative shape-stroke operation                                                                             | Unsupported.                                                                                                                             |
| `add_text_box`         | Allocates a durable created target keyed by `clientId` for later operations in the same atomic transaction | `add_text_box` creates a named shape and verifies its returned Office shape ID                                           | Partial only. Do not map until Office can provide canonical created-target identity and same-transaction references.                     |
| `delete_element`       | Deletes a durable/generated target atomically                                                              | `delete_shape` deletes by slide index and shape ID with verification/recovery                                            | Candidate only after canonical target identity and receipt mapping are proven.                                                           |
| `set_speaker_notes`    | Writes bounded notes transactionally                                                                       | No Office speaker-notes write tool                                                                                       | Unsupported.                                                                                                                             |
| `set_slide_background` | Solid slide background through the Desktop transaction                                                     | Slide-package XML can edit a background; native Office tool edits master backgrounds, not this canonical slide operation | Unsupported as a canonical operation. Package XML and master operations retain their existing explicit tools.                            |

Office-only or differently scoped operations stay outside the canonical Desktop map:

| Office capability                                     | Treatment                                                                                     |
| ----------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| `duplicate_slide`                                     | Keep as an Office proposal tool; no matching canonical operation.                             |
| `insert-image`                                        | Keep as an Office proposal/VFS tool; no matching canonical operation.                         |
| `edit_slide_chart` and `edit_slide_xml`               | Keep bounded allowlisted package-edit tools; never claim canonical shape-operation semantics. |
| `edit_slide_master`                                   | Keep native API 1.10 master/theme/layout semantics and current platform gating.               |
| `edit_slide_master_xml`                               | Keep existing non-Mac platform restriction and package verification.                          |
| `inspect_slide_masters`, screenshot, shape/text reads | Read-only Office capabilities; no transaction receipt mapping required.                       |

Until the gate passes, the Office tool registry must expose only its actual host capabilities. An
unsupported canonical operation must fail before proposal creation with a stable
`office_api_unsupported`/`invalid_tool_input` policy chosen by the eventual adapter; it must never
fall through to XML, silently drop fields such as rotation, or partially apply a transaction.

## Receipt and QC gaps to close before implementation

The canonical receipt currently has `applied`, `unchanged`, `conflict`, and `uncertain`. Office has
additional safety distinctions that must not be erased:

- `proposal_stale` can arise from selection/deck/target drift and is not automatically equivalent
  to canonical `target_stale`;
- `office_verify_failed` does not by itself prove `unchanged`;
- `office_concurrent_change` and `office_state_uncertain` require three-state reconciliation;
- `office_recovery_failed` is a critical host outcome with no lossless canonical receipt code;
- rejection and pre-write cancellation are proposal decisions, not mutation receipts; and
- cancellation racing a dispatched Office callback may produce verified applied or uncertain
  state, never an assumed cancellation receipt.

Before coding, approve either a backward-compatible extension to the pure receipt contract or an
Office result envelope that preserves these codes around a canonical receipt. Mapping distinct
Office outcomes to an inaccurate existing status is a release blocker.

Desktop deterministic/visual QC produces bounded `PresentationQualityReceipt` objects tied to a
transaction, durable slide ID, and quality-run ID. Office `verify_slides` currently reports bounded
geometry checks, while `screenshot_slide` is a model-visible image path; neither is the same
contract. Office must not advertise canonical QC until it can bind findings to canonical durable
identity and transaction revision, enforce the same finding/evidence budgets, report truncation,
and return explicit screenshot/transport/cancellation/stale-session unavailability.

## Future implementation shape

After the release prerequisite passes:

1. Add a pure Office capability-map module with an allowlist for exact semantic subsets.
2. Parse canonical input with `@wiswork/presentation-ops`, then translate only allowlisted
   operations into the existing Office declarative program.
3. Preserve the current proposal preview and suspend for confirmation before any Office.js call.
4. Revalidate canonical target identity/fingerprint during proposal validation.
5. Execute through the existing Office adapter, not the Desktop executor.
6. Perform existing delayed readback and recovery, then construct a canonical receipt only from
   the proven final state.
7. Keep each operation family wholly on legacy or canonical routing behind a Taskpane-local rollout
   switch. Never mix routes within one transaction.
8. Do not change Relay frames or require a Relay/PC/manifest rollout.

## Required automated test matrix

Each candidate operation and every supported batch combination must cover:

- exact parsing, unknown fields, unsafe prototype keys, non-finite geometry, text/identifier limits,
  operation-count limits, and unsupported canonical operations;
- durable slide and element identity, wrong type, missing target, ambiguous target, stale fingerprint,
  and expected deck revision drift;
- proposal suspension, explicit confirmation, rejection, new-turn cancellation, logout cancellation,
  cancellation before write, during Office callback, and after write before verification;
- immediate readback, delayed visibility within the bounded retry policy, retry exhaustion,
  verified applied, proven no-op, concurrent third state, and read failure;
- failure before the first operation, failure after each operation, full recovery, recovery failure,
  and no accidental continuation after uncertainty;
- duplicate transaction ID with identical and different payloads, in-flight duplicate requests,
  capacity exhaustion, and request/session isolation;
- structured receipt bounds, safe diagnostics, no slide text/image/path in receipts or diagnostics,
  and unchanged Relay behavior; and
- unsupported Desktop-only operations absent from Office tool descriptions and schemas.

QC tests must separately cover all deterministic codes, bounded evidence, finding truncation,
visual unavailable/cancelled/stale-session/capacity states, stable identity mapping, and screenshots
remaining outside diagnostics.

## Real-file and release acceptance

The convergence release is GO only when all of the following are recorded against a release
candidate that already passed the one-release Desktop observation prerequisite:

- packaged macOS and Windows Desktop runs on real PPTX fixtures containing tables, charts, images,
  groups, themes, masters, notes, and non-English fonts;
- PowerPoint Taskpane runs on supported Windows and Mac hosts and requirement-set variants;
- at least ten-slide edits cover plain text, geometry, add/delete, stale targets, delayed visibility,
  rejection, cancellation at every phase, partial Office sync, third state, recovery, and reopen;
- operation and receipt parity is checked semantically, without requiring identical host executor
  steps;
- unsupported operations are absent from the model-visible Office capability set;
- diagnostics contain only stable safe codes and bounded metadata;
- full repository format, lint, typecheck, test, build, branding, notice, and license gates pass; and
- telemetry/diagnostic review during a canary shows no increase in uncertain state, recovery failure,
  confirmation bypass, or pairing failure.

Any failed item keeps the Office family on the legacy route.

## Rollout and rollback

Ship candidate families one at a time behind a Taskpane-local switch, starting with the smallest
exact subset (plain `set_text` or non-rotated `set_geometry`). Do not remove the legacy route until
real-file acceptance and the canary observation window complete.

Rollback disables the affected family switch and redeploys the previous Taskpane bundle. Existing
Office tool names, proposals, chat history, and document files remain valid because there is no
file-format migration. Relay, WisWork PC, and manifest remain untouched. If receipt interpretation
or target identity is uncertain, stop new writes, retain diagnostics, inspect the document, and do
not automatically retry.

## Re-evaluation checklist

Re-run this gate only after a release tag contains the contract and Desktop implementation. Record
the qualifying tag, packaged acceptance artifacts, observation dates, any contract changes during
the period, and the approved lossless Office receipt mapping. Without all five, the answer remains
NO-GO.
