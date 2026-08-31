# Presentation Agent Planning and Verification Design

## Status

Approved direction. This design covers both the native WisWork PC Slides agent and the PowerPoint Office Add-in Taskpane.

## Design thesis

Presentation edits are complete only when the agent can prove both instruction fidelity and rendered quality. WisWork will therefore turn each supported editing request into a bounded plan and acceptance contract, execute through the host's existing transaction boundary, inspect the host-resolved rendering, perform at most two minimal corrective passes, and report evidence instead of an unsupported success or capability denial.

## Goals

- Give PC Slides and PowerPoint Taskpane the same user-facing planning, confirmation, verification, and completion semantics.
- Ask a question only when missing scope, reference, tolerance, or destructive intent would materially change the result.
- Show a concise plan before cross-slide, compound, or elevated-risk edits.
- Verify exact properties structurally and rendered outcomes visually.
- Allow at most two automatic corrective passes for reversible, already-supported low-risk edits.
- Preserve host-specific authority: PC durable targets and deck transactions; PowerPoint Office objects/package receipts and proposal confirmation.
- Never convert an applied-but-unverifiable mutation into “nothing changed,” and never claim verified success without evidence.

## Non-goals

- General autonomous redesign of arbitrary decks.
- Flattening slides into images or replacing editable content with raster output.
- Automatically repairing masters, charts, tables, animations, or arbitrary OOXML after visual review.
- A new unrestricted screenshot or code-execution tool.
- Sharing raw slide text, file paths, identifiers, screenshots, or credentials in diagnostics or persisted acceptance receipts.

## Product workflow

1. **Understand** — classify affected pages, reference pages, requested properties, risk, and available tool families.
2. **Clarify when necessary** — ask one bounded question when scope/reference/tolerance or permission is missing. Clear requests proceed without extra friction.
3. **Plan and confirm** — show a short plan for compound/cross-slide work. Existing confirmation boundaries remain authoritative for destructive or elevated-risk operations.
4. **Freeze contract** — compile the approved intent into a strict acceptance contract bound to document/session identity and the pre-edit revision.
5. **Execute** — use the existing host transaction/proposal path. A plan may not silently expand its targets.
6. **Verify structure** — check requested values, targets, revision, and unintended structural changes.
7. **Verify rendering** — capture only affected/reference pages from the current authoritative revision and run a task-specific visual reviewer.
8. **Correct, bounded** — at most two minimal passes, restricted to failed checks and low-risk supported tool families.
9. **Deliver evidence** — publish a bounded completion receipt and localized summary with undo/rollback affordance where available.

## When to ask and when to plan

### Clarification required

- “Make these consistent” has no reference page or explicit style source.
- The user mentions a selection/range that cannot be resolved authoritatively.
- Fixing overflow would require a material choice between resizing, rewriting, or changing layout.
- The request conflicts with a locked/master-owned object and changing the master is not explicitly authorized.
- A requested change would expand beyond the named pages or objects.

### Plan shown without blocking confirmation

- Two or more affected pages.
- Two or more property families, such as color plus geometry.
- A reference-page style transfer.
- A task that requires post-edit screenshot comparison.

### Explicit confirmation retained

- Object deletion or broad re-layout under the existing product policy.
- Master, chart, table, animation, package, and OOXML writes.
- Any correction that would broaden the frozen acceptance contract.

## Shared contracts

The shared package exposes strict bounded types and parsers. No model-provided executable validators are accepted.

```ts
type PresentationAcceptanceCheck =
  | {
      id: string
      kind: 'element_property'
      slide: number
      roleOrTarget: TargetRef
      property: SupportedProperty
      expected: SafeScalar
    }
  | {
      id: string
      kind: 'reference_match'
      slide: number
      referenceSlide: number
      role: SupportedRole
      properties: SupportedProperty[]
      tolerance: number
    }
  | { id: string; kind: 'render_quality'; slide: number; rules: RenderRule[] }

type PresentationAcceptanceContract = {
  version: 1
  taskId: string
  documentToken: string
  sessionToken: string
  baseRevision: string
  affectedSlides: number[]
  referenceSlides: number[]
  checks: PresentationAcceptanceCheck[]
  maxCorrectionPasses: 0 | 1 | 2
}

type PresentationCompletionReceipt = {
  version: 1
  taskId: string
  status: 'verified' | 'applied_unverified' | 'needs_user' | 'failed' | 'unchanged'
  mutationReceiptIds: string[]
  passedCheckIds: string[]
  failedCheckIds: string[]
  unavailableCheckIds: string[]
  correctionPasses: number
  affectedSlides: number[]
  rollbackId?: string
  safeCode?: string
}
```

Global bounds: at most 50 checks, 50 affected pages for deterministic checks, 8 screenshots per visual pass, 2 MiB serialized visual request, two correction passes, 64 safe observations, and no raw content in receipts.

## Verification architecture

### Deterministic verifier

Proves exact facts: target identity/type, text/style values, geometry within tolerance, page/background state, revision ownership, and absence of unauthorized target expansion. Checks are operation-specific; unsupported checks become unavailable, never successful.

### Visual verifier

Receives only the bounded acceptance contract summary, deterministic results, current rendered screenshots, optional reference screenshots, and safe structural hints. It returns strict data:

```ts
type VisualReviewResult = {
  status: 'pass' | 'needs_fix' | 'cannot_verify'
  failedCheckIds: string[]
  observations: SafeObservation[]
  fixIntents: SafeFixIntent[]
}
```

The visual reviewer has no direct host access. The orchestrator validates every fix intent against the frozen contract before compiling it through normal tools.

### PC Slides adapter

- Authority: durable slide/element identities, transaction revision, and current renderer session.
- Rendering: existing slide capture after authoritative deck refresh.
- Mutation: presentation transaction executor and existing history batches.
- Correction: text/style/geometry/fill/stroke/background operations already migrated to canonical transactions.
- Existing generic QC remains a deck-health signal, but task verification runs first and uses the user's acceptance checks.

### PowerPoint Taskpane adapter

- Authority: Office session/document identity, proposal target hashes, Office object state, and package verification receipts.
- Rendering: `screenshot_slide` from the post-write PowerPoint state, revision-bound by the orchestrator.
- Mutation: existing proposal controller and browser PowerPoint adapter.
- Correction: low-risk native text/style/geometry operations when supported; package/master/chart/table/OOXML changes remain confirmation-only and never visual-autofix.
- Office readback lag uses existing bounded convergence. A screenshot/readback failure after a proved write produces `applied_unverified`.

## Autonomy and trust boundaries

- Automatic: read-only inspection, plan construction, deterministic validation, screenshots of affected/reference pages, and up to two low-risk corrections within the approved contract.
- Confirmation: existing elevated-risk mutations and any proposed scope expansion.
- Recommendation only: unsupported or unsafe correction families.
- Refusal/fail closed: stale document/session/revision, ambiguous target, malformed contract, screenshot from a different revision, or receipt mismatch.
- Every applied completion exposes a bounded audit trail and existing undo/rollback capability when the host provides one.

## Failure semantics

- `verified`: mutation and every required check proved.
- `applied_unverified`: mutation proved applied, but one or more required visual checks were unavailable. Never auto-retry the mutation.
- `needs_user`: a material choice or elevated-risk correction requires confirmation.
- `failed`: no approved mutation was proved applied, or the operation failed before dispatch.
- `unchanged`: deterministic proof that no mutation occurred or the requested state already existed.
- `office_state_uncertain` and equivalent host uncertainty remain explicit and stop correction.

Cancellation before dispatch is unchanged/cancelled. Cancellation after dispatch must reconcile receipts and cannot erase an applied result.

## Final response policy

Final prose is rendered from the completion receipt. Free-form model text may add a short explanation but cannot contradict status or counts. The response lists affected pages, passed checks, unavailable/failed checks, correction count, and rollback availability. A run without a terminal receipt may not claim completion.

## Rollout and rollback

- Feature flags: shared contract generation, PC task verification, Taskpane task verification, and automatic correction are independently switchable.
- Start with text, style/color, geometry, fill/stroke, and solid background.
- Shadow-evaluate contracts before enabling automatic corrections.
- Rollback disables orchestration and returns both products to existing transaction/QC behavior without changing document formats or persisted chats.

## Verification and release gates

- Contract/parser fuzz and strict-bound tests.
- Agent Core tests for clarify/plan/confirm/receipt-finalization and cancellation.
- PC integration fixtures for cross-page color/geometry/reference matching, screenshot staleness, two-pass correction, and applied-unverified semantics.
- PowerPoint adapter tests for post-write screenshots, Office lag, stale proposal, elevated-risk no-autofix, and cancellation after dispatch.
- Shared golden tasks run through both adapters with equivalent receipt outcomes.
- False-success, false-denial, unintended-change, verified-completion, and manual-rework metrics.
- PC package and Taskpane production builds; no Relay, protocol, manifest, or pairing change unless implementation evidence later proves necessary.
