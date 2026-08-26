# Shared Agent Harness Design

## Design thesis

WisWork will use one shared Agent Harness implementation across the desktop editors and the Office Taskpane, while preserving one isolated harness instance per open document and keeping every host's tools, permissions, snapshots, and write verification inside that host.

## Goals

- Give Docs, Sheets, Slides, Markdown, and the Office Taskpane the same run lifecycle, cancellation, reset, stale-callback suppression, history access, and loop configuration contract.
- Preserve the existing `@wiswork/agent-core` ReAct engine and all observable host behavior.
- Make future model, timeout, compaction, loop-breaker, and lifecycle changes land once.
- Provide a conformance surface that can be tested without mounting a renderer UI.

## Non-goals

- Do not share one live Agent instance between documents or hosts.
- Do not move Office.js execution, confirmation, fingerprints, recovery, or requirement-set checks into WisWork PC or Relay.
- Do not merge Docs/Sheets/Slides/Markdown tool registries.
- Do not create implicit cross-document memory or permissions.
- Do not change Relay, model-provider, or Office tool wire protocols.

## Architecture

Create `@wiswork/agent-harness` as a small runtime layer above `@wiswork/agent-core`. The harness owns exactly one `AgentLoop`, exposes a stable lifecycle API and observable run state, guards callbacks by generation/disposal, and forwards typed lifecycle/tool events to a host presentation adapter. Each host supplies its own transport, skill, snapshots, formatting, and event projection.

```text
@wiswork/agent-harness
  └─ owns one @wiswork/agent-core AgentLoop
       ├─ desktop transport + Docs/Sheets/Slides/Markdown skill
       └─ Relay transport + Office host skill/confirmation wrapper
```

## Public contract

The package exports:

```ts
type AgentHarnessStatus = 'idle' | 'running' | 'done' | 'cancelled' | 'error'

interface AgentHarnessSnapshot {
  status: AgentHarnessStatus
  busy: boolean
  generation: number
  error?: string
}

interface AgentHarness<TSnapshot> {
  readonly snapshot: AgentHarnessSnapshot
  readonly messages: readonly AgentMessage[]
  subscribe(listener: () => void): () => void
  run(instruction: string, images?: AgentImage[]): boolean
  stop(): void
  reset(): void
  restore(messages: readonly AgentMessage[]): void
  dispose(): void
}

function createAgentHarness<TSnapshot>(
  options: AgentLoopOptions<TSnapshot>,
): AgentHarness<TSnapshot>
```

`run` returns `false` for an empty instruction, a busy harness, or a disposed harness. `stop` cancels the current run but preserves history. `reset` cancels work and clears history. `dispose` is terminal, cancels work, clears listeners, and suppresses all later callbacks. Host callbacks supplied in `options.events` are invoked only for the current generation and before the harness publishes the corresponding state transition.

The harness does not reinterpret tool results, errors, prompts, or confirmation decisions. Those remain owned by `AgentLoop` and the host skill wrapper.

## Host boundaries

### Shared

- AgentLoop ownership and construction
- idle/running/done/cancelled/error state
- cancel/reset/dispose generation rules
- stale callback suppression
- history restoration/access
- subscription mechanics

### Host-specific

- transport selection and authentication
- skill composition and dynamic context
- Office confirmation suspension and proposal UI
- snapshot capture and rollback UI
- chat persistence and presentation timeline
- document mutation, semantic verification, and recovery
- host diagnostics beyond generic run status

## Isolation and permissions

Every open document owns a distinct harness instance, history, AbortController path, skill, and transport handle. No harness may access another document's tools or messages. Office tools continue executing in the Taskpane's Office.js context and retain confirmation-first semantics. Relay remains an authenticated byte-stream broker and WisWork PC remains the authenticated model proxy; neither gains document authority.

## Failure handling

- A late event after `reset` or `dispose` is ignored by the harness and cannot repaint UI state.
- `stop` produces a cancelled terminal state only through the current generation.
- A synchronous construction or host callback error must not strand `busy=true`; core loop failures retain their existing stable error behavior.
- Disposing one document cannot cancel another harness.
- Host-specific confirmation/recovery failures retain their existing stable codes.

## Migration and rollback

Migration is source-compatible and staged: add the package, migrate Office, then migrate desktop editors. Each host replaces direct `new AgentLoop(...)` construction with `createAgentHarness(...)` but keeps its existing options and event callbacks. A host can be rolled back independently by restoring direct AgentLoop construction; no persisted data or protocol migration is involved.

## Verification

- Package tests cover lifecycle state, callback ordering, generation suppression, restore, independent instances, and disposal.
- Office conformance tests cover confirmation suspension, stop/reset/logout, and no late UI repaint.
- Desktop focused tests cover run, cancellation, persisted history/snapshots, and independent editors.
- Full tests/typechecks/builds for agent-core, agent-harness, Office, Docs, Sheets, Slides, Markdown, Shell, and AI Provider must pass.
- Manual release gate: concurrent Word/Excel/PowerPoint sessions plus one desktop editor, with one session cancelled and the others continuing.
