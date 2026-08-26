import type { AgentImage, AgentRunResult, AgentToolCall, ToolExecution } from '@wiswork/agent-core'

export type AgentRuntimeKind = 'legacy' | 'codex'

/** Fail closed to the established runtime so an invalid flag is also a rollback. */
export function selectAgentRuntime(value: unknown): AgentRuntimeKind {
  return value === 'codex' ? 'codex' : 'legacy'
}

export interface AgentTurnInput {
  text: string
  images?: AgentImage[] | undefined
}

export interface AgentRuntimeError {
  code: string
  message: string
}

export type AgentEvent<TSnapshot = unknown> =
  | { type: 'text'; text: string }
  | { type: 'tool-start'; call: AgentToolCall }
  | {
      type: 'tool-executed'
      call: AgentToolCall
      execution: ToolExecution
      snapshotBefore?: TSnapshot | undefined
    }
  | { type: 'turn-end' }
  | { type: 'done'; result: AgentRunResult }
  | ({ type: 'error' } & AgentRuntimeError)

export type AgentEventListener<TSnapshot = unknown> = (event: AgentEvent<TSnapshot>) => void

export interface AgentSession<TSnapshot = unknown> {
  readonly documentId: string
  readonly runtimeKind: AgentRuntimeKind
  readonly busy: boolean
  subscribe(listener: AgentEventListener<TSnapshot>): () => void
  startTurn(input: AgentTurnInput): void | Promise<void>
  cancel(): void
  close(): Promise<void>
}

export interface AgentRuntime<TOpen, TSnapshot = unknown> {
  readonly kind: AgentRuntimeKind
  openSession(options: TOpen): AgentSession<TSnapshot> | Promise<AgentSession<TSnapshot>>
  close(): Promise<void>
}
