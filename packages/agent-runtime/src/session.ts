import type {
  AgentImage,
  AgentLoopEvents,
  AgentLoopOptions,
  AgentMessage,
  AgentSkill,
} from '@wiswork/agent-core'
import type { AgentHarnessSnapshot } from '@wiswork/agent-harness'
import type { AgentRuntimeMode, EnhancedHost } from './contracts'

export interface AgentRuntimeSession {
  readonly mode: AgentRuntimeMode
  readonly snapshot: AgentHarnessSnapshot
  readonly messages: readonly AgentMessage[]
  subscribe(listener: () => void): () => void
  run(instruction: string, images?: AgentImage[]): boolean
  stop(): void
  reset(): void
  restore(messages: readonly AgentMessage[]): void
  dispose(): void
}

export interface AgentRuntimeSessionOptions<TSnapshot = unknown> extends Omit<
  AgentLoopOptions<TSnapshot>,
  'skill' | 'events'
> {
  readonly host: EnhancedHost
  readonly document: Readonly<{ id: string; generation: number }>
  readonly skill: AgentSkill
  readonly events?: AgentLoopEvents<TSnapshot>
}

export interface AgentRuntime {
  readonly mode: AgentRuntimeMode
  createSession<TSnapshot>(options: AgentRuntimeSessionOptions<TSnapshot>): AgentRuntimeSession
  dispose(): void | Promise<void>
}
