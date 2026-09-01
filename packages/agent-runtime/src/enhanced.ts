import type {
  AgentImage,
  AgentMessage,
  AgentRunResult,
  AgentToolCall,
  ToolExecutedEvent,
} from '@wiswork/agent-core'
import type { AgentSkill } from '@wiswork/agent-core'
import type { AgentHarnessSnapshot } from '@wiswork/agent-harness'
import type {
  PresentationCompletionFacts,
  PresentationCompletionReceipt,
} from '@wiswork/presentation-verification'
import type { AgentRuntime, AgentRuntimeSession, AgentRuntimeSessionOptions } from './session'

export type EnhancedSessionEvent =
  | { readonly type: 'text'; readonly text: string }
  | { readonly type: 'tool-start'; readonly call: AgentToolCall }
  | { readonly type: 'tool-executed'; readonly event: ToolExecutedEvent<unknown> }
  | { readonly type: 'turn-end' }
  | { readonly type: 'clarify'; readonly question: string }
  | { readonly type: 'plan'; readonly steps: string[]; readonly requiresConfirmation: boolean }
  | { readonly type: 'correction'; readonly pass: number; readonly maximum: number }
  | {
      readonly type: 'receipt'
      readonly event: { receipt: PresentationCompletionReceipt; facts: PresentationCompletionFacts }
    }
  | { readonly type: 'done'; readonly result: AgentRunResult }
  | { readonly type: 'error'; readonly code: string }

export interface EnhancedRuntimeClientSession {
  start(input: { readonly text: string; readonly images?: readonly AgentImage[] }): Promise<void>
  cancel(): Promise<void>
  close(): Promise<void>
  subscribe(listener: (event: EnhancedSessionEvent) => void): () => void
}

export interface EnhancedRuntimeClient {
  open(input: {
    readonly host: AgentRuntimeSessionOptions['host']
    readonly documentId: string
    readonly generation: number
    readonly skill: AgentSkill
  }): EnhancedRuntimeClientSession
  close(): void | Promise<void>
}

class EnhancedSession<TSnapshot> implements AgentRuntimeSession {
  readonly mode = 'enhanced' as const
  readonly #listeners = new Set<() => void>()
  readonly #remote: EnhancedRuntimeClientSession
  readonly #events: AgentRuntimeSessionOptions<TSnapshot>['events']
  #snapshot: AgentHarnessSnapshot = { status: 'idle', busy: false, generation: 0 }
  #messages: AgentMessage[] = []
  #disposed = false
  #unsubscribe: () => void

  constructor(
    remote: EnhancedRuntimeClientSession,
    options: AgentRuntimeSessionOptions<TSnapshot>,
  ) {
    this.#remote = remote
    this.#events = options.events
    this.#unsubscribe = remote.subscribe((event) => this.#event(event))
  }
  get snapshot() {
    return this.#snapshot
  }
  get messages() {
    return this.#messages
  }
  subscribe(listener: () => void) {
    this.#listeners.add(listener)
    return () => this.#listeners.delete(listener)
  }
  run(instruction: string, images?: AgentImage[]): boolean {
    if (this.#disposed || this.#snapshot.busy || !instruction) return false
    const generation = this.#snapshot.generation + 1
    this.#messages.push({ role: 'user', text: instruction, ...(images?.length ? { images } : {}) })
    this.#publish({ status: 'running', busy: true, generation })
    void this.#remote
      .start({ text: instruction, ...(images?.length ? { images } : {}) })
      .catch(() => {
        if (!this.#current(generation) || !this.#snapshot.busy) return
        this.#events?.onError?.('enhanced_turn_failed')
        this.#publish({ status: 'error', busy: false, generation, error: 'enhanced_turn_failed' })
      })
    return true
  }
  stop(): void {
    if (!this.#snapshot.busy || this.#disposed) return
    void this.#remote.cancel().catch(() => undefined)
  }
  reset(): void {
    if (this.#disposed) return
    void this.#remote.cancel().catch(() => undefined)
    this.#messages = []
    this.#publish({ status: 'idle', busy: false, generation: this.#snapshot.generation + 1 })
  }
  restore(messages: readonly AgentMessage[]): void {
    if (this.#disposed || this.#snapshot.busy) return
    this.#messages = structuredClone(messages) as AgentMessage[]
    this.#notify()
  }
  dispose(): void {
    if (this.#disposed) return
    this.#disposed = true
    this.#unsubscribe()
    this.#listeners.clear()
    void this.#remote.close().catch(() => undefined)
  }
  #current(generation: number) {
    return !this.#disposed && this.#snapshot.generation === generation
  }
  #event(event: EnhancedSessionEvent): void {
    if (this.#disposed || !this.#snapshot.busy) return
    const generation = this.#snapshot.generation
    if (event.type === 'text') {
      this.#events?.onText?.(event.text)
      return
    }
    if (event.type === 'tool-start') {
      this.#events?.onToolStart?.(event.call)
      return
    }
    if (event.type === 'tool-executed') {
      this.#events?.onToolExecuted?.(event.event as ToolExecutedEvent<TSnapshot>)
      return
    }
    if (event.type === 'turn-end') {
      this.#events?.onTurnEnd?.()
      return
    }
    if (event.type === 'clarify') {
      this.#events?.onPresentationClarify?.({ question: event.question })
      return
    }
    if (event.type === 'plan') {
      this.#events?.onPresentationPlan?.({
        steps: event.steps,
        requiresConfirmation: event.requiresConfirmation,
      })
      return
    }
    if (event.type === 'correction') {
      this.#events?.onPresentationCorrection?.({ pass: event.pass, maximum: event.maximum })
      return
    }
    if (event.type === 'receipt') {
      this.#events?.onPresentationReceipt?.(event.event)
      return
    }
    if (event.type === 'error') {
      this.#events?.onError?.(event.code)
      this.#publish({ status: 'error', busy: false, generation, error: event.code })
      return
    }
    this.#messages.push({ role: 'assistant', text: event.result.text })
    this.#events?.onDone?.(event.result)
    this.#publish({
      status: event.result.cancelled ? 'cancelled' : 'done',
      busy: false,
      generation,
    })
  }
  #publish(snapshot: AgentHarnessSnapshot) {
    this.#snapshot = snapshot
    this.#notify()
  }
  #notify() {
    for (const listener of [...this.#listeners]) {
      try {
        listener()
      } catch {}
    }
  }
}

export class EnhancedAgentRuntime implements AgentRuntime {
  readonly mode = 'enhanced' as const
  readonly #sessions = new Map<string, AgentRuntimeSession>()
  #disposed = false
  constructor(readonly client: EnhancedRuntimeClient) {}
  createSession<TSnapshot>(options: AgentRuntimeSessionOptions<TSnapshot>): AgentRuntimeSession {
    if (this.#disposed) throw new Error('agent_runtime_disposed')
    if (this.#sessions.has(options.document.id)) throw new Error('agent_session_exists')
    const session = new EnhancedSession(
      this.client.open({
        host: options.host,
        documentId: options.document.id,
        generation: options.document.generation,
        skill: options.skill,
      }),
      options,
    )
    this.#sessions.set(options.document.id, session)
    return session
  }
  async dispose(): Promise<void> {
    if (this.#disposed) return
    this.#disposed = true
    for (const session of this.#sessions.values()) session.dispose()
    this.#sessions.clear()
    await this.client.close()
  }
}
