import { AgentLoop, type AgentLoopOptions } from '@wiswork/agent-core'
import type {
  AgentEvent,
  AgentEventListener,
  AgentRuntime,
  AgentSession,
  AgentTurnInput,
} from './contracts'

export const LEGACY_ERROR_MAX_BYTES = 1_024

export interface LegacyAgentSessionOptions<TSnapshot = unknown> extends Omit<
  AgentLoopOptions<TSnapshot>,
  'events'
> {
  documentId: string
}

class LegacyAgentSession<TSnapshot> implements AgentSession<TSnapshot> {
  readonly runtimeKind = 'legacy' as const
  private readonly listeners = new Set<AgentEventListener<TSnapshot>>()
  private readonly loop: AgentLoop<TSnapshot>
  private closed = false
  private closePromise: Promise<void> | undefined

  constructor(
    readonly documentId: string,
    options: Omit<LegacyAgentSessionOptions<TSnapshot>, 'documentId'>,
    private readonly onClose: () => void,
  ) {
    this.loop = new AgentLoop<TSnapshot>({
      ...options,
      events: {
        onText: (text) => this.emit({ type: 'text', text }),
        onToolStart: (call) => this.emit({ type: 'tool-start', call }),
        onToolExecuted: ({ call, execution, snapshotBefore }) =>
          this.emit({ type: 'tool-executed', call, execution, snapshotBefore }),
        onTurnEnd: () => this.emit({ type: 'turn-end' }),
        onDone: (result) => this.emit({ type: 'done', result }),
        onError: () =>
          this.emit({
            type: 'error',
            code: 'legacy_turn_failed',
            message: 'Legacy turn failed.',
          }),
      },
    })
  }

  get busy(): boolean {
    return !this.closed && this.loop.busy
  }

  subscribe(listener: AgentEventListener<TSnapshot>): () => void {
    if (this.closed) return () => undefined
    this.listeners.add(listener)
    let subscribed = true
    return () => {
      if (!subscribed) return
      subscribed = false
      this.listeners.delete(listener)
    }
  }

  startTurn({ text, images }: AgentTurnInput): void {
    if (this.closed) return
    this.loop.run(text, images)
  }

  cancel(): void {
    if (this.closed) return
    this.loop.cancel()
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise
    this.closed = true
    this.listeners.clear()
    this.loop.reset()
    this.onClose()
    this.closePromise = Promise.resolve()
    return this.closePromise
  }

  private emit(event: AgentEvent<TSnapshot>): void {
    if (this.closed) return
    for (const listener of [...this.listeners]) {
      try {
        listener(event)
      } catch {
        // UI observers are isolated from the loop and from each other.
      }
    }
  }
}

export class LegacyAgentRuntime<TSnapshot = unknown> implements AgentRuntime<
  LegacyAgentSessionOptions<TSnapshot>,
  TSnapshot
> {
  readonly kind = 'legacy' as const
  private readonly sessions = new Map<string, LegacyAgentSession<TSnapshot>>()
  private closed = false
  private closePromise: Promise<void> | undefined

  openSession(options: LegacyAgentSessionOptions<TSnapshot>): AgentSession<TSnapshot> {
    if (this.closed) throw new Error('legacy agent runtime is closed')
    if (this.sessions.has(options.documentId)) {
      throw new Error(`agent session already open for document: ${options.documentId}`)
    }

    const { documentId, ...loopOptions } = options
    const session = new LegacyAgentSession(documentId, loopOptions, () => {
      if (this.sessions.get(documentId) === session) this.sessions.delete(documentId)
    })
    this.sessions.set(documentId, session)
    return session
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise
    this.closed = true
    const sessions = [...this.sessions.values()]
    this.closePromise = Promise.all(sessions.map((session) => session.close())).then(() => {
      this.sessions.clear()
    })
    return this.closePromise
  }
}
