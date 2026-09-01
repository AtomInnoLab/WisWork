import { createAgentHarness } from '@wiswork/agent-harness'
import type { AgentRuntime, AgentRuntimeSession, AgentRuntimeSessionOptions } from './session'

export class StandardAgentRuntime implements AgentRuntime {
  readonly mode = 'standard' as const
  readonly #sessions = new Set<AgentRuntimeSession>()
  #disposed = false

  createSession<TSnapshot>(options: AgentRuntimeSessionOptions<TSnapshot>): AgentRuntimeSession {
    if (this.#disposed) throw new Error('agent_runtime_disposed')
    const harness = createAgentHarness({
      transport: options.transport,
      skill: options.skill,
      events: options.events,
      ...(options.maxTurns === undefined ? {} : { maxTurns: options.maxTurns }),
      ...(options.maxHistory === undefined ? {} : { maxHistory: options.maxHistory }),
      ...(options.compaction === undefined ? {} : { compaction: options.compaction }),
      ...(options.captureSnapshot === undefined
        ? {}
        : { captureSnapshot: options.captureSnapshot }),
      ...(options.formatUserMessage === undefined
        ? {}
        : { formatUserMessage: options.formatUserMessage }),
      ...(options.systemSuffix === undefined ? {} : { systemSuffix: options.systemSuffix }),
    })
    const session: AgentRuntimeSession = Object.create(harness, {
      mode: { value: 'standard', enumerable: true },
    })
    this.#sessions.add(session as AgentRuntimeSession)
    const dispose = harness.dispose.bind(harness)
    Object.defineProperty(session, 'dispose', {
      value: () => {
        dispose()
        this.#sessions.delete(session as AgentRuntimeSession)
      },
    })
    return session
  }

  dispose(): void {
    if (this.#disposed) return
    this.#disposed = true
    for (const session of [...this.#sessions]) session.dispose()
    this.#sessions.clear()
  }
}
