import type { AgentStreamCallbacks, AgentStreamRequest, AgentTransport } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { EnhancedAgentRuntime, type EnhancedRuntimeClientSession } from '../src/enhanced'
import { StandardAgentRuntime } from '../src/standard'

const skill = { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() }
const options = (transport: AgentTransport, events = {}) => ({
  host: 'docs' as const,
  document: { id: 'doc', generation: 1 },
  skill,
  transport,
  events,
})

describe('Agent runtime facade', () => {
  it('wraps the established harness without changing Standard behavior', async () => {
    let callbacks!: AgentStreamCallbacks
    const transport: AgentTransport = {
      stream: (_request: AgentStreamRequest, next: AgentStreamCallbacks) => {
        callbacks = next
        return { cancel: vi.fn() }
      },
    }
    const done = vi.fn()
    const runtime = new StandardAgentRuntime()
    const session = runtime.createSession(options(transport, { onDone: done }))
    expect(session.mode).toBe('standard')
    expect(session.run('hello')).toBe(true)
    await vi.waitFor(() => expect(callbacks).toBeDefined())
    callbacks.onDelta('answer')
    callbacks.onDone()
    expect(done).toHaveBeenCalledWith(expect.objectContaining({ text: 'answer', cancelled: false }))
    expect(session.messages).toEqual([
      { role: 'user', text: 'hello' },
      { role: 'assistant', text: 'answer' },
    ])
  })

  it('maps Enhanced events to the same UI contract and never falls back to Standard', async () => {
    let emit!: (event: any) => void
    const remote: EnhancedRuntimeClientSession = {
      start: vi.fn(async () => undefined),
      cancel: vi.fn(async () => undefined),
      close: vi.fn(async () => undefined),
      subscribe: (listener) => ((emit = listener), () => undefined),
    }
    const client = { open: vi.fn(() => remote), close: vi.fn(async () => undefined) }
    const onText = vi.fn(),
      onDone = vi.fn(),
      onPlan = vi.fn()
    const runtime = new EnhancedAgentRuntime(client)
    const telemetry = { host: vi.fn(), component: vi.fn() }
    const session = runtime.createSession({
      ...options({ stream: vi.fn() } as never, {
        onText,
        onDone,
        onPresentationPlan: onPlan,
      }),
      telemetry,
    })
    expect(session.run('edit')).toBe(true)
    emit({ type: 'plan', steps: ['read', 'write'], requiresConfirmation: true })
    emit({ type: 'text', text: 'done' })
    emit({ type: 'done', result: { text: 'done', cancelled: false, turnLimit: false } })
    expect(onPlan).toHaveBeenCalledOnce()
    expect(onText).toHaveBeenCalledWith('done')
    expect(onDone).toHaveBeenCalledOnce()
    expect(session.snapshot.status).toBe('done')
    expect(client.open).toHaveBeenCalledOnce()
    expect(telemetry.host).toHaveBeenCalledWith('docs', 'dispatch', 'succeeded')
    expect(telemetry.host).toHaveBeenCalledWith('docs', 'plan', 'succeeded')
    expect(telemetry.host).toHaveBeenCalledWith('docs', 'complete', 'succeeded')
  })

  it('isolates replacement generations and fails Enhanced without replay', async () => {
    let emit!: (event: any) => void
    let reject!: () => void
    const remote: EnhancedRuntimeClientSession = {
      start: vi.fn(() => new Promise<void>((_, fail) => (reject = () => fail(new Error('crash'))))),
      cancel: vi.fn(async () => undefined),
      close: vi.fn(async () => undefined),
      subscribe: (listener) => ((emit = listener), () => undefined),
    }
    const runtime = new EnhancedAgentRuntime({ open: () => remote, close: async () => undefined })
    const error = vi.fn()
    const session = runtime.createSession(options({ stream: vi.fn() } as never, { onError: error }))
    session.run('first')
    session.reset()
    reject()
    await Promise.resolve()
    expect(error).not.toHaveBeenCalled()
    emit({ type: 'done', result: { text: 'stale', cancelled: false, turnLimit: false } })
    expect(session.messages).toEqual([])
  })
})
