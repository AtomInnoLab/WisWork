import { describe, expect, it, vi } from 'vitest'
import type {
  AgentLoopOptions,
  AgentMessage,
  AgentSkill,
  AgentStreamCallbacks,
  AgentTransport,
} from '@wiswork/agent-core'
import { createAgentHarness } from '../src'

interface ManualTransport extends AgentTransport {
  callbacks: AgentStreamCallbacks[]
  cancels: number
}

function manualTransport(): ManualTransport {
  const transport: ManualTransport = {
    callbacks: [],
    cancels: 0,
    stream(_request, callbacks) {
      transport.callbacks.push(callbacks)
      return {
        cancel: () => {
          transport.cancels++
        },
      }
    },
  }
  return transport
}

const skill: AgentSkill = {
  id: 'test',
  systemPrompt: 'system',
  tools: [],
  executeTool: () => ({ output: 'ok', summary: 'done' }),
}

function options(
  transport: AgentTransport,
  events: NonNullable<AgentLoopOptions['events']> = {},
): AgentLoopOptions {
  return { transport, skill, events, compaction: false }
}

const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('createAgentHarness', () => {
  it('publishes running and done state and rejects empty or concurrent runs', async () => {
    const transport = manualTransport()
    const harness = createAgentHarness(options(transport))
    const snapshots: string[] = []
    harness.subscribe(() => snapshots.push(`${harness.snapshot.status}:${harness.snapshot.busy}`))

    expect(harness.run('')).toBe(false)
    expect(harness.run('first')).toBe(true)
    expect(harness.snapshot).toMatchObject({ status: 'running', busy: true, generation: 1 })
    expect(harness.run('second')).toBe(false)

    await flush()
    transport.callbacks[0]!.onDelta('answer')
    transport.callbacks[0]!.onDone()
    await flush()

    expect(harness.snapshot).toEqual({ status: 'done', busy: false, generation: 1 })
    expect(harness.messages.at(-1)).toEqual({ role: 'assistant', text: 'answer' })
    expect(snapshots).toEqual(['running:true', 'done:false'])
  })

  it('stops a run as cancelled while preserving history', async () => {
    const transport = manualTransport()
    const harness = createAgentHarness(options(transport))
    expect(harness.run('keep me')).toBe(true)
    await flush()

    harness.stop()
    expect(transport.cancels).toBe(1)
    transport.callbacks[0]!.onDone()
    await flush()

    expect(harness.snapshot.status).toBe('cancelled')
    expect(harness.snapshot.busy).toBe(false)
    expect(harness.messages[0]).toMatchObject({ role: 'user', text: 'keep me' })
  })

  it('reset clears history and suppresses callbacks from the old generation', async () => {
    const onText = vi.fn()
    const onDone = vi.fn()
    const transport = manualTransport()
    const harness = createAgentHarness(options(transport, { onText, onDone }))
    harness.run('old')
    await flush()
    const oldCallbacks = transport.callbacks[0]!

    harness.reset()
    oldCallbacks.onDelta('late')
    oldCallbacks.onDone()
    await flush()

    expect(harness.snapshot).toEqual({ status: 'idle', busy: false, generation: 2 })
    expect(harness.messages).toEqual([])
    expect(onText).not.toHaveBeenCalled()
    expect(onDone).not.toHaveBeenCalled()
  })

  it('restores messages and publishes the change', () => {
    const harness = createAgentHarness(options(manualTransport()))
    const listener = vi.fn()
    harness.subscribe(listener)
    const restored: AgentMessage[] = [
      { role: 'user', text: 'question' },
      { role: 'assistant', text: 'answer' },
    ]

    harness.restore(restored)

    expect(harness.messages).toEqual(restored)
    expect(listener).toHaveBeenCalledOnce()
  })

  it('invokes host terminal callbacks before publishing the state transition', async () => {
    const order: string[] = []
    const transport = manualTransport()
    const harness = createAgentHarness(
      options(transport, {
        onDone: () => order.push(`host:${harness.snapshot.status}`),
      }),
    )
    harness.subscribe(() => order.push(`listener:${harness.snapshot.status}`))
    harness.run('go')
    await flush()
    order.length = 0

    transport.callbacks[0]!.onDone()
    await flush()

    expect(order).toEqual(['host:running', 'listener:done'])
  })

  it('publishes errors and does not stay busy when a host callback throws', async () => {
    const transport = manualTransport()
    const harness = createAgentHarness(
      options(transport, {
        onError: () => {
          throw new Error('host callback failed')
        },
      }),
    )
    harness.run('go')
    await flush()

    expect(() => transport.callbacks[0]!.onError('model_failed')).not.toThrow()

    expect(harness.snapshot).toEqual({
      status: 'error',
      busy: false,
      generation: 1,
      error: 'model_failed',
    })
  })

  it('dispose is terminal, clears listeners, and suppresses late callbacks', async () => {
    const onDone = vi.fn()
    const transport = manualTransport()
    const harness = createAgentHarness(options(transport, { onDone }))
    const listener = vi.fn()
    harness.subscribe(listener)
    harness.run('go')
    await flush()
    const callbacks = transport.callbacks[0]!
    listener.mockClear()

    harness.dispose()
    callbacks.onDone()
    await flush()

    expect(transport.cancels).toBe(1)
    expect(harness.run('again')).toBe(false)
    expect(onDone).not.toHaveBeenCalled()
    expect(listener).not.toHaveBeenCalled()
  })

  it('keeps two harness instances independent', async () => {
    const firstTransport = manualTransport()
    const secondTransport = manualTransport()
    const first = createAgentHarness(options(firstTransport))
    const second = createAgentHarness(options(secondTransport))
    first.run('first')
    second.run('second')
    await flush()

    first.stop()
    firstTransport.callbacks[0]!.onDone()
    secondTransport.callbacks[0]!.onDelta('ok')
    secondTransport.callbacks[0]!.onDone()
    await flush()

    expect(first.snapshot.status).toBe('cancelled')
    expect(second.snapshot.status).toBe('done')
    expect(second.messages[0]).toMatchObject({ role: 'user', text: 'second' })
    expect(secondTransport.cancels).toBe(0)
  })
})
