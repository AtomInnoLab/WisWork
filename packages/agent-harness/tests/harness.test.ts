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
  it('forwards presentation lifecycle events and receipt localization', async () => {
    const transport = manualTransport()
    const onPresentationPlan = vi.fn()
    const onPresentationReceipt = vi.fn(() => 'Localized receipt')
    const harness = createAgentHarness({
      ...options(transport, { onPresentationPlan, onPresentationReceipt }),
      skill: {
        ...skill,
        presentation: {
          prepare: () => ({
            kind: 'ready',
            contract: {
              version: 1,
              taskId: 'task-1',
              documentToken: 'doc-1',
              sessionToken: 'session-1',
              baseRevision: `sha256:${'a'.repeat(64)}`,
              affectedSlides: [2],
              referenceSlides: [],
              checks: [
                {
                  id: 'check-1',
                  kind: 'element_property',
                  slide: 2,
                  roleOrTarget: { kind: 'role', role: 'title' },
                  property: 'color',
                  expected: '#112233',
                },
              ],
              maxCorrectionPasses: 2,
            },
            plan: ['Edit'],
            requiresConfirmation: false,
          }),
          complete: () => ({
            kind: 'receipt',
            receipt: {
              version: 1,
              taskId: 'task-1',
              status: 'unchanged',
              mutationReceiptIds: [],
              passedCheckIds: ['check-1'],
              failedCheckIds: [],
              unavailableCheckIds: [],
              correctionPasses: 0,
              affectedSlides: [2],
            },
          }),
        },
      },
    })

    expect(harness.run('first')).toBe(true)
    await flush()
    transport.callbacks[0]!.onDone()
    await flush()
    await flush()

    expect(onPresentationPlan).toHaveBeenCalledWith({
      steps: ['Edit'],
      requiresConfirmation: false,
    })
    expect(onPresentationReceipt).toHaveBeenCalledOnce()
    expect(harness.messages.at(-1)).toEqual({ role: 'assistant', text: 'Localized receipt' })
  })

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

  it.each(['reset', 'dispose'] as const)(
    'does not launch work when a running listener calls %s',
    (action) => {
      const transport = manualTransport()
      const executeTool = vi.fn(() => ({ output: 'ok', summary: 'done' }))
      const harness = createAgentHarness({
        ...options(transport),
        skill: { ...skill, executeTool },
      })
      harness.subscribe(() => {
        if (harness.snapshot.status === 'running') harness[action]()
      })

      expect(harness.run('do not launch')).toBe(false)

      expect(transport.callbacks).toHaveLength(0)
      expect(executeTool).not.toHaveBeenCalled()
    },
  )

  it('cancels a pending launch when a running listener calls stop', () => {
    const order: string[] = []
    const transport = manualTransport()
    const executeTool = vi.fn(() => ({ output: 'ok', summary: 'done' }))
    const harness = createAgentHarness({
      ...options(transport, {
        onDone: (result) => order.push(`host:${result.cancelled}:${harness.snapshot.status}`),
      }),
      skill: { ...skill, executeTool },
    })
    harness.subscribe(() => {
      order.push(`listener:${harness.snapshot.status}`)
      if (harness.snapshot.status === 'running') harness.stop()
    })

    expect(harness.run('do not launch')).toBe(false)

    expect(harness.snapshot).toEqual({ status: 'cancelled', busy: false, generation: 1 })
    expect(transport.callbacks).toHaveLength(0)
    expect(executeTool).not.toHaveBeenCalled()
    expect(order).toEqual(['listener:running', 'host:true:running', 'listener:cancelled'])
  })

  it('rejects a reentrant run from the synchronous running notification', async () => {
    const transport = manualTransport()
    const harness = createAgentHarness(options(transport))
    const nested = vi.fn<(result: boolean) => void>()
    harness.subscribe(() => {
      if (harness.snapshot.status === 'running') nested(harness.run('nested'))
    })

    expect(harness.run('outer')).toBe(true)
    expect(nested).toHaveBeenCalledWith(false)
    await flush()
    expect(transport.callbacks).toHaveLength(1)
    expect(harness.messages.filter((message) => message.role === 'user')).toHaveLength(1)
  })

  it.each(['buildContext', 'formatUserMessage'] as const)(
    'recovers when %s throws synchronously and allows a later run',
    async (failurePoint) => {
      const transport = manualTransport()
      let shouldThrow = true
      const harness = createAgentHarness({
        ...options(transport),
        skill: {
          ...skill,
          buildContext: () => {
            if (failurePoint === 'buildContext' && shouldThrow) throw new Error('launch_failed')
            return 'context'
          },
        },
        formatUserMessage: (instruction, context) => {
          if (failurePoint === 'formatUserMessage' && shouldThrow) throw new Error('launch_failed')
          return `${instruction}:${context}`
        },
      })

      expect(harness.run('first')).toBe(true)
      expect(harness.snapshot).toEqual({
        status: 'error',
        busy: false,
        generation: 1,
        error: 'launch_failed',
      })

      shouldThrow = false
      expect(harness.run('second')).toBe(true)
      await flush()
      expect(transport.callbacks).toHaveLength(1)
    },
  )

  it('recovers from an asynchronous launch failure and allows a later run', async () => {
    const transport = manualTransport()
    const stream = transport.stream.bind(transport)
    let shouldThrow = true
    transport.stream = (request, callbacks) => {
      if (shouldThrow) throw new Error('stream_failed')
      return stream(request, callbacks)
    }
    const harness = createAgentHarness(options(transport))

    expect(harness.run('first')).toBe(true)
    await flush()
    expect(harness.snapshot).toEqual({
      status: 'error',
      busy: false,
      generation: 1,
      error: 'stream_failed',
    })

    shouldThrow = false
    expect(harness.run('second')).toBe(true)
    await flush()
    expect(transport.callbacks).toHaveLength(1)
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
