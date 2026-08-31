import { describe, expect, it, vi } from 'vitest'
import {
  AgentLoop,
  COMPLETED_VIA_TOOLS_TEXT,
  composeSkills,
  type AgentMessage,
  type AgentSkill,
  type AgentStreamCallbacks,
  type AgentToolCall,
  type AgentTransport,
  type FinalResponseReviewContext,
  type PresentationTaskHooks,
  type ToolExecution,
  type ToolExecutionOutcome,
  suspendToolExecution,
} from '../src'

/** transport scripted turn by turn; exposes the callbacks for manual driving */
function scriptedTransport(script: Array<(cb: AgentStreamCallbacks) => void>): AgentTransport & {
  requests: Array<{ messageCount: number; toolCount: number }>
  cancels: number
} {
  let turn = 0
  const transport = {
    requests: [] as Array<{ messageCount: number; toolCount: number }>,
    cancels: 0,
    lastCallbacks: null as AgentStreamCallbacks | null,
    stream(request: { messages: AgentMessage[]; tools: unknown[] }, cb: AgentStreamCallbacks) {
      transport.requests.push({
        messageCount: request.messages.length,
        toolCount: request.tools.length,
      })
      transport.lastCallbacks = cb
      const step = script[turn++]
      if (step) queueMicrotask(() => step(cb))
      return {
        cancel: () => {
          transport.cancels++
          queueMicrotask(() => cb.onDone())
        },
      }
    },
  }
  return transport
}

function makeSkill(execute?: (call: AgentToolCall) => ToolExecutionOutcome): AgentSkill {
  return {
    id: 'test',
    systemPrompt: 'system',
    tools: [{ name: 'do_thing', description: 'd', inputSchema: { type: 'object' } }],
    buildContext: () => 'CTX',
    executeTool: execute ?? (() => ({ output: 'ok', summary: 'done', mutated: true })),
  }
}

const flush = () => new Promise((r) => setTimeout(r, 0))

const contract = {
  version: 1 as const,
  taskId: 'task-1',
  documentToken: 'doc-1',
  sessionToken: 'session-1',
  baseRevision: `sha256:${'a'.repeat(64)}`,
  affectedSlides: [2],
  referenceSlides: [],
  checks: [
    {
      id: 'check-1',
      kind: 'element_property' as const,
      slide: 2,
      roleOrTarget: { kind: 'role' as const, role: 'title' as const },
      property: 'color' as const,
      expected: '#112233',
    },
  ],
  maxCorrectionPasses: 2 as const,
}

const receipt = (status: 'verified' | 'applied_unverified' = 'verified') => ({
  version: 1 as const,
  taskId: 'task-1',
  status,
  mutationReceiptIds: ['mutation-1'],
  passedCheckIds: status === 'verified' ? ['check-1'] : [],
  failedCheckIds: [],
  unavailableCheckIds: status === 'applied_unverified' ? ['check-1'] : [],
  correctionPasses: 0,
  affectedSlides: [2],
  ...(status === 'applied_unverified' ? { safeCode: 'screenshot_unavailable' as const } : {}),
})

const unchangedReceipt = () => ({
  version: 1 as const,
  taskId: 'task-1',
  status: 'unchanged' as const,
  mutationReceiptIds: [],
  passedCheckIds: ['check-1'],
  failedCheckIds: [],
  unavailableCheckIds: [],
  correctionPasses: 0,
  affectedSlides: [2],
})

const failedReceipt = () => ({
  version: 1 as const,
  taskId: 'task-1',
  status: 'failed' as const,
  mutationReceiptIds: [],
  passedCheckIds: [],
  failedCheckIds: ['check-1'],
  unavailableCheckIds: [],
  correctionPasses: 0,
  affectedSlides: [2],
  safeCode: 'mutation_failed' as const,
})

describe('AgentLoop', () => {
  describe('presentation task orchestration', () => {
    it('enrolls exact tool calls before the first dispatch and closes through a receipt', async () => {
      let dispatched = false
      const enroll = vi.fn(async (calls: readonly AgentToolCall[]) => {
        expect(dispatched).toBe(false)
        expect(calls).toHaveLength(1)
        return { kind: 'ready' as const, contract }
      })
      const transport = scriptedTransport([
        (cb) => {
          cb.onToolCall({ id: 'call-1', name: 'do_thing', input: { slideIndex: 1 } })
          cb.onDone()
        },
        (cb) => cb.onDone(),
      ])
      const done = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(() => {
            dispatched = true
            return { output: 'ok', summary: 'done', mutated: true }
          }),
          presentation: {
            prepare: () => ({ kind: 'bypass' }),
            enroll,
            complete: () => ({ kind: 'receipt', receipt: receipt() }),
          },
        },
        events: { onDone: done },
      })
      loop.run('edit')
      await flush()
      await flush()
      expect(enroll).toHaveBeenCalledOnce()
      expect(dispatched).toBe(true)
      expect(done).toHaveBeenCalledWith(
        expect.objectContaining({ presentation: expect.objectContaining({ status: 'verified' }) }),
      )
    })

    it('accepts authoritative host correction passes without double-counting model turns', async () => {
      const hostCorrected = { ...receipt(), correctionPasses: 2 }
      const done = vi.fn()
      const loop = new AgentLoop({
        transport: scriptedTransport([
          (cb) => {
            cb.onToolCall({ id: 'call-1', name: 'do_thing', input: { slideIndex: 1 } })
            cb.onDone()
          },
          (cb) => cb.onDone(),
        ]),
        skill: {
          ...makeSkill(() => ({ output: 'ok', summary: 'done', mutated: true })),
          presentation: {
            prepare: () => ({ kind: 'bypass' }),
            enroll: () => ({ kind: 'ready', contract }),
            complete: () => ({ kind: 'receipt', receipt: hostCorrected }),
          },
        },
        events: { onDone: done },
      })
      loop.run('edit')
      await flush()
      await flush()
      expect(done).toHaveBeenCalledWith(
        expect.objectContaining({ presentation: expect.objectContaining({ correctionPasses: 2 }) }),
      )
    })

    it('dispatches zero tools when authoritative enrollment fails', async () => {
      const execute = vi.fn(() => ({ output: 'ok', summary: 'done', mutated: true }))
      const transport = scriptedTransport([
        (cb) => {
          cb.onToolCall({ id: 'call-1', name: 'do_thing', input: { slideIndex: 1 } })
          cb.onDone()
        },
      ])
      const error = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(execute),
          presentation: {
            prepare: () => ({ kind: 'bypass' }),
            enroll: async () => {
              throw new Error('stale')
            },
            complete: () => ({ kind: 'receipt', receipt: receipt() }),
          },
        },
        events: { onError: error },
      })
      loop.run('edit')
      await flush()
      expect(execute).not.toHaveBeenCalled()
      expect(error).toHaveBeenCalledWith('presentation_enrollment_unavailable')
    })
    it('lets a simple presentation task bypass clarification and planning UI', async () => {
      const transport = scriptedTransport([
        (cb) => {
          cb.onDelta('No edit needed')
          cb.onDone()
        },
      ])
      const hooks: PresentationTaskHooks = {
        prepare: vi.fn(() => ({ kind: 'bypass' as const })),
        complete: vi.fn(),
      }
      const onPresentationClarify = vi.fn()
      const onPresentationPlan = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: { ...makeSkill(), presentation: hooks },
        events: { onPresentationClarify, onPresentationPlan },
      })

      loop.run('make title blue')
      await flush()

      expect(onPresentationClarify).not.toHaveBeenCalled()
      expect(onPresentationPlan).not.toHaveBeenCalled()
      expect(transport.requests).toHaveLength(1)
    })

    it('asks one bounded clarification and does not dispatch tools or the provider', async () => {
      const transport = scriptedTransport([])
      const executeTool = vi.fn()
      const onPresentationClarify = vi.fn()
      const onDone = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(executeTool),
          presentation: {
            prepare: () => ({ kind: 'clarify', question: 'Which slide is the reference?' }),
            complete: vi.fn(),
          },
        },
        events: { onPresentationClarify, onDone },
      })

      loop.run('make these consistent')
      await flush()

      expect(onPresentationClarify).toHaveBeenCalledOnce()
      expect(transport.requests).toHaveLength(0)
      expect(executeTool).not.toHaveBeenCalled()
      expect(onDone).toHaveBeenCalledWith(expect.objectContaining({ cancelled: false }))
    })

    it('emits a host-authored plan and rejects confirmation before any dispatch', async () => {
      const transport = scriptedTransport([])
      const executeTool = vi.fn()
      const onPresentationPlan = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(executeTool),
          presentation: {
            prepare: () => ({
              kind: 'ready',
              plan: ['Update slides 2–3', 'Verify rendered output'],
              requiresConfirmation: true,
              contract,
            }),
            confirm: vi.fn(() => false),
            complete: vi.fn(),
          },
        },
        events: { onPresentationPlan },
      })

      loop.run('update and verify slides 2–3')
      await flush()

      expect(onPresentationPlan).toHaveBeenCalledWith({
        steps: ['Update slides 2–3', 'Verify rendered output'],
        requiresConfirmation: true,
      })
      expect(transport.requests).toHaveLength(0)
      expect(executeTool).not.toHaveBeenCalled()
    })

    it('uses a valid contract-bound receipt as terminal response facts', async () => {
      const transport = scriptedTransport([
        (cb) => {
          cb.onToolCall({ id: 'write', name: 'do_thing', input: {} })
          cb.onDone()
        },
        (cb) => {
          cb.onDelta('Anything the model says')
          cb.onDone()
        },
      ])
      const onDone = vi.fn()
      const onText = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(),
          presentation: {
            prepare: () => ({ kind: 'ready', contract }),
            complete: () => ({ kind: 'receipt', receipt: receipt() }),
          },
        },
        events: { onDone, onText },
      })

      loop.run('edit slide 2')
      await flush()
      await flush()

      expect(onDone).toHaveBeenCalledWith(
        expect.objectContaining({
          text: 'presentation:verified;slides=2;passed=1;failed=0;unavailable=0;corrections=0;rollback=false',
          presentation: expect.objectContaining({ status: 'verified', affectedSlides: [2] }),
        }),
      )
      expect(loop.messages.at(-1)).toEqual({
        role: 'assistant',
        text: 'presentation:verified;slides=2;passed=1;failed=0;unavailable=0;corrections=0;rollback=false',
      })
    })

    it('reconciles a tool-free model success claim into authoritative unchanged truth', async () => {
      const complete = vi.fn(() => ({ kind: 'receipt' as const, receipt: unchangedReceipt() }))
      const transport = scriptedTransport([
        (cb) => {
          cb.onDelta('Everything was successfully changed!')
          cb.onDone()
        },
      ])
      const onDone = vi.fn()
      const onText = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(),
          presentation: { prepare: () => ({ kind: 'ready', contract }), complete },
        },
        events: { onDone, onText },
      })

      loop.run('make title blue')
      await flush()

      expect(complete).toHaveBeenCalledWith(expect.objectContaining({ mutated: false }))
      expect(onDone).toHaveBeenCalledWith(
        expect.objectContaining({
          text: 'presentation:unchanged;slides=2;passed=1;failed=0;unavailable=0;corrections=0;rollback=false',
          presentation: expect.objectContaining({ status: 'unchanged' }),
        }),
      )
      expect(JSON.stringify(loop.messages)).not.toContain('successfully changed')
      expect(onText).toHaveBeenLastCalledWith(
        'presentation:unchanged;slides=2;passed=1;failed=0;unavailable=0;corrections=0;rollback=false',
      )
    })

    it('reconciles an all-mutated-false tool round into authoritative failed truth', async () => {
      const complete = vi.fn(() => ({ kind: 'receipt' as const, receipt: failedReceipt() }))
      const transport = scriptedTransport([
        (cb) => {
          cb.onToolCall({ id: 'write', name: 'do_thing', input: {} })
          cb.onDone()
        },
        (cb) => {
          cb.onDelta('The edit succeeded')
          cb.onDone()
        },
      ])
      const onDone = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(() => ({ output: 'not applied', summary: 'failed', mutated: false })),
          presentation: { prepare: () => ({ kind: 'ready', contract }), complete },
        },
        events: { onDone },
      })

      loop.run('edit')
      await flush()
      await flush()

      expect(complete).toHaveBeenCalledWith(expect.objectContaining({ mutated: false }))
      expect(onDone).toHaveBeenCalledWith(
        expect.objectContaining({
          text: 'presentation:failed;slides=2;passed=0;failed=1;unavailable=0;corrections=0;rollback=false;code=mutation_failed',
          presentation: expect.objectContaining({ status: 'failed' }),
        }),
      )
    })

    it('never retries applied-unverified and does not accept terminal success without a valid receipt', async () => {
      const complete = vi.fn(() => ({
        kind: 'receipt' as const,
        receipt: receipt('applied_unverified'),
      }))
      const transport = scriptedTransport([
        (cb) => {
          cb.onToolCall({ id: 'write', name: 'do_thing', input: {} })
          cb.onDone()
        },
        (cb) => {
          cb.onDelta('done')
          cb.onDone()
        },
      ])
      const onDone = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(),
          presentation: { prepare: () => ({ kind: 'ready', contract }), complete },
        },
        events: { onDone },
      })

      loop.run('edit slide 2')
      await flush()
      await flush()

      expect(complete).toHaveBeenCalledOnce()
      expect(transport.requests).toHaveLength(2)
      expect(onDone).toHaveBeenCalledWith(
        expect.objectContaining({
          presentation: expect.objectContaining({ status: 'applied_unverified' }),
        }),
      )

      const invalidDone = vi.fn()
      const invalidLoop = new AgentLoop({
        transport: scriptedTransport([
          (cb) => {
            cb.onToolCall({ id: 'write', name: 'do_thing', input: {} })
            cb.onDone()
          },
          (cb) => {
            cb.onDelta('verified')
            cb.onDone()
          },
        ]),
        skill: {
          ...makeSkill(),
          presentation: {
            prepare: () => ({ kind: 'ready', contract }),
            complete: () => ({
              kind: 'receipt',
              receipt: { ...receipt(), mutationReceiptIds: [] },
            }),
          },
        },
        events: { onDone: invalidDone },
      })
      invalidLoop.run('edit')
      await flush()
      await flush()
      expect(invalidDone).not.toHaveBeenCalledWith(
        expect.objectContaining({ presentation: expect.objectContaining({ status: 'verified' }) }),
      )
    })

    it('runs at most contract-bounded corrective turns and cancellation preserves dispatch truth', async () => {
      const complete = vi
        .fn()
        .mockReturnValueOnce({ kind: 'correct', instruction: 'Correct check-1 only' })
        .mockReturnValueOnce({
          kind: 'receipt',
          receipt: {
            ...receipt(),
            correctionPasses: 1,
            mutationReceiptIds: ['mutation-1', 'mutation-2'],
          },
        })
      const transport = scriptedTransport([
        (cb) => {
          cb.onToolCall({ id: 'write', name: 'do_thing', input: {} })
          cb.onDone()
        },
        (cb) => {
          cb.onDelta('premature')
          cb.onDone()
        },
        (cb) => {
          cb.onToolCall({ id: 'fix', name: 'do_thing', input: {} })
          cb.onDone()
        },
        (cb) => {
          cb.onDelta('fixed')
          cb.onDone()
        },
      ])
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(),
          presentation: { prepare: () => ({ kind: 'ready', contract }), complete },
        },
      })
      loop.run('edit')
      await flush()
      await flush()
      await flush()
      await flush()
      expect(complete).toHaveBeenCalledTimes(2)
      expect(transport.requests).toHaveLength(4)
      expect(loop.messages.map((message) => message.role)).toEqual([
        'user',
        'assistant',
        'tool',
        'assistant',
        'user',
        'assistant',
        'tool',
        'assistant',
      ])
    })

    it('cancels before dispatch with zero-mutation receipt truth', async () => {
      const complete = vi.fn(() => ({ kind: 'receipt' as const, receipt: unchangedReceipt() }))
      const transport = scriptedTransport([() => {}])
      const onDone = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(),
          presentation: { prepare: () => ({ kind: 'ready', contract }), complete },
        },
        events: { onDone },
      })

      loop.run('edit')
      await flush()
      loop.cancel()
      await flush()

      expect(complete).toHaveBeenCalledWith(
        expect.objectContaining({ mutated: false, cancelled: true }),
      )
      expect(onDone).toHaveBeenCalledWith(
        expect.objectContaining({
          cancelled: true,
          presentation: expect.objectContaining({ status: 'unchanged' }),
        }),
      )
      expect(loop.messages.map((message) => message.role)).toEqual(['user', 'assistant'])
    })

    it('reconciles a cancellation after dispatch and reports the applied receipt truth', async () => {
      const complete = vi.fn(() => ({
        kind: 'receipt' as const,
        receipt: receipt('applied_unverified'),
      }))
      const transport = scriptedTransport([
        (cb) => {
          cb.onToolCall({ id: 'write', name: 'do_thing', input: {} })
          cb.onDone()
        },
        () => {},
      ])
      const onDone = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(),
          presentation: { prepare: () => ({ kind: 'ready', contract }), complete },
        },
        events: { onDone },
      })

      loop.run('edit')
      await flush()
      await flush()
      loop.cancel()
      await flush()

      expect(complete).toHaveBeenCalledWith(
        expect.objectContaining({ mutated: true, cancelled: true }),
      )
      expect(onDone).toHaveBeenCalledWith(
        expect.objectContaining({
          cancelled: true,
          presentation: expect.objectContaining({ status: 'applied_unverified' }),
        }),
      )
      expect(loop.messages.map((message) => message.role)).toEqual([
        'user',
        'assistant',
        'tool',
        'assistant',
      ])
    })

    it('does not abort an async completion reconciliation when cancel races after dispatch', async () => {
      let settle!: () => void
      let reconciliationAborted = false
      const complete = vi.fn(
        ({ signal }: { signal?: AbortSignal }) =>
          new Promise<{ kind: 'receipt'; receipt: ReturnType<typeof receipt> }>((resolve) => {
            settle = () => resolve({ kind: 'receipt', receipt: receipt('applied_unverified') })
            signal?.addEventListener('abort', () => {
              reconciliationAborted = true
            })
          }),
      )
      const callbacks: AgentStreamCallbacks[] = []
      const transport: AgentTransport = {
        stream: (_request, cb) => {
          callbacks.push(cb)
          return { cancel: () => queueMicrotask(() => cb.onDone()) }
        },
      }
      const onDone = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(),
          presentation: { prepare: () => ({ kind: 'ready', contract }), complete },
        },
        events: { onDone },
      })

      loop.run('edit')
      await flush()
      callbacks[0]!.onToolCall({ id: 'write', name: 'do_thing', input: {} })
      callbacks[0]!.onDone()
      await flush()
      callbacks[1]!.onDelta('done')
      callbacks[1]!.onDone()
      await flush()
      loop.cancel()
      settle()
      await flush()

      expect(complete.mock.calls[0]![0].signal?.aborted).toBe(false)
      expect(reconciliationAborted).toBe(false)
      expect(onDone).toHaveBeenCalledWith(
        expect.objectContaining({
          presentation: expect.objectContaining({ status: 'applied_unverified' }),
        }),
      )
    })

    it('ignores a pending completion rejection after reset without polluting state or events', async () => {
      let rejectCompletion!: (error: Error) => void
      const complete = vi.fn(
        () =>
          new Promise<never>((_resolve, reject) => {
            rejectCompletion = reject
          }),
      )
      const callbacks: AgentStreamCallbacks[] = []
      const transport: AgentTransport = {
        stream: (_request, cb) => {
          callbacks.push(cb)
          return { cancel: () => queueMicrotask(() => cb.onDone()) }
        },
      }
      const onDone = vi.fn()
      const onError = vi.fn()
      const onText = vi.fn()
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(),
          presentation: { prepare: () => ({ kind: 'ready', contract }), complete },
        },
        events: { onDone, onError, onText },
      })

      loop.run('edit')
      await flush()
      callbacks[0]!.onToolCall({ id: 'write', name: 'do_thing', input: {} })
      callbacks[0]!.onDone()
      await flush()
      callbacks[1]!.onDelta('success')
      callbacks[1]!.onDone()
      await flush()
      expect(complete).toHaveBeenCalledOnce()

      loop.reset()
      onText.mockClear()
      rejectCompletion(new Error('late reconciliation failure'))
      await flush()

      expect(loop.busy).toBe(false)
      expect(loop.messages).toEqual([])
      expect(onDone).not.toHaveBeenCalled()
      expect(onError).not.toHaveBeenCalled()
      expect(onText).not.toHaveBeenCalled()
    })

    it('replaces streamed model success with authoritative error text before reporting failure', async () => {
      const transport = scriptedTransport([
        (cb) => {
          cb.onDelta('Successfully changed everything')
          cb.onDone()
        },
      ])
      const calls: string[] = []
      const onText = vi.fn((text: string) => calls.push(`text:${text}`))
      const onError = vi.fn((error: string) => calls.push(`error:${error}`))
      const loop = new AgentLoop({
        transport,
        skill: {
          ...makeSkill(),
          presentation: {
            prepare: () => ({ kind: 'ready', contract }),
            complete: () => ({ kind: 'receipt', receipt: { invalid: true } as never }),
          },
        },
        events: { onText, onError },
      })

      loop.run('edit')
      await flush()

      expect(calls).toEqual([
        'text:Successfully changed everything',
        'text:presentation:error;code=presentation_receipt_invalid',
        'error:presentation_receipt_invalid',
      ])
      expect(loop.messages.at(-1)).toEqual({
        role: 'assistant',
        text: 'presentation:error;code=presentation_receipt_invalid',
      })
    })
  })

  it('reviews one normal tool-free completion and retries without finishing the UI turn', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('I cannot make that supported edit.')
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('Corrected response')
        cb.onDone()
      },
    ])
    const reviewFinalResponse = vi.fn(() => '[System] Use the supported editing tools now.')
    const skill: AgentSkill = { ...makeSkill(), reviewFinalResponse }
    const onDone = vi.fn()
    const onTurnEnd = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone, onTurnEnd } })

    loop.run('make the edit')
    await flush()
    await flush()

    expect(reviewFinalResponse).toHaveBeenCalledTimes(1)
    expect(reviewFinalResponse).toHaveBeenCalledWith({
      text: 'I cannot make that supported edit.',
      mutated: false,
    })
    expect(transport.requests).toHaveLength(2)
    expect(onDone).toHaveBeenCalledTimes(1)
    expect(onDone).toHaveBeenCalledWith({
      text: 'Corrected response',
      cancelled: false,
      turnLimit: false,
    })
    expect(onTurnEnd).not.toHaveBeenCalled()
    expect(loop.messages).toEqual([
      { role: 'user', text: 'make the edit\n\nCTX' },
      { role: 'assistant', text: 'I cannot make that supported edit.' },
      { role: 'user', text: '[System] Use the supported editing tools now.' },
      { role: 'assistant', text: 'Corrected response' },
    ])
  })

  it('clears rejected prose before a corrective tool turn and only finishes after the final turn', async () => {
    const callbacks: AgentStreamCallbacks[] = []
    const transport: AgentTransport = {
      stream: (_request, cb) => {
        callbacks.push(cb)
        return { cancel: () => cb.onDone() }
      },
    }
    const skill: AgentSkill = { ...makeSkill(), reviewFinalResponse: () => 'use the tool' }
    const onText = vi.fn()
    const onDone = vi.fn()
    const onTurnEnd = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onText, onDone, onTurnEnd } })

    loop.run('edit')
    await flush()
    callbacks[0]!.onDelta('unsupported denial')
    callbacks[0]!.onDone()
    await flush()
    expect(onText).toHaveBeenLastCalledWith('')
    expect(onDone).not.toHaveBeenCalled()
    callbacks[1]!.onToolCall({ id: 'fix', name: 'do_thing', input: {} })
    callbacks[1]!.onDone()
    await flush()
    expect(onDone).not.toHaveBeenCalled()
    callbacks[2]!.onDelta('edit completed')
    callbacks[2]!.onDone()
    await flush()

    expect(onText.mock.calls.map(([text]) => text)).toEqual([
      'unsupported denial',
      '',
      'edit completed',
    ])
    expect(onTurnEnd).toHaveBeenCalledTimes(1)
    expect(onDone).toHaveBeenCalledTimes(1)
    expect(onDone).toHaveBeenCalledWith({
      text: 'edit completed',
      cancelled: false,
      turnLimit: false,
    })
  })

  it('keeps rejected prose cleared when a corrective tool flow is cancelled', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('unsupported denial')
        cb.onDone()
      },
      (cb) => {
        cb.onToolCall({ id: 'fix', name: 'do_thing', input: {} })
        cb.onDone()
      },
      () => {
        // wait for explicit cancellation after the corrective tool executes
      },
    ])
    const skill: AgentSkill = { ...makeSkill(), reviewFinalResponse: () => 'use the tool' }
    const onText = vi.fn()
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onText, onDone } })

    loop.run('edit')
    await flush()
    await flush()
    expect(onText).toHaveBeenLastCalledWith('')
    expect(onDone).not.toHaveBeenCalled()
    loop.cancel()
    await flush()

    expect(onText).toHaveBeenLastCalledWith('')
    expect(onDone).toHaveBeenCalledWith({ text: '', cancelled: true, turnLimit: false })
  })

  it('bounds terminal text passed to completion review while retaining both ends', async () => {
    const longResponse = `BEGIN-${'x'.repeat(6_000)}-END`
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta(longResponse)
        cb.onDone()
      },
    ])
    const reviewFinalResponse = vi.fn((_context: FinalResponseReviewContext) => undefined)
    const skill: AgentSkill = { ...makeSkill(), reviewFinalResponse }
    const loop = new AgentLoop({ transport, skill })

    loop.run('question')
    await flush()

    const reviewed = reviewFinalResponse.mock.calls[0]![0].text
    expect(reviewed.length).toBeLessThanOrEqual(4_096)
    expect(reviewed).toMatch(/^BEGIN-/)
    expect(reviewed).toMatch(/-END$/)
    expect(reviewed).toContain('[response truncated for review]')
  })

  it.each([
    ['non-string', () => 42 as unknown as string],
    ['empty', () => '   '],
    ['too many characters', () => 'x'.repeat(2_001)],
    ['too many UTF-8 bytes', () => '修'.repeat(1_500)],
  ])('fails open for an invalid %s completion correction', async (_label, review) => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('final answer')
        cb.onDone()
      },
    ])
    const skill: AgentSkill = { ...makeSkill(), reviewFinalResponse: review }
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })

    loop.run('question')
    await flush()

    expect(transport.requests).toHaveLength(1)
    expect(loop.messages).toEqual([
      { role: 'user', text: 'question\n\nCTX' },
      { role: 'assistant', text: 'final answer' },
    ])
    expect(onDone).toHaveBeenCalledWith({
      text: 'final answer',
      cancelled: false,
      turnLimit: false,
    })
  })

  it('fails open when credential sanitization expands a correction past its bounds', async () => {
    const expandingCorrection = Array.from({ length: 180 }, () => 'x://a:b@').join(' ')
    expect(expandingCorrection.length).toBeLessThanOrEqual(2_000)
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('final answer')
        cb.onDone()
      },
    ])
    const skill: AgentSkill = {
      ...makeSkill(),
      reviewFinalResponse: () => expandingCorrection,
    }
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })

    loop.run('question')
    await flush()

    expect(transport.requests).toHaveLength(1)
    expect(loop.messages).toEqual([
      { role: 'user', text: 'question\n\nCTX' },
      { role: 'assistant', text: 'final answer' },
    ])
    expect(onDone).toHaveBeenCalledWith({
      text: 'final answer',
      cancelled: false,
      turnLimit: false,
    })
  })

  it('sanitizes a valid correction before storing it in model history', async () => {
    const secret = `sk-${'a'.repeat(20)}`
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('denial')
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('corrected')
        cb.onDone()
      },
    ])
    const skill: AgentSkill = {
      ...makeSkill(),
      reviewFinalResponse: () => `Retry without ${secret}`,
    }
    const loop = new AgentLoop({ transport, skill })

    loop.run('edit')
    await flush()
    await flush()

    expect(loop.messages[2]).toEqual({
      role: 'user',
      text: 'Retry without [REDACTED_API_KEY]',
    })
  })

  it('allows at most one completion-review retry per run', async () => {
    const respond = (text: string) => (cb: AgentStreamCallbacks) => {
      cb.onDelta(text)
      cb.onDone()
    }
    const transport = scriptedTransport([respond('first denial'), respond('second denial')])
    const reviewFinalResponse = vi.fn(() => 'try again')
    const skill: AgentSkill = { ...makeSkill(), reviewFinalResponse }
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })

    loop.run('edit')
    await flush()
    await flush()

    expect(reviewFinalResponse).toHaveBeenCalledTimes(1)
    expect(transport.requests).toHaveLength(2)
    expect(onDone).toHaveBeenCalledWith({
      text: 'second denial',
      cancelled: false,
      turnLimit: false,
    })
  })

  it('fails open when the completion-review hook throws', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('ordinary final response')
        cb.onDone()
      },
    ])
    const skill: AgentSkill = {
      ...makeSkill(),
      reviewFinalResponse: () => {
        throw new Error('private policy failure')
      },
    }
    const onDone = vi.fn()
    const onError = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone, onError } })

    loop.run('question')
    await flush()

    expect(onDone).toHaveBeenCalledWith({
      text: 'ordinary final response',
      cancelled: false,
      turnLimit: false,
    })
    expect(onError).not.toHaveBeenCalled()
  })

  it('does not review a cancelled completion', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('partial')
      },
    ])
    const reviewFinalResponse = vi.fn(() => 'try again')
    const skill: AgentSkill = { ...makeSkill(), reviewFinalResponse }
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })

    loop.run('edit')
    await flush()
    loop.cancel()
    await flush()

    expect(reviewFinalResponse).not.toHaveBeenCalled()
    expect(transport.requests).toHaveLength(1)
    expect(onDone).toHaveBeenCalledWith({ text: 'partial', cancelled: true, turnLimit: false })
  })

  it('keeps provider history paired when the corrective turn is cancelled', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('denial')
        cb.onDone()
      },
      (cb) => cb.onDelta('correcting'),
    ])
    const skill: AgentSkill = { ...makeSkill(), reviewFinalResponse: () => 'retry with tools' }
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })

    loop.run('edit')
    await flush()
    loop.cancel()
    await flush()

    expect(loop.messages.map((message) => message.role)).toEqual([
      'user',
      'assistant',
      'user',
      'assistant',
    ])
    expect(onDone).toHaveBeenCalledWith({
      text: 'correcting',
      cancelled: true,
      turnLimit: false,
    })
  })

  it('rolls the whole run back when the corrective request fails', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('denial')
        cb.onDone()
      },
      (cb) => cb.onError('network dropped'),
    ])
    const skill: AgentSkill = { ...makeSkill(), reviewFinalResponse: () => 'retry with tools' }
    const onError = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onError } })
    loop.restore([
      { role: 'user', text: 'earlier question' },
      { role: 'assistant', text: 'earlier answer' },
    ])

    loop.run('edit')
    await flush()
    await flush()

    expect(onError).toHaveBeenCalledWith('network dropped')
    expect(loop.messages).toEqual([
      { role: 'user', text: 'earlier question' },
      { role: 'assistant', text: 'earlier answer' },
    ])
  })

  it('assigns invocation nonces per run while deduplicating a retried transport callback', async () => {
    const seen: AgentToolCall[] = []
    const transport = scriptedTransport([
      (cb) => {
        const repeated = { id: 'same', name: 'do_thing', input: { value: 1 } }
        cb.onToolCall(repeated)
        cb.onToolCall({ ...repeated })
        cb.onDone()
      },
      (cb) => cb.onDone(),
      (cb) => {
        cb.onToolCall({ id: 'same', name: 'do_thing', input: { value: 1 } })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const loop = new AgentLoop({
      transport,
      skill: makeSkill((call) => {
        seen.push(call)
        return { output: 'ok', summary: 'ok', mutated: false }
      }),
    })
    loop.run('first')
    await flush()
    await flush()
    loop.run('second')
    await flush()
    await flush()
    expect(seen).toHaveLength(3)
    expect(seen[0]!.invocationId).toBeTruthy()
    expect(seen[1]!.invocationId).toBe(seen[0]!.invocationId)
    expect(seen[2]!.invocationId).not.toBe(seen[0]!.invocationId)
  })

  it('runs a plain-text turn to completion', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('Hello')
        cb.onDelta(', world')
        cb.onDone()
      },
    ])
    const onDone = vi.fn()
    const onText = vi.fn()
    const loop = new AgentLoop({ transport, skill: makeSkill(), events: { onDone, onText } })
    loop.run('question')
    await flush()
    expect(onText).toHaveBeenLastCalledWith('Hello, world')
    expect(onDone).toHaveBeenCalledWith({
      text: 'Hello, world',
      cancelled: false,
      turnLimit: false,
    })
    expect(loop.busy).toBe(false)
    // user message carries the skill context
    expect(loop.messages[0]).toEqual({ role: 'user', text: 'question\n\nCTX' })
    expect(loop.messages[1]).toEqual({ role: 'assistant', text: 'Hello, world' })
  })

  it('attaches images to the user message (and omits the field without any)', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('I can see it')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({ transport, skill: makeSkill() })
    const images = [{ base64: 'AAAA', mime: 'image/png' }]
    loop.run('describe the image', images)
    await flush()
    expect(loop.messages[0]).toEqual({ role: 'user', text: 'describe the image\n\nCTX', images })
    expect('images' in (loop.messages[0] as object)).toBe(true)
  })

  it('executes tools and loops back to the model', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('Let me edit the document first')
        cb.onToolCall({ id: 't1', name: 'do_thing', input: { a: 1 } })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('All done')
        cb.onDone()
      },
    ])
    const executed: AgentToolCall[] = []
    const skill = makeSkill((call) => {
      executed.push(call)
      return { output: 'result-1', summary: 'changed 1 spot', mutated: true }
    })
    const onToolExecuted = vi.fn()
    const onDone = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill,
      captureSnapshot: () => 'SNAP',
      events: { onToolExecuted, onDone },
    })
    loop.run('make a change')
    await flush()
    await flush()

    expect(executed).toHaveLength(1)
    // first mutation carries the pre-tool snapshot
    expect(onToolExecuted).toHaveBeenCalledWith(expect.objectContaining({ snapshotBefore: 'SNAP' }))
    expect(onDone).toHaveBeenCalledWith({ text: 'All done', cancelled: false, turnLimit: false })
    // history: user / assistant+tools / tool results / assistant
    expect(loop.messages.map((m) => m.role)).toEqual(['user', 'assistant', 'tool', 'assistant'])
    const toolMsg = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMsg.results[0]).toEqual({
      id: 't1',
      name: 'do_thing',
      output: 'result-1',
      isError: undefined,
    })
    // second request included the tool round-trip
    expect(transport.requests[1].messageCount).toBe(3)
  })

  it('settles the run when snapshot capture throws before a tool', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
    ])
    const onError = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      captureSnapshot: () => {
        throw new Error('snapshot_failed')
      },
      events: { onError },
    })

    loop.run('make a change')
    await flush()

    expect(onError).toHaveBeenCalledWith('snapshot_failed')
    expect(loop.busy).toBe(false)
    expect(loop.messages).toEqual([])
  })

  it('settles the run when starting the next provider turn throws', async () => {
    let turn = 0
    const transport: AgentTransport = {
      stream(_request, callbacks) {
        turn++
        if (turn === 2) throw new Error('second_turn_stream_failed')
        queueMicrotask(() => {
          callbacks.onToolCall({ id: 't1', name: 'do_thing', input: {} })
          callbacks.onDone()
        })
        return { cancel() {} }
      },
    }
    const onError = vi.fn()
    const loop = new AgentLoop({ transport, skill: makeSkill(), events: { onError } })

    loop.run('make a change')
    await flush()
    await flush()

    expect(onError).toHaveBeenCalledWith('second_turn_stream_failed')
    expect(loop.busy).toBe(false)
    expect(loop.messages).toEqual([])
  })

  it('pauses a tool round without starting the next provider turn and resumes it once', async () => {
    let resolveExecution!: (execution: ToolExecution) => void
    const result = new Promise<ToolExecution>((resolve) => {
      resolveExecution = resolve
    })
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 'approval-1', name: 'do_thing', input: { text: 'approved text' } })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('Applied')
        cb.onDone()
      },
    ])
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => suspendToolExecution(result)),
      events: { onToolExecuted },
    })

    loop.run('write it')
    await flush()
    expect(loop.busy).toBe(true)
    expect(transport.requests).toHaveLength(1)
    expect(onToolExecuted).not.toHaveBeenCalled()

    resolveExecution({ output: 'written', summary: 'wrote text', mutated: true })
    resolveExecution({ output: 'duplicate', summary: 'must be ignored', mutated: true })
    await flush()
    await flush()
    expect(transport.requests).toHaveLength(2)
    expect(onToolExecuted).toHaveBeenCalledTimes(1)
    expect(loop.messages.map((message) => message.role)).toEqual([
      'user',
      'assistant',
      'tool',
      'assistant',
    ])
    expect((loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>).results).toEqual([
      {
        id: 'approval-1',
        name: 'do_thing',
        output: 'written',
        isError: undefined,
      },
    ])
  })

  it('cancel releases a suspended tool immediately and ignores its late resolution', async () => {
    let resolveExecution!: (execution: ToolExecution) => void
    const result = new Promise<ToolExecution>((resolve) => {
      resolveExecution = resolve
    })
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 'approval-1', name: 'do_thing', input: {} })
        cb.onDone()
      },
    ])
    const onDone = vi.fn()
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => suspendToolExecution(result)),
      events: { onDone, onToolExecuted },
    })

    loop.run('write it')
    await flush()
    loop.cancel()
    await flush()
    expect(loop.busy).toBe(false)
    expect(onDone).toHaveBeenCalledWith({ text: '', cancelled: true, turnLimit: false })
    expect(transport.requests).toHaveLength(1)

    resolveExecution({ output: 'too late', summary: 'late', mutated: true })
    await flush()
    expect(onToolExecuted).not.toHaveBeenCalled()
    expect(transport.requests).toHaveLength(1)
  })

  it('reset releases a suspended tool and prevents its late result from restoring history', async () => {
    let resolveExecution!: (execution: ToolExecution) => void
    const result = new Promise<ToolExecution>((resolve) => {
      resolveExecution = resolve
    })
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 'approval-1', name: 'do_thing', input: {} })
        cb.onDone()
      },
    ])
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => suspendToolExecution(result)),
      events: { onToolExecuted },
    })

    loop.run('write it')
    await flush()
    loop.reset()
    await flush()
    expect(loop.busy).toBe(false)
    expect(loop.messages).toEqual([])

    resolveExecution({ output: 'too late', summary: 'late', mutated: true })
    await flush()
    expect(loop.messages).toEqual([])
    expect(onToolExecuted).not.toHaveBeenCalled()
  })

  it('fails closed for malformed and rejected suspensions without duplicating execution', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 'approval-1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() =>
        suspendToolExecution(Promise.reject(new Error('proposal disappeared'))),
      ),
      events: { onToolExecuted },
    })

    loop.run('write it')
    await flush()
    await flush()
    expect(onToolExecuted).toHaveBeenCalledTimes(1)
    expect(onToolExecuted).toHaveBeenCalledWith(
      expect.objectContaining({
        execution: expect.objectContaining({ output: 'invalid_tool_output', isError: true }),
      }),
    )
    expect(transport.requests).toHaveLength(2)
  })

  it('turns a malformed suspended final result into one stable tool error', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 'approval-1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() =>
        suspendToolExecution(Promise.resolve(undefined as unknown as ToolExecution)),
      ),
      events: { onToolExecuted },
    })

    loop.run('write it')
    await flush()
    await flush()
    expect(onToolExecuted).toHaveBeenCalledTimes(1)
    expect(onToolExecuted.mock.calls[0]?.[0].execution).toMatchObject({
      output: 'invalid_tool_output',
      isError: true,
    })
  })

  it('fails closed when a suspended Promise has a throwing then getter', async () => {
    const target = Promise.resolve({ output: 'must not escape', summary: 'bad' })
    const hostileResult = new Proxy(target, {
      get(value, property, receiver) {
        if (property === 'then') throw new Error('hostile then getter')
        return Reflect.get(value, property, receiver)
      },
    })
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 'approval-1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('Recovered')
        cb.onDone()
      },
    ])
    const onDone = vi.fn()
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => suspendToolExecution(hostileResult)),
      events: { onDone, onToolExecuted },
    })

    loop.run('write it')
    await flush()
    await flush()
    expect(onToolExecuted).toHaveBeenCalledTimes(1)
    expect(onToolExecuted.mock.calls[0]?.[0].execution).toMatchObject({
      output: 'invalid_tool_output',
      isError: true,
    })
    expect(onDone).toHaveBeenCalledWith({ text: 'Recovered', cancelled: false, turnLimit: false })
    expect(loop.busy).toBe(false)
  })

  it('stops the current tool batch after a controlling execution and resumes the provider once', async () => {
    let resolveExecution!: (execution: ToolExecution) => void
    const result = new Promise<ToolExecution>((resolve) => {
      resolveExecution = resolve
    })
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 'write-1', name: 'do_thing', input: { text: 'first' } })
        cb.onToolCall({ id: 'write-2', name: 'do_thing', input: { text: 'second' } })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('Stopped after rejection')
        cb.onDone()
      },
    ])
    const executeTool = vi.fn((call: AgentToolCall) =>
      call.id === 'write-1'
        ? suspendToolExecution(result)
        : { output: 'must not execute', summary: 'bad', mutated: true },
    )
    const onToolStart = vi.fn()
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(executeTool),
      events: { onToolStart, onToolExecuted },
    })

    loop.run('make two writes')
    await flush()
    resolveExecution({
      output: 'user_rejected_change',
      summary: 'Change rejected',
      isError: true,
      stopToolBatch: true,
    })
    await flush()
    await flush()

    expect(executeTool).toHaveBeenCalledTimes(1)
    expect(onToolStart).toHaveBeenCalledTimes(1)
    expect(onToolExecuted).toHaveBeenCalledTimes(1)
    expect(transport.requests).toHaveLength(2)
    const toolMessage = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMessage.results).toHaveLength(2)
    expect(toolMessage.results[0]).toMatchObject({
      id: 'write-1',
      output: 'user_rejected_change',
      isError: true,
    })
    expect(toolMessage.results[1]).toMatchObject({
      id: 'write-2',
      isError: true,
    })
    expect(toolMessage.results[1]?.output).toContain('tool batch')
  })

  it('fails closed when stopToolBatch is not a boolean', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 'write-1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(
        () =>
          ({
            output: 'bad control',
            summary: 'bad',
            stopToolBatch: 'yes',
          }) as unknown as ToolExecution,
      ),
      events: { onToolExecuted },
    })

    loop.run('write')
    await flush()
    await flush()
    expect(onToolExecuted).toHaveBeenCalledWith(
      expect.objectContaining({
        execution: expect.objectContaining({ output: 'invalid_tool_output', isError: true }),
      }),
    )
  })

  it('preserves model-visible tool image content in history for the next request', async () => {
    const image = { base64: 'iVBORw0KGgo=', mime: 'image/png' }
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('seen')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => ({
        output: 'captured',
        summary: 'image',
        modelContent: [{ type: 'image', image }],
      })),
    })
    loop.run('capture')
    await flush()
    await flush()
    expect(loop.messages).toContainEqual({
      role: 'tool',
      results: [
        {
          id: 't1',
          name: 'do_thing',
          output: 'captured',
          isError: undefined,
          content: [{ type: 'image', image }],
        },
      ],
    })
  })

  it('turns invalid model-visible content into a stable tool error and keeps running', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('recovered')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => ({
        output: 'captured',
        summary: 'image',
        modelContent: [
          { type: 'image', image: { base64: 'R0lGODlhAQABAAAAACw=', mime: 'image/gif' } },
        ],
      })),
    })

    loop.run('capture')
    await flush()
    await flush()

    expect(loop.busy).toBe(false)
    expect(loop.messages).toContainEqual({
      role: 'tool',
      results: [
        {
          id: 't1',
          name: 'do_thing',
          output: 'invalid_tool_output',
          isError: true,
        },
      ],
    })
    expect(loop.messages.at(-1)).toEqual({ role: 'assistant', text: 'recovered' })
  })

  it('emits onToolStart before each execution, paired with onToolExecuted', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: { a: 1 } })
        cb.onToolCall({ id: 't2', name: 'do_thing', input: { a: 2 } })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const order: string[] = []
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => {
        order.push('exec')
        return { output: 'ok', summary: 's' }
      }),
      events: {
        onToolStart: (call) => order.push(`start:${call.id}`),
        onToolExecuted: ({ call }) => order.push(`done:${call.id}`),
      },
    })
    loop.run('x')
    await flush()
    await flush()
    expect(order).toEqual(['start:t1', 'exec', 'done:t1', 'start:t2', 'exec', 'done:t2'])
  })

  it('only the first mutating tool carries a snapshot', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onToolCall({ id: 't2', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    const onToolExecuted = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      captureSnapshot: () => 'SNAP',
      events: { onToolExecuted },
    })
    loop.run('x')
    await flush()
    await flush()
    expect(onToolExecuted).toHaveBeenCalledTimes(2)
    expect(onToolExecuted.mock.calls[0][0].snapshotBefore).toBe('SNAP')
    expect(onToolExecuted.mock.calls[1][0].snapshotBefore).toBeUndefined()
  })

  it('after maxTurns, adds a final tool-less turn that yields a partial answer', async () => {
    const alwaysTool = (cb: AgentStreamCallbacks) => {
      cb.onToolCall({ id: 'x', name: 'do_thing', input: {} })
      cb.onDone()
    }
    const finalize = (cb: AgentStreamCallbacks) => {
      cb.onDelta('partial conclusion')
      cb.onDone()
    }
    const transport = scriptedTransport([alwaysTool, alwaysTool, finalize])
    const onDone = vi.fn()
    const reviewFinalResponse = vi.fn(() => 'try again')
    const loop = new AgentLoop({
      transport,
      skill: { ...makeSkill(), reviewFinalResponse },
      maxTurns: 2,
      events: { onDone },
    })
    loop.run('x')
    await flush()
    await flush()
    await flush()
    await flush()
    // The third turn is the finalizing one: no tools, and a system note was inserted into history
    expect(transport.requests).toHaveLength(3)
    expect(transport.requests[2].toolCount).toBe(0)
    const note = loop.messages.find((m) => m.role === 'user' && m.text.includes('turn limit'))
    expect(note).toBeDefined()
    expect(onDone).toHaveBeenCalledWith({
      text: 'partial conclusion',
      cancelled: false,
      turnLimit: true,
    })
    expect(reviewFinalResponse).not.toHaveBeenCalled()
  })

  it('stops an unchanged tool-call loop without imposing a total turn limit', async () => {
    const repeat = (cb: AgentStreamCallbacks) => {
      cb.onToolCall({ id: crypto.randomUUID(), name: 'do_thing', input: { page: 1 } })
      cb.onDone()
    }
    const transport = scriptedTransport(Array.from({ length: 6 }, () => repeat))
    const execute = vi.fn(() => ({ output: 'ok', summary: 'done', mutated: true }))
    const onError = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(execute),
      events: { onError },
    })

    loop.run('x')
    for (let i = 0; i < 12; i++) await flush()

    expect(execute).toHaveBeenCalledTimes(4)
    expect(onError).toHaveBeenCalledWith('tool_loop_detected')
    expect(loop.busy).toBe(false)
  })

  it('cancel drops pending tool calls and finalizes the run', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('partial')
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        // no onDone: waits for cancel
      },
    ])
    const executed: AgentToolCall[] = []
    const skill = makeSkill((call) => {
      executed.push(call)
      return { output: 'ok', summary: 's' }
    })
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })
    loop.run('x')
    await flush()
    loop.cancel()
    await flush()
    expect(transport.cancels).toBe(1)
    expect(executed).toHaveLength(0)
    expect(onDone).toHaveBeenCalledWith({ text: 'partial', cancelled: true, turnLimit: false })
    // assistant message stored without toolCalls (no results would follow)
    expect(loop.messages[1]).toEqual({ role: 'assistant', text: 'partial' })
  })

  it('cancel during tool execution aborts the signal, skips remaining tools and starts no new turn', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onToolCall({ id: 't2', name: 'do_thing', input: {} })
        cb.onDone()
      },
      // The second turn should never be requested (no return to the model after cancel)
      (cb) => cb.onDone(),
    ])
    const seen: Array<{ id: string; abortedAfterCancel: boolean }> = []
    let loop: AgentLoop | null = null
    const skill = makeSkill()
    skill.executeTool = (call, signal) => {
      // Simulate the user hitting stop while a long tool is running
      loop?.cancel()
      seen.push({ id: call.id, abortedAfterCancel: signal?.aborted === true })
      return { output: 'ok', summary: 's', mutated: true }
    }
    const onDone = vi.fn()
    loop = new AgentLoop({ transport, skill, events: { onDone } })
    loop.run('x')
    await flush()
    await flush()
    // Only the first tool ran; after cancel the signal is aborted right away (long tools' inner loops break on it)
    expect(seen).toEqual([{ id: 't1', abortedAfterCancel: true }])
    // The second tool didn't run, but a paired error result was added (tool_use/tool_result stay paired)
    const toolMsg = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMsg.results.map((r) => r.id)).toEqual(['t1', 't2'])
    expect(toolMsg.results[1].isError).toBe(true)
    // No further model request; the run finishes as cancelled
    expect(transport.requests).toHaveLength(1)
    expect(onDone).toHaveBeenCalledTimes(1)
    expect(onDone).toHaveBeenCalledWith({ text: '', cancelled: true, turnLimit: false })
    expect(loop.busy).toBe(false)
  })

  it('executeTool receives a live (non-aborted) signal during normal runs', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onDone(),
    ])
    let sawSignal: AbortSignal | undefined
    const skill = makeSkill()
    skill.executeTool = (_call, signal) => {
      sawSignal = signal
      return { output: 'ok', summary: 's' }
    }
    const loop = new AgentLoop({ transport, skill })
    loop.run('x')
    await flush()
    await flush()
    expect(sawSignal).toBeInstanceOf(AbortSignal)
    expect(sawSignal?.aborted).toBe(false)
  })

  it('trims history at user boundaries', async () => {
    const script = Array.from({ length: 4 }, () => (cb: AgentStreamCallbacks) => {
      cb.onDelta('answer')
      cb.onDone()
    })
    const transport = scriptedTransport(script)
    const loop = new AgentLoop({ transport, skill: makeSkill(), maxHistory: 3 })
    for (let i = 0; i < 4; i++) {
      loop.run(`question${i}`)
      await flush()
    }
    // trimmed to start at a user message
    expect(loop.messages.length).toBeLessThanOrEqual(4)
    expect(loop.messages[0].role).toBe('user')
  })

  it('lets an unbounded Cowork run continue past the former default turn cap', async () => {
    // 21 tool turns → 1 user + 21×(assistant+tool) = 43 messages > maxHistory 40
    const script: Array<(cb: AgentStreamCallbacks) => void> = Array.from(
      { length: 21 },
      (_, i) => (cb: AgentStreamCallbacks) => {
        cb.onToolCall({ id: `t${i}`, name: 'do_thing', input: { page: i + 1 } })
        cb.onDone()
      },
    )
    script.push((cb) => {
      cb.onDelta('all done')
      cb.onDone()
    })
    const transport = scriptedTransport(script)
    const onDone = vi.fn()
    const onError = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      compaction: false,
      events: { onDone, onError },
    })
    loop.run('big job')
    for (let i = 0; i < 50; i++) await flush()
    expect(transport.requests).toHaveLength(22)
    // every request carried the full history including the run's user message
    expect(transport.requests.every((r) => r.messageCount > 0)).toBe(true)
    expect(loop.messages).toHaveLength(44)
    expect(loop.messages[0]).toMatchObject({ role: 'user' })
    expect(onError).not.toHaveBeenCalled()
    expect(onDone).toHaveBeenCalledWith({ text: 'all done', cancelled: false, turnLimit: false })
  })

  it('boundary trim is abandoned when the window holds no user message (never empties history)', async () => {
    const script: Array<(cb: AgentStreamCallbacks) => void> = Array.from(
      { length: 4 },
      (_, i) => (cb: AgentStreamCallbacks) => {
        cb.onToolCall({ id: `t${i}`, name: 'do_thing', input: {} })
        cb.onDone()
      },
    )
    script.push((cb) => {
      cb.onDelta('done')
      cb.onDone()
    })
    script.push((cb) => {
      cb.onDelta('second answer')
      cb.onDone()
    })
    const transport = scriptedTransport(script)
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      maxTurns: 10,
      maxHistory: 5,
      compaction: false,
    })
    loop.run('long first job') // ends with 10 messages, the only user message at index 0
    for (let i = 0; i < 20; i++) await flush()
    expect(loop.messages).toHaveLength(10)
    loop.run('follow-up')
    await flush()
    // last-5 window has no user boundary → trim abandoned, nothing lost
    expect(loop.messages[0]).toMatchObject({
      role: 'user',
      text: expect.stringContaining('long first job'),
    })
    expect(transport.requests.at(-1)!.messageCount).toBe(11)
  })

  it('restore seeds history and the next run sends it to the model', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('picking up from before')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({ transport, skill: makeSkill() })
    loop.restore([
      { role: 'user', text: 'previous question' },
      { role: 'assistant', text: 'previous answer' },
    ])
    expect(loop.messages).toHaveLength(2)

    loop.run('follow-up')
    await flush()
    // The request carries the 2 restored messages + the new user message
    expect(transport.requests[0].messageCount).toBe(3)
    expect(loop.messages[0]).toEqual({ role: 'user', text: 'previous question' })
  })

  it('restore is a no-op when history exists or the loop is busy', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('answer')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({ transport, skill: makeSkill() })
    loop.run('opening message')
    await flush()
    const before = loop.messages.length
    loop.restore([{ role: 'user', text: 'should-not-be-injected' }])
    expect(loop.messages.length).toBe(before)
    expect(
      loop.messages.some(
        (m) => m.role === 'user' && 'text' in m && m.text === 'should-not-be-injected',
      ),
    ).toBe(false)
  })

  it('a failed run rolls its user message back out of history', async () => {
    const transport = scriptedTransport([
      (cb) => cb.onError('Not signed in'),
      (cb) => {
        cb.onDelta('answer to the second question')
        cb.onDone()
      },
    ])
    const onError = vi.fn()
    const loop = new AgentLoop({ transport, skill: makeSkill(), events: { onError } })
    loop.run('change all headings to red')
    await flush()
    expect(onError).toHaveBeenCalledWith('Not signed in')
    // the failed instruction is gone, so it can't be re-executed by the next run
    expect(loop.messages).toHaveLength(0)

    loop.run('what does this report propose?')
    await flush()
    // the next request carries only the new user message — no adjacent user turns
    expect(transport.requests[1].messageCount).toBe(1)
    expect(loop.messages.map((m) => m.role)).toEqual(['user', 'assistant'])
  })

  it('a mid-run failure after tool turns rolls back the whole run', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => cb.onError('network dropped'),
    ])
    const onError = vi.fn()
    const loop = new AgentLoop({ transport, skill: makeSkill(), events: { onError } })
    loop.restore([
      { role: 'user', text: 'earlier question' },
      { role: 'assistant', text: 'earlier answer' },
    ])
    loop.run('do more work')
    await flush()
    await flush()
    expect(onError).toHaveBeenCalledWith('network dropped')
    // history is back to the pre-run state: no dangling user/assistant/tool from the failed run
    expect(loop.messages).toEqual([
      { role: 'user', text: 'earlier question' },
      { role: 'assistant', text: 'earlier answer' },
    ])
  })

  it('restore drops unanswered user messages (trailing and adjacent)', () => {
    const transport = scriptedTransport([])
    const loop = new AgentLoop({ transport, skill: makeSkill() })
    loop.restore([
      { role: 'user', text: 'failed and never answered' },
      { role: 'user', text: 'answered question' },
      { role: 'assistant', text: 'the answer' },
      { role: 'user', text: 'trailing unanswered' },
    ])
    expect(loop.messages).toEqual([
      { role: 'user', text: 'answered question' },
      { role: 'assistant', text: 'the answer' },
    ])
  })

  it('restore keeps edits-only turns: empty assistant text gets a placeholder instead of dropping the pair', () => {
    const transport = scriptedTransport([])
    const loop = new AgentLoop({ transport, skill: makeSkill() })
    loop.restore([
      { role: 'user', text: 'translate the intro' },
      { role: 'assistant', text: '' }, // edits-only run persisted without a summary
      { role: 'user', text: 'now shorten it' },
      { role: 'assistant', text: 'shortened' },
    ])
    expect(loop.messages).toHaveLength(4)
    expect(loop.messages[0]).toEqual({ role: 'user', text: 'translate the intro' })
    const placeholder = loop.messages[1] as { role: string; text: string }
    expect(placeholder.role).toBe('assistant')
    expect(placeholder.text).not.toBe('') // providers reject empty assistant content blocks
  })

  it('after tools mutate, an empty final model turn still stores non-empty history for follow-ups', async () => {
    // Regression for upstream#12 / #22: first AI prompt mutates via tools with no
    // prose, second prompt must not inherit an empty assistant content block.
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: { a: 1 } })
        cb.onDone()
      },
      (cb) => cb.onDone(), // model returns no text after tools
      (cb) => {
        cb.onDelta('second prompt ok')
        cb.onDone()
      },
    ])
    const onDone = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      events: { onDone },
    })
    loop.run('first change')
    await flush()
    await flush()

    // onDone reports the raw (empty) turn text so app UIs keep their own
    // localized fallbacks; only the history entry gets the placeholder.
    expect(onDone).toHaveBeenCalledWith({
      text: '',
      cancelled: false,
      turnLimit: false,
    })
    const afterFirst = loop.messages
    expect(afterFirst.map((m) => m.role)).toEqual(['user', 'assistant', 'tool', 'assistant'])
    const finalAssistant = afterFirst[3] as Extract<AgentMessage, { role: 'assistant' }>
    expect(finalAssistant.text).toBe(COMPLETED_VIA_TOOLS_TEXT)
    expect(finalAssistant.text.length).toBeGreaterThan(0)

    onDone.mockClear()
    loop.run('follow-up')
    await flush()
    expect(onDone).toHaveBeenCalledWith({
      text: 'second prompt ok',
      cancelled: false,
      turnLimit: false,
    })
    // Follow-up request carries the prior (now non-empty) terminal assistant
    expect(transport.requests[2]!.messageCount).toBeGreaterThanOrEqual(5)
  })

  it('restore trims oversized history at a user boundary', () => {
    const transport = scriptedTransport([])
    const loop = new AgentLoop({ transport, skill: makeSkill(), maxHistory: 2 })
    loop.restore([
      { role: 'user', text: 'q1' },
      { role: 'assistant', text: 'a1' },
      { role: 'user', text: 'q2' },
      { role: 'assistant', text: 'a2' },
    ])
    expect(loop.messages.length).toBeLessThanOrEqual(2)
    expect(loop.messages[0]).toEqual({ role: 'user', text: 'q2' })
  })
})

describe('AgentLoop compaction', () => {
  it('over budget before run: folds old conversation via LLM summary; new request carries the summary, not the originals', async () => {
    const bigAnswer = 'old reply'.padEnd(800, 'y')
    const transport = scriptedTransport([
      // Two real turns first: turn 1 produces a long over-budget reply
      (cb) => {
        cb.onDelta(bigAnswer)
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('second-turn reply')
        cb.onDone()
      },
      // The 3rd run triggers compaction: this stream call is the summary request
      (cb) => {
        cb.onDelta('Goal: build a deck; 5 slides done')
        cb.onDone()
      },
      // Only then comes the actual model turn for run 3
      (cb) => {
        cb.onDelta('continuing')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      compaction: { maxBytes: 500, keepRecentBytes: 100 },
    })
    loop.run('old instruction')
    await flush()
    loop.run('second question')
    await flush()
    loop.run('follow-up')
    await flush()
    expect(transport.requests).toHaveLength(4)
    const msgs = loop.messages
    expect(msgs[0]!.role).toBe('user')
    expect((msgs[0] as { text: string }).text).toContain('[Summary of earlier conversation')
    expect((msgs[0] as { text: string }).text).toContain('5 slides done')
    // The folded original text is gone from history
    expect(msgs.some((m) => 'text' in m && m.text === bigAnswer)).toBe(false)
    // The new user message comes after the summary
    expect(msgs.some((m) => m.role === 'user' && 'text' in m && m.text.includes('follow-up'))).toBe(
      true,
    )
  })

  it('falls back to a mechanical digest when the LLM summary fails; the turn proceeds normally', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('done'.padEnd(800, 'y'))
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('second turn')
        cb.onDone()
      },
      (cb) => cb.onError('summary failed'),
      (cb) => {
        cb.onDelta('answer')
        cb.onDone()
      },
    ])
    const onDone = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      compaction: { maxBytes: 500, keepRecentBytes: 100 },
      events: { onDone },
    })
    loop.run('key instruction')
    await flush()
    loop.run('second question')
    await flush()
    loop.run('continue')
    await flush()
    expect(onDone).toHaveBeenLastCalledWith({ text: 'answer', cancelled: false, turnLimit: false })
    // The mechanical digest kept the gist of the user instruction
    expect((loop.messages[0] as { text: string }).text).toContain('key instruction')
    expect((loop.messages[0] as { text: string }).text).toContain(
      '[Summary of earlier conversation',
    )
  })

  it('restore over budget folds mechanically (no LLM request)', () => {
    const transport = scriptedTransport([])
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      compaction: { maxBytes: 500, keepRecentBytes: 200 },
    })
    loop.restore([
      { role: 'user', text: 'very old instruction'.padEnd(600, 'x') },
      { role: 'assistant', text: 'early reply' },
      { role: 'user', text: 'recent question' },
      { role: 'assistant', text: 'recent answer' },
    ])
    expect(transport.requests).toHaveLength(0)
    expect((loop.messages[0] as { text: string }).text).toContain(
      '[Summary of earlier conversation',
    )
    expect((loop.messages[0] as { text: string }).text).toContain('very old instruction')
    expect(loop.messages.some((m) => 'text' in m && m.text === 'recent question')).toBe(true)
  })

  it('over budget mid-run truncates stale tool outputs (the 2 most recent keep the full text)', async () => {
    const big = 'z'.repeat(3_000)
    const script = Array.from({ length: 3 }, (_, i) => (cb: AgentStreamCallbacks) => {
      cb.onToolCall({ id: `t${i}`, name: 'do_thing', input: {} })
      cb.onDone()
    })
    script.push((cb) => {
      cb.onDelta('done')
      cb.onDone()
    })
    const transport = scriptedTransport(script)
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(() => ({ output: big, summary: 'big output', mutated: false })),
      compaction: { maxBytes: 5_000, keepRecentBytes: 2_000, disableLlmSummary: true },
    })
    loop.run('do the work')
    await flush()
    await flush()
    await flush()
    await flush()
    const toolMsgs = loop.messages.filter((m) => m.role === 'tool')
    expect(toolMsgs).toHaveLength(3)
    const first = toolMsgs[0] as { results: Array<{ output: string }> }
    const last = toolMsgs[2] as { results: Array<{ output: string }> }
    expect(first.results[0]!.output).toContain('…(output truncated: too long)')
    expect(first.results[0]!.output.length).toBeLessThan(1_200)
    expect(last.results[0]!.output).toBe(big)
  })

  it('counts and drops stale tool media before the next provider request', async () => {
    const requests: AgentMessage[][] = []
    let turn = 0
    const transport: AgentTransport = {
      stream(request, callbacks) {
        requests.push(request.messages)
        const current = turn++
        queueMicrotask(() => {
          if (current < 2) callbacks.onToolCall({ id: `t${current}`, name: 'do_thing', input: {} })
          else callbacks.onDelta('done')
          callbacks.onDone()
        })
        return { cancel: vi.fn() }
      },
    }
    const images = ['A'.repeat(400), 'B'.repeat(400)]
    const loop = new AgentLoop({
      transport,
      skill: makeSkill((call) => ({
        output: 'captured',
        summary: 'image',
        modelContent: [
          {
            type: 'image',
            image: { base64: images[Number(call.id.slice(1))]!, mime: 'image/png' },
          },
        ],
      })),
      compaction: { maxBytes: 650, keepRecentBytes: 300, disableLlmSummary: true },
    })

    loop.run('capture twice')
    await flush()
    await flush()
    await flush()

    const thirdRequest = JSON.stringify(requests[2])
    expect(thirdRequest).not.toContain(images[0])
    expect(thirdRequest).toContain(images[1])
    const oldResult = loop.messages.filter((message) => message.role === 'tool').at(0) as Extract<
      AgentMessage,
      { role: 'tool' }
    >
    expect(oldResult.results[0]!.content).toBeUndefined()
  })

  it('compaction: false disables both folding and truncation', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('answer')
        cb.onDone()
      },
    ])
    const loop = new AgentLoop({ transport, skill: makeSkill(), compaction: false })
    const bigText = 'x'.repeat(200_000)
    loop.restore([
      { role: 'user', text: bigText },
      { role: 'assistant', text: 'ok' },
    ])
    loop.run('continue')
    await flush()
    // Only the real model turn, no summary request; the original text was not folded
    expect(transport.requests).toHaveLength(1)
    expect((loop.messages[0] as { text: string }).text).toBe(bigText)
  })

  it('tool executor exceptions become error results and the loop continues', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {} })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('OK')
        cb.onDone()
      },
    ])
    const skill = makeSkill(() => {
      throw new Error('boom')
    })
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })
    loop.run('x')
    await flush()
    await flush()
    const toolMsg = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMsg.results[0].isError).toBe(true)
    expect(toolMsg.results[0].output).toBe('boom')
    expect(onDone).toHaveBeenCalledWith({ text: 'OK', cancelled: false, turnLimit: false })
  })

  it('calls with inputError are not executed; an is_error result is fed back so the model can retry', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({ id: 't1', name: 'do_thing', input: {}, inputError: 'bad json' })
        cb.onDone()
      },
      (cb) => {
        cb.onToolCall({ id: 't2', name: 'do_thing', input: { a: 1 } })
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('done')
        cb.onDone()
      },
    ])
    const executed: AgentToolCall[] = []
    const skill = makeSkill((call) => {
      executed.push(call)
      return { output: 'ok', summary: 'ok' }
    })
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })
    loop.run('x')
    await flush()
    await flush()
    await flush()
    expect(executed.map((c) => c.id)).toEqual(['t2'])
    const toolMsg = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMsg.results[0].isError).toBe(true)
    expect(toolMsg.results[0].output).toContain('bad json')
    expect(onDone).toHaveBeenCalledWith({ text: 'done', cancelled: false, turnLimit: false })
  })

  it('a truncated tool call is fed back as "split the call", not as a JSON error', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onToolCall({
          id: 't1',
          name: 'do_thing',
          input: {},
          inputError: 'Unexpected end of JSON input',
          truncated: true,
        })
        cb.onStopReason?.('max_tokens')
        cb.onDone()
      },
      (cb) => {
        cb.onDelta('done')
        cb.onDone()
      },
    ])
    const executed: AgentToolCall[] = []
    const skill = makeSkill((call) => {
      executed.push(call)
      return { output: 'ok', summary: 'ok' }
    })
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onDone } })
    loop.run('x')
    await flush()
    await flush()
    expect(executed).toHaveLength(0)
    const toolMsg = loop.messages[2] as Extract<AgentMessage, { role: 'tool' }>
    expect(toolMsg.results[0].isError).toBe(true)
    expect(toolMsg.results[0].output).toContain('smaller tool calls')
    expect(toolMsg.results[0].output).not.toContain('JSON failed to parse')
    // the follow-up turn completed normally, and a non-final max_tokens does not mark the result truncated
    expect(onDone).toHaveBeenCalledWith({ text: 'done', cancelled: false, turnLimit: false })
  })

  it('a max_tokens stop on the final text turn surfaces truncated: true in onDone', async () => {
    const transport = scriptedTransport([
      (cb) => {
        cb.onDelta('partial reply cut off mid-')
        cb.onStopReason?.('max_tokens')
        cb.onDone()
      },
    ])
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill: makeSkill(), events: { onDone } })
    loop.run('x')
    await flush()
    expect(onDone).toHaveBeenCalledWith({
      text: 'partial reply cut off mid-',
      cancelled: false,
      turnLimit: false,
      truncated: true,
    })
  })

  it('the parse-failure counter is consecutive: a successful call resets it', async () => {
    const badTurn = (cb: AgentStreamCallbacks) => {
      cb.onToolCall({ id: 'bad', name: 'do_thing', input: {}, inputError: 'bad json' })
      cb.onDone()
    }
    const goodTurn = (cb: AgentStreamCallbacks) => {
      cb.onToolCall({ id: 'good', name: 'do_thing', input: { a: 1 } })
      cb.onDone()
    }
    const finalTurn = (cb: AgentStreamCallbacks) => {
      cb.onDelta('recovered')
      cb.onDone()
    }
    // 2 fails, success, 2 fails: 4 total but never 3 in a row → the run must complete
    const transport = scriptedTransport([badTurn, badTurn, goodTurn, badTurn, badTurn, finalTurn])
    const onError = vi.fn()
    const onDone = vi.fn()
    const loop = new AgentLoop({
      transport,
      skill: makeSkill(),
      maxTurns: 10,
      events: { onError, onDone },
    })
    loop.run('x')
    for (let i = 0; i < 12; i++) await flush()
    expect(onError).not.toHaveBeenCalled()
    expect(onDone).toHaveBeenCalledWith({ text: 'recovered', cancelled: false, turnLimit: false })
  })

  it('terminates the run after consecutive input-parse failures hit the limit', async () => {
    const badTurn = (cb: AgentStreamCallbacks) => {
      cb.onToolCall({ id: 't', name: 'do_thing', input: {}, inputError: 'bad json' })
      cb.onDone()
    }
    const transport = scriptedTransport([badTurn, badTurn, badTurn, badTurn])
    const executed: AgentToolCall[] = []
    const skill = makeSkill((call) => {
      executed.push(call)
      return { output: 'ok', summary: 'ok' }
    })
    const onError = vi.fn()
    const onDone = vi.fn()
    const loop = new AgentLoop({ transport, skill, events: { onError, onDone } })
    loop.run('x')
    for (let i = 0; i < 6; i++) await flush()
    expect(executed).toHaveLength(0)
    expect(transport.requests).toHaveLength(3)
    expect(onError).toHaveBeenCalledWith(expect.stringContaining('retries stopped'))
    expect(onDone).not.toHaveBeenCalled()
    expect(loop.busy).toBe(false)
  })
})

describe('composeSkills', () => {
  it('merges prompts, tools and context, and routes execution', async () => {
    const a: AgentSkill = {
      id: 'a',
      systemPrompt: 'PA',
      tools: [{ name: 'tool_a', description: '', inputSchema: {} }],
      buildContext: () => 'CA',
      executeTool: () => ({ output: 'from-a', summary: 'a' }),
    }
    const b: AgentSkill = {
      id: 'b',
      systemPrompt: 'PB',
      tools: [{ name: 'tool_b', description: '', inputSchema: {} }],
      executeTool: () => ({ output: 'from-b', summary: 'b' }),
    }
    const merged = composeSkills('ab', 'INTRO', [a, b])
    expect(merged.systemPrompt).toBe('INTRO\n\nPA\n\nPB')
    expect(merged.tools.map((t) => t.name)).toEqual(['tool_a', 'tool_b'])
    expect(merged.buildContext?.()).toBe('CA')
    const routed = await merged.executeTool({ id: '1', name: 'tool_b', input: {} })
    expect('output' in routed ? routed.output : undefined).toBe('from-b')
    const unknown = await merged.executeTool({ id: '2', name: 'nope', input: {} })
    expect('isError' in unknown ? unknown.isError : undefined).toBe(true)
  })

  it('rejects duplicate tool names', () => {
    const tool = { name: 'same', description: '', inputSchema: {} }
    const make = (id: string): AgentSkill => ({
      id,
      systemPrompt: '',
      tools: [tool],
      executeTool: () => ({ output: '', summary: '' }),
    })
    expect(() => composeSkills('x', '', [make('a'), make('b')])).toThrow(/duplicate/)
  })

  it('forwards the first child completion review that returns a correction', () => {
    const pass = vi.fn(() => undefined)
    const correct = vi.fn(() => 'static correction')
    const skipped = vi.fn(() => 'later correction')
    const merged = composeSkills('combined', '', [
      { ...makeSkill(), id: 'pass', tools: [], reviewFinalResponse: pass },
      { ...makeSkill(), id: 'correct', tools: [], reviewFinalResponse: correct },
      { ...makeSkill(), id: 'skipped', tools: [], reviewFinalResponse: skipped },
    ])
    const context = { text: 'answer', mutated: false }

    expect(merged.reviewFinalResponse?.(context)).toBe('static correction')
    expect(pass).toHaveBeenCalledWith(context)
    expect(correct).toHaveBeenCalledWith(context)
    expect(skipped).not.toHaveBeenCalled()
  })
})
