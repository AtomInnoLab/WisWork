import { act, createElement, StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import type { AgentStreamCallbacks, AgentTransport } from '@wiswork/agent-core'
import {
  createAgentController,
  beginSlidesHostRun,
  classifySlidesQcFailure,
  completeSlidesHostRun,
  stopSlidesHostRun,
  recordSlidesRunAttachments,
  useAgentControllerCleanup,
} from '../src/renderer/ai/agent-controller'

function manualTransport(): AgentTransport & {
  callbacks: AgentStreamCallbacks[]
  cancels: number
} {
  const transport = {
    callbacks: [] as AgentStreamCallbacks[],
    cancels: 0,
    stream(_request: unknown, callbacks: AgentStreamCallbacks) {
      transport.callbacks.push(callbacks)
      return { cancel: () => transport.cancels++ }
    },
  }
  return transport
}

const skill = {
  id: 'slides-host',
  systemPrompt: 'system',
  tools: [{ name: 'execute_slide_script', description: 'edit', inputSchema: { type: 'object' } }],
  executeTool: vi.fn(() => ({ output: 'ok', summary: 'slide changed', mutated: true })),
}
const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('Slides interactive agent controller', () => {
  it('does not let the preceding startTurn completion settle a follow-up run', async () => {
    let documentId: string | null = null
    let onEvent: ((event: any) => void) | undefined
    let finishFirst!: () => void
    const first = new Promise<void>((resolve) => {
      finishFirst = resolve
    })
    const done = vi.fn()
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi
        .fn()
        .mockReturnValueOnce(first)
        .mockImplementation(() => new Promise<void>(() => {})),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn((listener) => {
        onEvent = listener
        return () => {}
      }),
      onToolCall: vi.fn(() => () => {}),
    }
    const controller = createAgentController(
      { transport: manualTransport(), skill, events: { onDone: done } },
      { host: 'slides', api },
    )
    controller.activate()
    await flush()
    controller.run('first')
    await flush()
    onEvent?.({ type: 'done', result: { text: '', cancelled: false, turnLimit: false } })
    await flush()
    expect(controller.run('follow-up')).toBe(true)
    await flush()
    finishFirst()
    await flush()
    expect(controller.snapshot.busy).toBe(true)
    expect(done).toHaveBeenCalledOnce()
    controller.dispose()
  })
  it('replaces an Enhanced deck registration and rejects callbacks from the old generation', async () => {
    let documentId: string | null = null
    const toolListeners: Array<(request: any) => void> = []
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi.fn(async () => undefined),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn(() => () => undefined),
      onToolCall: vi.fn((listener) => {
        toolListeners.push(listener)
        return () => undefined
      }),
    }
    const executeTool = vi.fn(async () => ({ output: 'ok', summary: 'edited', mutated: true }))
    const controller = createAgentController(
      { transport: manualTransport(), skill: { ...skill, executeTool } },
      { host: 'slides', api },
    )
    controller.activate()
    await flush()
    expect(api.register).toHaveBeenCalledWith(expect.objectContaining({ generation: 0 }))
    const oldListener = toolListeners[0]!
    controller.reset()
    await flush()
    await flush()
    expect(api.unregister).toHaveBeenCalledWith(expect.any(String), 0)
    expect(api.register).toHaveBeenLastCalledWith(expect.objectContaining({ generation: 1 }))
    oldListener({
      documentId,
      generation: 0,
      call: { id: 'stale', name: 'execute_slide_script', input: {} },
    })
    await flush()
    expect(executeTool).not.toHaveBeenCalled()
    expect(api.toolResult).not.toHaveBeenCalled()
    controller.dispose()
  })

  it('keeps Enhanced slide mutations inside the renderer presentation lifecycle', async () => {
    let documentId: string | null = null
    let onToolCall: ((request: any) => void) | undefined
    let onEvent: ((event: any) => void) | undefined
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi.fn(() => new Promise<void>(() => undefined)),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn((listener) => {
        onEvent = listener
        return () => undefined
      }),
      onToolCall: vi.fn((listener) => {
        onToolCall = listener
        return () => undefined
      }),
    }
    const contract: any = {
      version: 1,
      taskId: 'task-1',
      documentToken: 'doc-1',
      sessionToken: 'session-1',
      baseRevision: `sha256:${'a'.repeat(64)}`,
      affectedSlides: [1],
      referenceSlides: [],
      checks: [
        {
          id: 'check-1',
          kind: 'element_property',
          slide: 1,
          roleOrTarget: { kind: 'role', role: 'title' },
          property: 'color',
          expected: '#112233',
        },
      ],
      maxCorrectionPasses: 2,
    }
    const prepare = vi.fn(() => ({
      kind: 'ready' as const,
      contract,
      plan: ['Apply bounded deck edits'],
      requiresConfirmation: true,
    }))
    const confirm = vi.fn(async () => true)
    const enroll = vi.fn(() => ({ kind: 'ready' as const, contract }))
    const complete = vi.fn(() => ({
      kind: 'receipt' as const,
      receipt: {
        version: 1,
        taskId: 'task-1',
        status: 'verified' as const,
        mutationReceiptIds: ['mutation-1'],
        passedCheckIds: ['check-1'],
        failedCheckIds: [],
        unavailableCheckIds: [],
        correctionPasses: 0,
        affectedSlides: [1],
      },
    }))
    const enhancedSkill: any = {
      ...skill,
      presentation: { prepare, confirm, enroll, complete },
    }
    const captureSnapshot = vi.fn(() => 'deck-before')
    const done = vi.fn()
    const controller = createAgentController(
      {
        transport: manualTransport(),
        skill: enhancedSkill,
        captureSnapshot,
        events: { onDone: done },
      },
      { host: 'slides', api },
    )
    controller.activate()
    await flush()
    expect(controller.run('edit safely')).toBe(true)
    await flush()
    expect(prepare).toHaveBeenCalledOnce()
    expect(confirm).toHaveBeenCalledOnce()
    onToolCall?.({
      documentId,
      generation: 0,
      call: { id: 'edit-1', name: 'execute_slide_script', input: {} },
    })
    await new Promise((resolve) => setTimeout(resolve, 30))
    await flush()
    await flush()
    expect(enroll).toHaveBeenCalledOnce()
    expect(enhancedSkill.executeTool).toHaveBeenCalled()
    expect(captureSnapshot).toHaveBeenCalledOnce()
    expect(api.toolResult).toHaveBeenCalledOnce()
    onEvent?.({ type: 'done', result: { text: '', cancelled: false, turnLimit: false } })
    await flush()
    expect(complete).toHaveBeenCalledOnce()
    expect(done).toHaveBeenCalledWith(
      expect.objectContaining({ presentation: expect.objectContaining({ status: 'verified' }) }),
    )
    controller.dispose()
  })

  it('uses Standard when standalone Slides has no Codex IPC handler', async () => {
    const transport = manualTransport()
    const api: any = {
      status: vi.fn(async () => {
        throw new Error("No handler registered for 'codex:pc-host:status'")
      }),
    }
    const controller = createAgentController({ transport, skill }, { host: 'slides', api })
    controller.activate()
    await flush()
    expect(controller.run('standard slides')).toBe(true)
    await flush()
    expect(transport.callbacks).toHaveLength(1)
    controller.dispose()
  })

  it('selects Enhanced without dispatching the Standard transport', async () => {
    const transport = manualTransport()
    let documentId: string | null = null
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi.fn(async () => undefined),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn(() => () => undefined),
      onToolCall: vi.fn(() => () => undefined),
    }
    const controller = createAgentController({ transport, skill }, { host: 'slides', api })
    controller.activate()
    await flush()
    expect(controller.run('enhanced slides')).toBe(true)
    await flush()
    expect(api.startTurn).toHaveBeenCalledOnce()
    expect(transport.callbacks).toHaveLength(0)
    controller.dispose()
  })

  it('settles an Enhanced run when startTurn completes without a terminal event', async () => {
    let documentId: string | null = null
    let onEvent: ((event: any) => void) | undefined
    let finishTurn!: () => void
    const turn = new Promise<void>((resolve) => {
      finishTurn = resolve
    })
    const done = vi.fn()
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi.fn(() => turn),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn((listener) => {
        onEvent = listener
        return () => undefined
      }),
      onToolCall: vi.fn(() => () => undefined),
    }
    const controller = createAgentController(
      { transport: manualTransport(), skill, events: { onDone: done } },
      { host: 'slides', api },
    )
    controller.activate()
    await flush()
    expect(controller.run('enhanced slides')).toBe(true)
    await flush()
    expect(done).not.toHaveBeenCalled()

    finishTurn()
    await flush()
    expect(done).toHaveBeenCalledOnce()
    expect(controller.snapshot.busy).toBe(false)
    onEvent?.({ type: 'done', result: { text: '', cancelled: false, turnLimit: false } })
    expect(done).toHaveBeenCalledOnce()
    controller.dispose()
  })

  it('collects concurrent Enhanced tool calls into one renderer batch', async () => {
    let documentId: string | null = null
    let onToolCall: ((request: any) => void) | undefined
    const executeTool = vi.fn(async (call: any) => ({
      output: call.id,
      summary: call.id,
      mutated: false,
    }))
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi.fn(() => new Promise<void>(() => undefined)),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn(() => () => undefined),
      onToolCall: vi.fn((listener) => {
        onToolCall = listener
        return () => undefined
      }),
    }
    const controller = createAgentController(
      { transport: manualTransport(), skill: { ...skill, executeTool } },
      { host: 'slides', api },
    )
    controller.activate()
    await flush()
    expect(controller.run('inspect and edit')).toBe(true)
    await flush()

    onToolCall?.({
      documentId,
      generation: 0,
      call: { id: 'read', name: 'execute_slide_script', input: {} },
    })
    onToolCall?.({
      documentId,
      generation: 0,
      call: { id: 'edit', name: 'execute_slide_script', input: {} },
    })
    await new Promise((resolve) => setTimeout(resolve, 30))
    await flush()
    await flush()

    expect(executeTool).toHaveBeenCalledTimes(2)
    expect(api.toolResult).toHaveBeenCalledTimes(2)
    controller.dispose()
  })

  it('keeps image-search display data local when returning an Enhanced tool result', async () => {
    let documentId: string | null = null
    let onToolCall: ((request: any) => void) | undefined
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi.fn(() => new Promise<void>(() => undefined)),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn(() => () => undefined),
      onToolCall: vi.fn((listener) => {
        onToolCall = listener
        return () => undefined
      }),
    }
    const executeTool = vi.fn(async () => ({
      output: 'https://example.test/image.png',
      summary: 'Found 1 image',
      mutated: false,
      display: { kind: 'images' as const, items: [{ url: 'https://example.test/image.png' }] },
    }))
    const controller = createAgentController(
      { transport: manualTransport(), skill: { ...skill, executeTool } },
      { host: 'slides', api },
    )
    controller.activate()
    await flush()
    expect(controller.run('find an image')).toBe(true)
    await flush()

    onToolCall?.({
      documentId,
      generation: 0,
      call: { id: 'image', name: 'execute_slide_script', input: {} },
    })
    await new Promise((resolve) => setTimeout(resolve, 30))
    await flush()

    expect(api.toolResult).toHaveBeenCalledWith(
      expect.objectContaining({
        execution: {
          output: 'https://example.test/image.png',
          summary: 'Found 1 image',
          mutated: false,
        },
      }),
    )
    controller.dispose()
  })

  it('settles a renderer tool-result turn when the Enhanced terminal event arrives first', async () => {
    let documentId: string | null = null
    let onEvent: ((event: any) => void) | undefined
    let onToolCall: ((request: any) => void) | undefined
    let releaseTool!: () => void
    const toolPending = new Promise<void>((resolve) => {
      releaseTool = resolve
    })
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi.fn(() => new Promise<void>(() => undefined)),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn((listener) => {
        onEvent = listener
        return () => undefined
      }),
      onToolCall: vi.fn((listener) => {
        onToolCall = listener
        return () => undefined
      }),
    }
    const executeTool = vi.fn(async () => {
      await toolPending
      return { output: 'ok', summary: 'read', mutated: false }
    })
    const controller = createAgentController(
      { transport: manualTransport(), skill: { ...skill, executeTool } },
      { host: 'slides', api },
    )
    controller.activate()
    await flush()
    expect(controller.run('inspect')).toBe(true)
    await flush()

    onToolCall?.({
      documentId,
      generation: 0,
      call: { id: 'read', name: 'execute_slide_script', input: {} },
    })
    await new Promise((resolve) => setTimeout(resolve, 30))
    expect(executeTool).toHaveBeenCalledOnce()
    onEvent?.({ type: 'done', result: { text: '', cancelled: false, turnLimit: false } })
    releaseTool()
    await flush()
    await flush()

    expect(api.toolResult).toHaveBeenCalledOnce()
    expect(controller.snapshot.busy).toBe(false)
    controller.dispose()
  })

  it('cancels the Enhanced runtime when the renderer presentation loop fails locally', async () => {
    let documentId: string | null = null
    let onToolCall: ((request: any) => void) | undefined
    const error = vi.fn()
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => undefined),
      startTurn: vi.fn(() => new Promise<void>(() => undefined)),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn(() => () => undefined),
      onToolCall: vi.fn((listener) => {
        onToolCall = listener
        return () => undefined
      }),
    }
    const controller = createAgentController(
      {
        transport: manualTransport(),
        skill: {
          ...skill,
          presentation: {
            prepare: () => ({ kind: 'bypass' as const }),
            enroll: async () => {
              throw new Error('local enrollment failed')
            },
            complete: async () => {
              throw new Error('not reached')
            },
          },
        },
        events: { onError: error },
      },
      { host: 'slides', api },
    )
    controller.activate()
    await flush()
    expect(controller.run('generate')).toBe(true)
    await flush()

    onToolCall?.({
      documentId,
      generation: 0,
      call: { id: 'edit', name: 'execute_slide_script', input: {} },
    })
    await new Promise((resolve) => setTimeout(resolve, 30))
    await flush()

    expect(error).toHaveBeenCalledWith('presentation_enrollment_unavailable')
    expect(api.cancelTurn).toHaveBeenCalledWith(documentId)
    controller.dispose()
  })

  it('waits for an old Enhanced registration to close before reactivating the deck', async () => {
    const transport = manualTransport()
    let documentId: string | null = null
    let registered = false
    let releaseUnregister!: () => void
    const unregisterPending = new Promise<void>((resolve) => {
      releaseUnregister = resolve
    })
    const errors: string[] = []
    const api: any = {
      status: vi.fn(async () => ({ activeAgentRuntime: 'enhanced', documentId })),
      register: vi.fn(async (input: any) => {
        if (registered) throw new Error('enhanced_document_exists')
        registered = true
        documentId = input.documentId
      }),
      unregister: vi.fn(async () => {
        await unregisterPending
        registered = false
      }),
      startTurn: vi.fn(async () => undefined),
      cancelTurn: vi.fn(async () => undefined),
      toolResult: vi.fn(async () => undefined),
      onEvent: vi.fn(() => () => undefined),
      onToolCall: vi.fn(() => () => undefined),
    }
    const controller = createAgentController(
      { transport, skill, events: { onError: (code) => errors.push(code) } },
      { host: 'slides', api },
    )

    controller.activate()
    await flush()
    expect(api.register).toHaveBeenCalledOnce()
    controller.deactivate()
    controller.activate()
    await flush()
    expect(api.register).toHaveBeenCalledOnce()

    releaseUnregister()
    await flush()
    await flush()
    expect(api.register).toHaveBeenCalledTimes(2)
    expect(errors).toEqual([])
    expect(controller.run('first turn after reactivation')).toBe(true)
    await flush()
    expect(api.startTurn).toHaveBeenCalledOnce()
    controller.dispose()
  })

  it('classifies cancellation without exposing raw QC orchestration errors', () => {
    const controller = new AbortController()
    controller.abort()
    expect(classifySlidesQcFailure(new Error('private deck data'), controller.signal)).toBe(
      'cancelled',
    )
    expect(
      classifySlidesQcFailure(Object.assign(new Error('cancelled'), { name: 'AbortError' })),
    ).toBe('cancelled')
    expect(classifySlidesQcFailure(new Error('private deck data'))).toBe('failed')
  })

  it('runs, stops, resets, and restores deck history', async () => {
    const transport = manualTransport()
    const controller = createAgentController({ transport, skill })
    const restored = [
      { role: 'user' as const, text: 'deck request' },
      { role: 'assistant' as const, text: 'deck answer' },
    ]
    controller.restore(restored)
    expect(controller.messages).toEqual(restored)
    controller.run('change the deck')
    await flush()
    controller.stop()
    transport.callbacks[0]!.onDone()
    await flush()
    expect(
      controller.messages.some(
        (message) => message.role === 'user' && message.text === 'change the deck',
      ),
    ).toBe(true)
    controller.reset()
    expect(controller.messages).toEqual([])
  })

  it('preserves snapshot/history-batch, attachment, clarification, and QC host hooks', async () => {
    const transport = manualTransport()
    const hooks = {
      snapshot: vi.fn(),
      finishBatch: vi.fn(),
      persistAttachments: vi.fn(),
      clarify: vi.fn(),
      qc: vi.fn(),
    }
    const controller = createAgentController({
      transport,
      skill,
      captureSnapshot: () => 'deck-before',
      events: {
        onToolStart: () => hooks.clarify(),
        onToolExecuted: ({ snapshotBefore }) => hooks.snapshot(snapshotBefore),
        onDone: () => {
          hooks.finishBatch()
          hooks.persistAttachments(['brief.pdf'])
          hooks.qc()
        },
      },
    })
    controller.run('generate slides')
    await flush()
    transport.callbacks[0]!.onToolCall({ id: 'edit-1', name: 'execute_slide_script', input: {} })
    transport.callbacks[0]!.onDone()
    await flush()
    transport.callbacks[1]!.onDone()
    await flush()
    expect(hooks.clarify).toHaveBeenCalledOnce()
    expect(hooks.snapshot).toHaveBeenCalledWith('deck-before')
    expect(hooks.finishBatch).toHaveBeenCalledOnce()
    expect(hooks.persistAttachments).toHaveBeenCalledWith(['brief.pdf'])
    expect(hooks.qc).toHaveBeenCalledOnce()
  })

  it('unmount disposal cancels and suppresses late snapshot, history, attachment, clarification, and QC hooks', async () => {
    ;(
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true
    const transport = manualTransport()
    const hostHook = vi.fn()
    const ref = {
      current: createAgentController({
        transport,
        skill,
        captureSnapshot: () => 'deck-before',
        events: {
          onToolStart: hostHook,
          onToolExecuted: hostHook,
          onDone: hostHook,
        },
      }),
    }
    const Probe = () => {
      useAgentControllerCleanup(ref)
      return null
    }
    const root = createRoot(document.createElement('div'))
    act(() => root.render(createElement(Probe)))
    ref.current!.run('generate slides')
    await flush()
    const callbacks = transport.callbacks[0]!
    act(() => root.unmount())
    callbacks.onToolCall({ id: 'late', name: 'execute_slide_script', input: {} })
    callbacks.onDone()
    await flush()
    expect(ref.current).toBeNull()
    expect(transport.cancels).toBe(1)
    expect(hostHook).not.toHaveBeenCalled()
  })

  it('survives StrictMode replay, restores and runs, then suppresses late callbacks after real unmount', async () => {
    const transport = manualTransport()
    const done = vi.fn()
    const ref = { current: createAgentController({ transport, skill, events: { onDone: done } }) }
    const Probe = () => {
      useAgentControllerCleanup(ref)
      return null
    }
    const root = createRoot(document.createElement('div'))
    act(() => root.render(createElement(StrictMode, null, createElement(Probe))))
    await flush()
    ref.current?.restore([{ role: 'user', text: 'restored deck' }])
    expect(ref.current?.run('after replay')).toBe(true)
    await flush()
    const callbacks = transport.callbacks[0]!
    act(() => root.unmount())
    callbacks.onDone()
    await flush()
    expect(ref.current).toBeNull()
    expect(done).not.toHaveBeenCalled()
  })

  it('uses the production Slides coordinator for attachments, history snapshots, QC, and clarification stop', async () => {
    const order: string[] = []
    const attachments = [{ name: 'brief.pdf', path: '/brief.pdf' }]
    recordSlidesRunAttachments(attachments, (_attachments) => {
      expect(_attachments).toBe(attachments)
      order.push('attachments')
    })
    await beginSlidesHostRun({
      beginHistoryBatch: async () => {
        order.push('begin-history')
        return true
      },
      isCurrent: () => true,
      markHistoryActive: () => order.push('history-active'),
      finishHistoryBatch: async () => undefined,
      run: () => (order.push('run'), true),
    })
    await completeSlidesHostRun({
      cancelled: false,
      finishHistoryBatch: async () => order.push('snapshot'),
      hasQcPages: () => true,
      clearQcPages: () => order.push('clear-qc'),
      runQc: () => order.push('qc'),
      setBusy: (busy) => order.push(`busy:${busy}`),
    })
    stopSlidesHostRun({
      dismissClarification: () => order.push('clarification'),
      abortQc: () => order.push('abort-qc'),
      stop: () => order.push('stop'),
    })
    expect(order).toEqual([
      'attachments',
      'begin-history',
      'history-active',
      'run',
      'snapshot',
      'busy:false',
      'qc',
      'clarification',
      'abort-qc',
      'stop',
    ])
  })

  it('settles busy state and schedules QC even when snapshot finalization rejects', async () => {
    const settled = vi.fn()
    const qc = vi.fn()
    await expect(
      completeSlidesHostRun({
        cancelled: false,
        finishHistoryBatch: () => Promise.reject(new Error('snapshot failed')),
        hasQcPages: () => true,
        clearQcPages: vi.fn(),
        runQc: qc,
        setBusy: settled,
      }),
    ).rejects.toThrow('snapshot failed')
    expect(settled).toHaveBeenCalledWith(false)
    expect(qc).toHaveBeenCalledOnce()
  })

  it('does not start QC when the deck generation becomes stale during snapshot finalization', async () => {
    let finish!: () => void
    let current = true
    const qc = vi.fn()
    const clear = vi.fn()
    const settled = vi.fn()
    const publishSnapshot = vi.fn()
    const completion = completeSlidesHostRun({
      cancelled: false,
      finishHistoryBatch: () =>
        new Promise<void>((resolve) => {
          finish = resolve
        }),
      isCurrent: () => current,
      hasQcPages: () => true,
      clearQcPages: clear,
      runQc: qc,
      setBusy: settled,
      publishHistorySnapshot: publishSnapshot,
    })

    current = false
    finish()
    await completion

    expect(qc).not.toHaveBeenCalled()
    expect(clear).not.toHaveBeenCalled()
    expect(settled).not.toHaveBeenCalled()
    expect(publishSnapshot).not.toHaveBeenCalled()
  })

  it('closes a pending history batch without running when the launch becomes stale', async () => {
    const controller = createAgentController({ transport: manualTransport(), skill })
    const ref = { current: controller }
    const Probe = () => {
      useAgentControllerCleanup(ref)
      return null
    }
    const root = createRoot(document.createElement('div'))
    act(() => root.render(createElement(Probe)))
    let resolveBegin!: (opened: boolean) => void
    const active = vi.fn()
    const run = vi.fn(() => true)
    const finish = vi.fn(async () => undefined)
    const pending = beginSlidesHostRun({
      beginHistoryBatch: () =>
        new Promise<boolean>((resolve) => {
          resolveBegin = resolve
        }),
      isCurrent: () => ref.current === controller,
      markHistoryActive: active,
      finishHistoryBatch: finish,
      run,
    })
    act(() => root.unmount())
    resolveBegin(true)
    expect(await pending).toBe(false)
    expect(finish).toHaveBeenCalledOnce()
    expect(active).toHaveBeenCalledOnce()
    expect(run).not.toHaveBeenCalled()
  })

  it('activates the history batch before run callbacks can synchronously fail', async () => {
    let active = false
    const finish = vi.fn(async () => {
      if (!active) return
      active = false
    })
    const launched = await beginSlidesHostRun({
      beginHistoryBatch: async () => true,
      isCurrent: () => true,
      markHistoryActive: () => {
        active = true
      },
      finishHistoryBatch: finish,
      run: () => {
        // AgentHarness reports launch errors through the host callback before
        // returning from run(). This models AiPanel's synchronous onError path.
        void finish()
        return true
      },
    })
    await flush()
    expect(launched).toBe(true)
    expect(finish).toHaveBeenCalledOnce()
    expect(active).toBe(false)
  })
})
