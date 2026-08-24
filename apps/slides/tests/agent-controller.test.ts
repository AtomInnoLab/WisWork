import { act, createElement, StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import type { AgentStreamCallbacks, AgentTransport } from '@wiswork/agent-core'
import {
  createAgentController,
  beginSlidesHostRun,
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
      'run',
      'history-active',
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
    expect(active).not.toHaveBeenCalled()
    expect(run).not.toHaveBeenCalled()
  })
})
