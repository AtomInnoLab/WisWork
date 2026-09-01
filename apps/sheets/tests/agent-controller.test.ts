// @vitest-environment jsdom
import { act, createElement, StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import type { AgentStreamCallbacks, AgentTransport } from '@wiswork/agent-core'
import {
  createAsyncGenerationGate,
  createAgentController,
  createSheetsChatLoadCoordinator,
  classifySheetsDocumentTransition,
  bindSheetsSession,
  getSheetsDocumentIdentity,
  disposeAgentController,
  restoreSheetsSession,
  selectSheetsExecution,
  settleSheetsApplyPromises,
  useAgentControllerCleanup,
} from '../src/renderer/ai/agent-controller'

function manualTransport(): AgentTransport & {
  callbacks: AgentStreamCallbacks[]
  cancels: number
  requests: unknown[]
} {
  const transport = {
    callbacks: [] as AgentStreamCallbacks[],
    cancels: 0,
    requests: [] as unknown[],
    stream(request: unknown, callbacks: AgentStreamCallbacks) {
      transport.requests.push(request)
      transport.callbacks.push(callbacks)
      return { cancel: () => transport.cancels++ }
    },
  }
  return transport
}

const skill = {
  id: 'sheets-host',
  systemPrompt: 'system',
  tools: [{ name: 'apply_plan', description: 'apply', inputSchema: { type: 'object' } }],
  executeTool: vi.fn(() => ({ output: 'ok', summary: 'applied', mutated: true })),
}
const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('Sheets agent controller', () => {
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
    const controller = createAgentController({ transport, skill }, { host: 'sheets', api })
    controller.activate()
    await flush()
    expect(controller.run('enhanced sheets')).toBe(true)
    await flush()
    expect(api.startTurn).toHaveBeenCalledOnce()
    expect(transport.callbacks).toHaveLength(0)
    controller.dispose()
  })
  it('atomically rejects every side effect from a stale session chat load', () => {
    const coordinator = createSheetsChatLoadCoordinator()
    const effects: string[] = []
    const sessionA = coordinator.begin()
    const sessionB = coordinator.begin()

    expect(
      coordinator.commit(sessionA, () => {
        effects.push('setHistoricChat(A)', 'chatRefIdsRef=A', 'restore(A)')
      }),
    ).toBe(false)
    expect(effects).toEqual([])
    expect(
      coordinator.commit(sessionB, () => {
        effects.push('setHistoricChat(B)', 'chatRefIdsRef=B', 'restore(B)')
      }),
    ).toBe(true)
    expect(effects).toEqual(['setHistoricChat(B)', 'chatRefIdsRef=B', 'restore(B)'])
  })

  it('invalidates an async prelaunch so late attachment work cannot start the old run', async () => {
    const gate = createAsyncGenerationGate()
    const run = vi.fn()
    let finishImages!: () => void
    const images = new Promise<void>((resolve) => {
      finishImages = resolve
    })
    const token = gate.begin()
    const launch = images.then(() => gate.commit(token, run))

    gate.invalidate()
    finishImages()
    await launch

    expect(run).not.toHaveBeenCalled()
  })

  it('runs, stops without clearing history, resets, and restores session-bound history', async () => {
    const transport = manualTransport()
    const controller = createAgentController({ transport, skill })
    const restored = [
      { role: 'user' as const, text: 'workbook session request' },
      { role: 'assistant' as const, text: 'workbook session answer' },
    ]
    controller.restore(restored)
    expect(controller.messages).toEqual(restored)
    expect(controller.run('apply workbook plan')).toBe(true)
    await flush()
    controller.stop()
    transport.callbacks[0]!.onDone()
    await flush()
    expect(
      controller.messages.some(
        (message) => message.role === 'user' && message.text === 'apply workbook plan',
      ),
    ).toBe(true)
    controller.reset()
    expect(controller.messages).toEqual([])
  })

  it('delivers apply completion to the workbook autosave hook', async () => {
    const transport = manualTransport()
    const applyFinished = vi.fn()
    const autosave = vi.fn()
    const controller = createAgentController({
      transport,
      skill,
      captureSnapshot: () => 'workbook-before',
      events: {
        onToolExecuted: ({ snapshotBefore }) => applyFinished(snapshotBefore),
        onDone: () => autosave(),
      },
    })
    controller.run('apply workbook plan')
    await flush()
    transport.callbacks[0]!.onToolCall({ id: 'apply-1', name: 'apply_plan', input: {} })
    transport.callbacks[0]!.onDone()
    await flush()
    transport.callbacks[1]!.onDone()
    await flush()
    expect(applyFinished).toHaveBeenCalledWith('workbook-before')
    expect(autosave).toHaveBeenCalledOnce()
  })

  it('reset suppresses late apply and autosave callbacks from the previous run', async () => {
    const transport = manualTransport()
    const hostHook = vi.fn()
    const controller = createAgentController({
      transport,
      skill,
      events: { onToolExecuted: hostHook, onDone: hostHook },
    })
    controller.run('apply workbook plan')
    await flush()
    const callbacks = transport.callbacks[0]!
    controller.reset()
    callbacks.onToolCall({ id: 'late', name: 'apply_plan', input: {} })
    callbacks.onDone()
    await flush()
    expect(controller.snapshot.status).toBe('idle')
    expect(controller.messages).toEqual([])
    expect(hostHook).not.toHaveBeenCalled()
  })

  it('disposes on unmount and blocks late apply/autosave callbacks', async () => {
    ;(
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true
    const transport = manualTransport()
    const applyFinished = vi.fn()
    const autosave = vi.fn()
    const ref = {
      current: createAgentController({
        transport,
        skill,
        captureSnapshot: () => 'workbook-before',
        events: {
          onToolExecuted: applyFinished,
          onDone: autosave,
        },
      }),
    }
    const Probe = () => {
      useAgentControllerCleanup(ref)
      return null
    }
    const root = createRoot(document.createElement('div'))
    act(() => root.render(createElement(Probe)))
    ref.current!.run('apply workbook plan')
    await flush()
    const callbacks = transport.callbacks[0]!
    act(() => root.unmount())
    callbacks.onToolCall({ id: 'late', name: 'apply_plan', input: {} })
    callbacks.onDone()
    await flush()
    expect(ref.current).toBeNull()
    expect(transport.cancels).toBe(1)
    expect(applyFinished).not.toHaveBeenCalled()
    expect(autosave).not.toHaveBeenCalled()
  })

  it('survives StrictMode replay, can restore and run, then terminally disposes on real unmount', async () => {
    const transport = manualTransport()
    const late = vi.fn()
    const ref = {
      current: createAgentController({ transport, skill, events: { onDone: late } }),
    }
    const Probe = () => {
      useAgentControllerCleanup(ref)
      return null
    }
    const root = createRoot(document.createElement('div'))
    act(() => root.render(createElement(StrictMode, null, createElement(Probe))))
    await flush()
    expect(
      restoreSheetsSession(ref.current, 'session-a', () => 'session-a', [
        { role: 'user', text: 'restored' },
      ]),
    ).toBe(true)
    expect(ref.current?.run('after replay')).toBe(true)
    await flush()
    const callbacks = transport.callbacks[0]!
    act(() => root.unmount())
    callbacks.onDone()
    await flush()
    expect(ref.current).toBeNull()
    expect(late).not.toHaveBeenCalled()
  })

  it('uses the production Sheets coordinator for planner choice, session binding, and apply-before-autosave', async () => {
    expect(selectSheetsExecution(false)).toBe('planner')
    expect(selectSheetsExecution(true)).toBe('agent')
    const controller = createAgentController({ transport: manualTransport(), skill })
    expect(
      restoreSheetsSession(controller, 'stale-session', () => 'current-session', [
        { role: 'user', text: 'wrong workbook' },
      ]),
    ).toBe(false)
    expect(controller.messages).toEqual([])

    let finishApply!: (ok: boolean) => void
    const order: string[] = []
    const applies = [
      new Promise<boolean>((resolve) => {
        finishApply = (ok) => {
          order.push('apply')
          resolve(ok)
        }
      }),
    ]
    const settled = settleSheetsApplyPromises(applies, async () => {
      order.push('autosave')
    })
    expect(applies).toEqual([])
    expect(order).toEqual([])
    finishApply(true)
    expect(await settled).toBe(true)
    expect(order).toEqual(['apply', 'autosave'])
  })

  it('resets A history before binding B and rejects a stale A load', async () => {
    const transport = manualTransport()
    const controller = createAgentController({ transport, skill })
    const binding = { current: undefined as string | undefined }
    bindSheetsSession(controller, binding, 'session-a')
    restoreSheetsSession(controller, 'session-a', () => binding.current, [
      { role: 'user', text: 'secret from A' },
      { role: 'assistant', text: 'A answer' },
    ])
    bindSheetsSession(controller, binding, 'session-b')
    expect(controller.messages).toEqual([])
    expect(
      restoreSheetsSession(controller, 'session-a', () => binding.current, [
        { role: 'user', text: 'late A' },
      ]),
    ).toBe(false)
    restoreSheetsSession(controller, 'session-b', () => binding.current, [
      { role: 'user', text: 'B context' },
      { role: 'assistant', text: 'B answer' },
    ])
    controller.run('next B request')
    await flush()
    expect(JSON.stringify(transport.requests[0])).not.toContain('secret from A')
    expect(JSON.stringify(transport.requests[0])).toContain('B context')
  })

  it('preserves context across a sidecar session rotation for the same workbook path', () => {
    const controller = createAgentController({ transport: manualTransport(), skill })
    const binding = { current: undefined as string | undefined }
    const path = '/workbooks/forecast.xlsx'
    const identityA = getSheetsDocumentIdentity({
      sessionId: 'session-a',
      documentInstanceId: 'document-a',
      path,
    })
    const identityB = getSheetsDocumentIdentity({
      sessionId: 'session-b',
      documentInstanceId: 'document-a',
      path,
    })

    bindSheetsSession(controller, binding, identityA)
    controller.restore([
      { role: 'user', text: 'keep this workbook context' },
      { role: 'assistant', text: 'kept answer' },
    ])
    bindSheetsSession(controller, binding, identityB)

    expect(identityB).toBe(identityA)
    expect(controller.messages).toEqual([
      { role: 'user', text: 'keep this workbook context' },
      { role: 'assistant', text: 'kept answer' },
    ])
  })

  it('resets context and rejects stale loads when a different workbook is opened', () => {
    const controller = createAgentController({ transport: manualTransport(), skill })
    const binding = { current: undefined as string | undefined }
    const identityA = getSheetsDocumentIdentity({
      sessionId: 'session-a',
      documentInstanceId: 'document-a',
      path: '/workbooks/a.xlsx',
    })
    const identityB = getSheetsDocumentIdentity({
      sessionId: 'session-b',
      documentInstanceId: 'document-b',
      path: '/workbooks/b.xlsx',
    })
    bindSheetsSession(controller, binding, identityA)
    controller.restore([{ role: 'user', text: 'A context' }])

    bindSheetsSession(controller, binding, identityB)

    expect(controller.messages).toEqual([])
    expect(
      restoreSheetsSession(controller, identityA, () => binding.current, [
        { role: 'user', text: 'late A load' },
      ]),
    ).toBe(false)
  })

  it('treats reopening the same path as a new workbook instance', () => {
    const path = '/workbooks/forecast.xlsx'
    expect(
      getSheetsDocumentIdentity({
        sessionId: 'session-a',
        documentInstanceId: 'document-a',
        path,
      }),
    ).not.toBe(
      getSheetsDocumentIdentity({
        sessionId: 'session-b',
        documentInstanceId: 'document-b',
        path,
      }),
    )
  })

  it('rebinds persistence only when a sidecar rotation keeps the document instance', () => {
    expect(classifySheetsDocumentTransition('document-a', 'document-a')).toBe('rebind')
    expect(classifySheetsDocumentTransition('document-a', 'document-b')).toBe('open')
  })

  it('dispose helper is terminal and clears the owning ref', () => {
    const controller = createAgentController({ transport: manualTransport(), skill })
    const ref = { current: controller }
    disposeAgentController(ref)
    expect(ref.current).toBeNull()
    expect(controller.run('late run')).toBe(false)
  })
})
