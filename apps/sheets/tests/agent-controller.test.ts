// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import type { AgentStreamCallbacks, AgentTransport } from '@wiswork/agent-core'
import {
  createAgentController,
  disposeAgentController,
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
  id: 'sheets-host',
  systemPrompt: 'system',
  tools: [{ name: 'apply_plan', description: 'apply', inputSchema: { type: 'object' } }],
  executeTool: vi.fn(() => ({ output: 'ok', summary: 'applied', mutated: true })),
}
const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('Sheets agent controller', () => {
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

  it('dispose helper is terminal and clears the owning ref', () => {
    const controller = createAgentController({ transport: manualTransport(), skill })
    const ref = { current: controller }
    disposeAgentController(ref)
    expect(ref.current).toBeNull()
    expect(controller.run('late run')).toBe(false)
  })
})
