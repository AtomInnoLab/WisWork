import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import type { AgentStreamCallbacks, AgentTransport } from '@wiswork/agent-core'
import {
  createAgentController,
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
})
