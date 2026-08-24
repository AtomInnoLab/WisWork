import { describe, expect, it, vi } from 'vitest'
import { act, createElement, StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import type { AgentStreamCallbacks, AgentTransport } from '@wiswork/agent-core'
import {
  createAgentController,
  disposeAgentController,
  useAgentControllerCleanup,
  createAgentLaunchOwner,
  createAgentRunStartingGuard,
  shouldResetAgentSession,
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
  id: 'markdown-host',
  systemPrompt: 'system',
  tools: [{ name: 'edit', description: 'edit', inputSchema: { type: 'object' } }],
  executeTool: vi.fn(() => ({ output: 'ok', summary: 'edited', mutated: true })),
}
const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('Markdown agent controller', () => {
  it('keeps an untitled session on first save but isolates different documents', () => {
    expect(shouldResetAgentSession(null, '/saved.md')).toBe(false)
    expect(shouldResetAgentSession('/a.md', '/b.md')).toBe(true)
    expect(shouldResetAgentSession('/a.md', null)).toBe(true)
  })

  it('synchronously rejects concurrent prelaunch sends and allows Stop to clear the guard', () => {
    const guard = createAgentRunStartingGuard()
    const first = guard.begin()
    expect(first).not.toBeNull()
    expect(guard.begin()).toBeNull()
    guard.clear()
    expect(guard.begin()).not.toBeNull()
  })
  it('survives StrictMode replay, then terminally suppresses callbacks after real unmount', async () => {
    const transport = manualTransport()
    const late = vi.fn()
    const ref = { current: createAgentController({ transport, skill, events: { onDone: late } }) }
    const Probe = () => {
      useAgentControllerCleanup(ref)
      return null
    }
    const root = createRoot(document.createElement('div'))
    act(() => root.render(createElement(StrictMode, null, createElement(Probe))))
    await flush()
    expect(ref.current?.run('after replay')).toBe(true)
    await flush()
    const callbacks = transport.callbacks[0]!
    act(() => root.unmount())
    callbacks.onDone()
    await flush()
    expect(ref.current).toBeNull()
    expect(transport.cancels).toBe(1)
    expect(late).not.toHaveBeenCalled()
  })

  it('invalidates pending settings collection before an old instruction can launch', async () => {
    const owner = createAgentLaunchOwner()
    let resolve!: (settings: string) => void
    const pending = new Promise<string>((done) => (resolve = done))
    const run = vi.fn()
    const launch = owner.launch(
      () => pending,
      () => run('old'),
    )
    owner.invalidate()
    resolve('late-settings')
    expect(await launch).toBe(false)
    expect(run).not.toHaveBeenCalled()
  })
  it('disposes and clears the controller when its owning component unmounts', async () => {
    ;(
      globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }
    ).IS_REACT_ACT_ENVIRONMENT = true
    const transport = manualTransport()
    const ref = { current: createAgentController({ transport, skill }) }
    const Probe = () => {
      useAgentControllerCleanup(ref)
      return null
    }
    const container = document.createElement('div')
    const root = createRoot(container)
    act(() => root.render(createElement(Probe)))
    ref.current!.run('pending')
    await flush()
    act(() => root.unmount())
    expect(transport.cancels).toBe(1)
    expect(ref.current).toBeNull()
  })

  it('preserves history on stop, clears it on reset, and restores persisted history', async () => {
    const transport = manualTransport()
    const controller = createAgentController({ transport, skill })
    const restored = [
      { role: 'user' as const, text: 'restored question' },
      { role: 'assistant' as const, text: 'restored answer' },
    ]
    controller.restore(restored)
    expect(controller.messages).toEqual(restored)

    controller.run('new request')
    await flush()
    controller.stop()
    transport.callbacks[0]!.onDone()
    await flush()
    expect(
      controller.messages.some(
        (message) => message.role === 'user' && message.text === 'new request',
      ),
    ).toBe(true)

    controller.reset()
    expect(controller.messages).toEqual([])
  })

  it('delivers pre-edit snapshots and completion to Markdown autosave hooks', async () => {
    const transport = manualTransport()
    const rememberSnapshot = vi.fn()
    const autosave = vi.fn()
    const controller = createAgentController({
      transport,
      skill,
      captureSnapshot: () => 'markdown-before',
      events: {
        onToolExecuted: ({ snapshotBefore }) => rememberSnapshot(snapshotBefore),
        onDone: () => autosave(true),
      },
    })
    controller.run('edit markdown')
    await flush()
    transport.callbacks[0]!.onToolCall({ id: 'edit-1', name: 'edit', input: {} })
    transport.callbacks[0]!.onDone()
    await flush()
    transport.callbacks[1]!.onDone()
    await flush()

    expect(rememberSnapshot).toHaveBeenCalledWith('markdown-before')
    expect(autosave).toHaveBeenCalledWith(true)
  })

  it('disposal cancels work, nulls the ref, and blocks late snapshot/autosave hooks', async () => {
    const transport = manualTransport()
    const rememberSnapshot = vi.fn()
    const autosave = vi.fn()
    const controller = createAgentController({
      transport,
      skill,
      captureSnapshot: () => 'markdown-before',
      events: {
        onToolExecuted: ({ snapshotBefore }) => rememberSnapshot(snapshotBefore),
        onDone: () => autosave(),
      },
    })
    const ref = { current: controller }
    controller.run('edit markdown')
    await flush()
    const callbacks = transport.callbacks[0]!

    disposeAgentController(ref)
    callbacks.onToolCall({ id: 'late', name: 'edit', input: {} })
    callbacks.onDone()
    await flush()

    expect(ref.current).toBeNull()
    expect(transport.cancels).toBe(1)
    expect(rememberSnapshot).not.toHaveBeenCalled()
    expect(autosave).not.toHaveBeenCalled()
  })
})
