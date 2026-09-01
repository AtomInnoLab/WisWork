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
  id: 'docs-host',
  systemPrompt: 'system',
  tools: [{ name: 'edit', description: 'edit', inputSchema: { type: 'object' } }],
  executeTool: vi.fn(() => ({ output: 'ok', summary: 'edited', mutated: true })),
}
const flush = () => new Promise((resolve) => setTimeout(resolve, 0))

describe('Docs agent controller', () => {
  it('uses Standard when standalone Docs has no Codex IPC handler', async () => {
    const transport = manualTransport()
    const api: any = {
      status: vi.fn(async () => {
        throw new Error("No handler registered for 'codex:pc-host:status'")
      }),
    }
    const controller = createAgentController({ transport, skill }, { host: 'docs', api })
    controller.activate()
    await flush()
    expect(controller.run('standard docs')).toBe(true)
    await flush()
    expect(transport.callbacks).toHaveLength(1)
    controller.dispose()
  })

  it('selects Enhanced once at activation and never starts the Standard transport', async () => {
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
    const controller = createAgentController({ transport, skill }, { host: 'docs', api })
    controller.activate()
    await flush()
    expect(controller.run('enhanced docs')).toBe(true)
    await flush()
    expect(api.register).toHaveBeenCalledOnce()
    expect(api.startTurn).toHaveBeenCalledOnce()
    expect(transport.callbacks).toHaveLength(0)
    controller.dispose()
  })
  it('keeps an untitled session on first save but isolates different documents', () => {
    expect(shouldResetAgentSession(null, '/saved.docx')).toBe(false)
    expect(shouldResetAgentSession('/a.docx', '/b.docx')).toBe(true)
    expect(shouldResetAgentSession('/a.docx', null)).toBe(true)
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

  it('invalidates pending attachment collection before an old instruction can launch', async () => {
    const owner = createAgentLaunchOwner()
    let resolve!: (images: string[]) => void
    const pending = new Promise<string[]>((done) => (resolve = done))
    const run = vi.fn()
    const launch = owner.launch(
      () => pending,
      (images) => run('old', images),
    )
    owner.invalidate()
    resolve(['late-image'])
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
      { role: 'user' as const, text: 'restored with attachments transcript' },
      { role: 'assistant' as const, text: 'restored response' },
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

  it('delivers pre-edit snapshots and attachments to Docs persistence hooks', async () => {
    const transport = manualTransport()
    const rememberSnapshot = vi.fn()
    const persistAttachments = vi.fn()
    const sentAttachments = [{ name: 'brief.pdf', path: '/brief.pdf' }]
    const controller = createAgentController({
      transport,
      skill,
      captureSnapshot: () => ({ type: 'doc', content: [] }),
      events: {
        onToolExecuted: ({ snapshotBefore }) => rememberSnapshot(snapshotBefore),
        onDone: () => persistAttachments(sentAttachments),
      },
    })
    controller.run('edit docs')
    await flush()
    transport.callbacks[0]!.onToolCall({ id: 'edit-1', name: 'edit', input: {} })
    transport.callbacks[0]!.onDone()
    await flush()
    transport.callbacks[1]!.onDone()
    await flush()

    expect(rememberSnapshot).toHaveBeenCalledWith({ type: 'doc', content: [] })
    expect(persistAttachments).toHaveBeenCalledWith(sentAttachments)
  })

  it('disposal cancels work, nulls the ref, and blocks late snapshot/attachment persistence', async () => {
    const transport = manualTransport()
    const rememberSnapshot = vi.fn()
    const persistAttachments = vi.fn()
    const sentAttachments = [{ name: 'brief.pdf', path: '/brief.pdf' }]
    const controller = createAgentController({
      transport,
      skill,
      captureSnapshot: () => ({ type: 'doc', content: [] }),
      events: {
        onToolExecuted: ({ snapshotBefore }) => rememberSnapshot(snapshotBefore),
        onDone: () => persistAttachments(sentAttachments),
      },
    })
    const ref = { current: controller }
    controller.run('edit docs')
    await flush()
    const callbacks = transport.callbacks[0]!

    disposeAgentController(ref)
    callbacks.onToolCall({ id: 'late', name: 'edit', input: {} })
    callbacks.onDone()
    await flush()

    expect(ref.current).toBeNull()
    expect(transport.cancels).toBe(1)
    expect(rememberSnapshot).not.toHaveBeenCalled()
    expect(persistAttachments).not.toHaveBeenCalled()
  })
})
