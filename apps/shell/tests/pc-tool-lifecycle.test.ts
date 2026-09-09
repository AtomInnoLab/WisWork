import { describe, expect, it, vi } from 'vitest'
import {
  createEnhancedRendererClient,
  EnhancedAgentRuntime,
  ENHANCED_HOSTS,
  PC_HOST_CODEX_CHANNELS as C,
} from '@wiswork/agent-runtime'
import { registerPcCodexHosts } from '../src/main/pc-codex-hosts'

function fixture() {
  const handlers = new Map<string, (...args: any[]) => any>()
  const listeners = new Map<string, Set<(value: any) => void>>()
  const sent: Array<{ channel: string; value: any }> = []
  let document: any
  const owner = {
    id: 91,
    isDestroyed: () => false,
    send(channel: string, value: any) {
      sent.push({ channel, value })
      for (const listener of listeners.get(channel) ?? []) listener(value)
    },
  }
  const listen = (channel: string, listener: (value: any) => void) => {
    const channelListeners = listeners.get(channel) ?? new Set()
    listeners.set(channel, channelListeners)
    channelListeners.add(listener)
    return () => channelListeners.delete(listener)
  }
  const invoke = (channel: string, ...args: unknown[]) =>
    handlers.get(channel)!({ sender: owner }, ...args)
  const registrar = registerPcCodexHosts({
    ipcMain: { handle: (channel, handler) => handlers.set(channel, handler) },
    runtime: {
      configuredAgentRuntime: 'enhanced',
      registerDocument(input: unknown) {
        document = input
        return { close: async () => undefined }
      },
    } as never,
    policy: {
      globalEnabled: true,
      rawOfficeEnabled: false,
      hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as never,
    },
    hostForOwner: (candidate) => (candidate === owner ? 'docs' : null),
  })
  const bridge = {
    register: async (input: unknown) => invoke(C.register, input),
    unregister: async (id: string, generation: number) => invoke(C.unregister, id, generation),
    status: async () => ({ activeAgentRuntime: 'enhanced' as const, documentId: 'document' }),
    startTurn: async () => undefined,
    cancelTurn: async () => {
      // Actual engine order: revoke the grant (cancelAll), then publish terminal.
      document.toolSession.cancelAll(document.toolSession.credentials)
      document.onEvent({ type: 'terminal', status: 'cancelled' })
    },
    subscribe: (_id: string, listener: (value: any) => void) => listen(C.event, listener),
    onToolCall: (listener: (value: any) => void) => listen(C.toolCall, listener),
    onToolCancel: (listener: (value: any) => void) => listen(C.toolCancel, listener),
    toolResult: async (value: unknown) => invoke(C.toolResult, value),
  }
  const register = () =>
    bridge.register({
      host: 'docs',
      documentId: 'document',
      generation: 1,
      systemPrompt: '',
      tools: ['read_blocks', 'replace_blocks'].map((name) => ({
        name,
        description: name,
        inputSchema: { type: 'object' },
      })),
      mutatingTools: ['replace_blocks'],
    })
  return {
    bridge,
    register,
    registrar,
    sent,
    invoke,
    get document() {
      return document
    },
  }
}

describe('PC semantic tool lifecycle', () => {
  it.each(['invalid_tool_call', 'tool_call_in_progress'])(
    'projects %s without executing tools and ignores duplicate lifecycle frames',
    async (errorCode) => {
      const f = fixture()
      await f.register()
      const start = {
        type: 'tool-start',
        turnId: 'turn-a',
        callId: 'blocked',
        toolName: 'read_blocks',
      }
      f.document.onEvent(start)
      f.document.onEvent(start)
      f.document.onEvent({ ...start, type: 'tool-complete', isError: true, errorCode })
      f.document.onEvent({ ...start, type: 'tool-complete', isError: true, errorCode })
      expect(f.sent.filter(({ channel }) => channel === C.toolCall)).toHaveLength(0)
      expect(f.sent.filter(({ value }) => value.type === 'tool-start')).toHaveLength(1)
      expect(f.sent.filter(({ value }) => value.type === 'tool-executed')).toEqual([
        expect.objectContaining({
          value: expect.objectContaining({
            event: expect.objectContaining({
              call: expect.objectContaining({ id: 'blocked', invocationId: 'turn-a:blocked' }),
              execution: expect.objectContaining({ output: errorCode, isError: true }),
            }),
          }),
        }),
      ])
      await f.registrar.close()
    },
  )

  it('terminates pending activity before turn completion and rejects stale receipts', async () => {
    const f = fixture()
    await f.register()
    const frame = { turnId: 'old-turn', callId: 'old', toolName: 'read_blocks' }
    f.document.onEvent({ type: 'tool-start', ...frame })
    const pending = f.document.toolSession.callTool(f.document.toolSession.credentials, {
      id: 'old',
      invocationId: 'old-turn:old',
      name: 'read_blocks',
      input: {},
    })
    f.document.onEvent({ type: 'terminal', status: 'cancelled' })
    await expect(pending).resolves.toMatchObject({ isError: true })
    const events = f.sent.filter(({ channel }) => channel === C.event).map(({ value }) => value)
    expect(events.map(({ type }) => type)).toEqual(['tool-start', 'tool-executed', 'done'])
    expect(events[1].event.execution.output).toBe('tool_cancelled')
    f.document.onEvent({
      type: 'tool-start',
      turnId: 'new-turn',
      callId: 'new',
      toolName: 'read_blocks',
    })
    const count = f.sent.length
    f.document.onEvent({ type: 'tool-complete', ...frame, isError: false })
    expect(() =>
      f.invoke(C.toolResult, {
        documentId: 'document',
        generation: 1,
        callId: 'old',
        execution: { output: 'late', summary: 'late success' },
      }),
    ).toThrow('enhanced_untrusted_request')
    expect(f.sent).toHaveLength(count)
    await f.registrar.close()
  })

  it('keeps one rich receipt when gateway lifecycle surrounds renderer execution', async () => {
    const f = fixture()
    await f.register()
    const frame = { turnId: 'turn-a', callId: 'read', toolName: 'read_blocks' }
    f.document.onEvent({ type: 'tool-start', ...frame })
    const result = f.document.toolSession.callTool(f.document.toolSession.credentials, {
      id: 'read',
      invocationId: 'turn-a:read',
      name: 'read_blocks',
      input: { block: 'one' },
    })
    f.invoke(C.toolResult, {
      documentId: 'document',
      generation: 1,
      callId: 'read',
      execution: { output: 'actual content', summary: 'Read one block', mutated: false },
    })
    await result
    f.document.onEvent({ type: 'tool-complete', ...frame, isError: false })
    expect(f.sent.filter(({ value }) => value.type === 'tool-start')).toHaveLength(1)
    const complete = f.sent.filter(({ value }) => value.type === 'tool-executed')
    expect(complete).toHaveLength(1)
    expect(complete[0].value.event).toMatchObject({
      call: { input: { block: 'one' } },
      execution: { output: 'actual content', summary: 'Read one block' },
    })
    await f.registrar.close()
  })

  it('cannot publish a cancelled read into the next run even when the renderer finishes late', async () => {
    const f = fixture()
    let release!: (value: any) => void
    const executed = vi.fn()
    const runtime = new EnhancedAgentRuntime(createEnhancedRendererClient(f.bridge))
    const session = runtime.createSession({
      host: 'docs',
      document: { id: 'document', generation: 1 },
      skill: {
        id: 'docs',
        systemPrompt: '',
        tools: [{ name: 'read_blocks', description: 'read', inputSchema: {} }],
        executeTool: () =>
          new Promise((resolve) => {
            release = resolve
          }),
      },
      events: { onToolExecuted: executed },
    })
    session.run('first')
    await vi.waitFor(() => expect(f.document).toBeDefined())
    const pending = f.document.toolSession.callTool(f.document.toolSession.credentials, {
      id: 'old-read',
      name: 'read_blocks',
      input: {},
    })
    await vi.waitFor(() => expect(release).toBeDefined())
    session.stop()
    await vi.waitFor(() => expect(session.snapshot.busy).toBe(false))
    await expect(pending).resolves.toMatchObject({ output: 'tool_cancelled' })
    executed.mockClear()
    expect(session.run('second')).toBe(true)
    release({ output: 'old content', summary: 'old successful read' })
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(executed).not.toHaveBeenCalled()
    expect(f.sent).toContainEqual({
      channel: C.toolCancel,
      value: { documentId: 'document', generation: 1, callId: 'old-read' },
    })
    await runtime.dispose()
    await f.registrar.close()
  })
})
