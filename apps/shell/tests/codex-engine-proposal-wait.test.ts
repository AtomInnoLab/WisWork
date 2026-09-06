import { afterEach, expect, it, vi } from 'vitest'
import { createProductionCodexBootstrap } from '../src/main/codex-engine'

const mock = vi.hoisted(() => ({
  notify: undefined as any,
  document: undefined as any,
  revoke: vi.fn(),
  startThread: vi.fn(async () => ({ thread: { id: 'thread' } })),
  startTurn: vi.fn(async () => ({ turn: { id: 'turn' } })),
}))
vi.mock('@wiswork/codex-bridge', async (original) => ({
  ...(await original<any>()),
  startResponsesBridge: async () => ({ baseUrl: '', secret: '', close: async () => {} }),
  startDynamicMcpGateway: async () => ({
    url: '',
    secret: '',
    register: (document: any) => {
      mock.document = document
      return () => {}
    },
    beginTurn: () => ({ capability: 'capability' }),
    bindTurn: () => {},
    revokeTurn: mock.revoke,
    close: async () => {},
  }),
  CodexProcessManager: class {
    crashed = new Promise(() => {})
    async stop() {}
    async start() {
      return {
        initialize: async () => {},
        onNotification: (listener: any) => {
          mock.notify = listener
          return () => {}
        },
        startThread: mock.startThread,
        startTurn: mock.startTurn,
        interruptTurn: async () => {},
      }
    }
  },
}))
afterEach(() => {
  vi.useRealTimers()
  vi.clearAllMocks()
  mock.startThread.mockImplementation(async () => ({ thread: { id: 'thread' } }))
  mock.startTurn.mockImplementation(async () => ({ turn: { id: 'turn' } }))
})

it('diagnoses thread and turn start boundaries without retaining request content', async () => {
  const diagnostics: string[] = []
  mock.startThread.mockRejectedValueOnce(new Error('private thread failure'))
  const first = await createProductionCodexBootstrap({
    fetchWithAuth: vi.fn(),
    diagnostics: (code) => diagnostics.push(code),
  }).start({ executablePath: '', onCrash: vi.fn() })
  first.registerDocument!({
    ownerId: 'owner',
    documentId: 'doc',
    host: 'slides',
    generation: 1,
    session: { credentials: {}, listTools: () => [], close: () => {} } as any,
  })
  await expect(
    first.startTurn({ documentId: 'doc', host: 'slides', generation: 1, text: 'PRIVATE PROMPT' }),
  ).rejects.toThrow()
  expect(diagnostics).toEqual(
    expect.arrayContaining(['enhanced_thread_starting', 'enhanced_thread_start_failed']),
  )
  expect(diagnostics.join(',')).not.toContain('PRIVATE')
  await first.close()

  diagnostics.length = 0
  mock.startTurn.mockRejectedValueOnce(new Error('private turn failure'))
  const second = await createProductionCodexBootstrap({
    fetchWithAuth: vi.fn(),
    diagnostics: (code) => diagnostics.push(code),
  }).start({ executablePath: '', onCrash: vi.fn() })
  second.registerDocument!({
    ownerId: 'owner',
    documentId: 'doc',
    host: 'slides',
    generation: 1,
    session: { credentials: {}, listTools: () => [], close: () => {} } as any,
  })
  await expect(
    second.startTurn({ documentId: 'doc', host: 'slides', generation: 1, text: 'PRIVATE PROMPT' }),
  ).rejects.toThrow()
  expect(diagnostics).toEqual(
    expect.arrayContaining([
      'enhanced_thread_starting',
      'enhanced_thread_started',
      'enhanced_turn_starting',
      'enhanced_turn_start_failed',
    ]),
  )
  expect(diagnostics.join(',')).not.toContain('PRIVATE')
  await second.close()
})

it.each(['answered', 'cancelled', 'failed'])(
  'keeps a questionnaire inside one host run; outcome=%s',
  async (outcome) => {
    const cancelled = outcome === 'cancelled'
    const engine = await createProductionCodexBootstrap({ fetchWithAuth: vi.fn() }).start({
      executablePath: '',
      onCrash: vi.fn(),
    })
    engine.registerDocument!({
      ownerId: 'owner',
      documentId: 'doc',
      host: 'slides',
      generation: 1,
      session: {
        credentials: {},
        listTools: () => [{ name: 'build_deck', annotations: { readOnlyHint: false } }],
        close: () => {},
      } as any,
    })
    let done = false
    const running = engine
      .startTurn({ documentId: 'doc', host: 'slides', generation: 1, text: 'make slides' })
      .then(
        () => {
          done = true
        },
        (error: Error) => error,
      )
    await new Promise((r) => setTimeout(r, 0))
    mock.document.onToolEvent({
      type: 'tool-start',
      callId: 'survey',
      toolName: 'ask_clarification',
    })
    if (outcome !== 'answered')
      mock.notify({
        method: 'turn/completed',
        params: { threadId: 'thread', turn: { id: 'turn', status: 'completed' } },
      })
    await new Promise((r) => setTimeout(r, 0))
    expect(done).toBe(false)
    if (cancelled) await engine.cancelTurn('doc')
    mock.document.onToolEvent({
      type: 'tool-complete',
      callId: 'survey',
      toolName: 'ask_clarification',
      isError: outcome === 'failed',
    })
    await new Promise((r) => setTimeout(r, 0))
    if (outcome === 'failed') {
      expect(await running).toMatchObject({ message: 'enhanced_questionnaire_incomplete' })
      expect(done).toBe(false)
      await engine.close()
      return
    }
    if (cancelled) {
      await running
      expect(mock.startTurn).toHaveBeenCalledTimes(1)
      await engine.close()
      return
    }
    expect(mock.startTurn).toHaveBeenCalledTimes(1)
    expect(mock.revoke).not.toHaveBeenCalled()
    mock.document.onToolEvent({
      type: 'tool-complete',
      callId: 'build',
      toolName: 'build_deck',
      isError: false,
    })
    mock.notify({
      method: 'turn/completed',
      params: { threadId: 'thread', turn: { id: 'turn', status: 'completed' } },
    })
    await running
    expect(done).toBe(true)
    await engine.close()
  },
)

it.each(['applied', 'tool_failed', 'mutation_expired', 'mutation_cancelled', 'cancel', 'close'])(
  'waits past idle timeout after model completion until proposal %s',
  async (output) => {
    vi.useFakeTimers()
    const engine = await createProductionCodexBootstrap({ fetchWithAuth: vi.fn() }).start({
      executablePath: '',
      onCrash: vi.fn(),
    })
    const events: any[] = []
    engine.registerDocument!({
      ownerId: 'owner',
      documentId: 'doc',
      host: 'docs',
      generation: 1,
      session: { credentials: {}, listTools: () => [], close: () => {} } as any,
      onEvent: (event) => events.push(event),
    })
    let settle!: (value: any) => void
    const settled = new Promise<any>((resolve) => {
      settle = resolve
    })
    let result: unknown = 'pending'
    const running = engine
      .startTurn({ documentId: 'doc', host: 'docs', generation: 1, text: 'edit' })
      .then(
        () => {
          result = 'done'
        },
        (error) => {
          result = error.message
        },
      )
    await vi.advanceTimersByTimeAsync(0)
    mock.document.onProposal({
      proposalId: 'proposal',
      call: {},
      expiresAt: Date.now() + 300_000,
      summary: {},
      settled,
    })
    await vi.advanceTimersByTimeAsync(120_000)
    expect(result).toBe('pending')
    mock.notify({
      method: 'turn/completed',
      params: { threadId: 'thread', turn: { id: 'turn', status: 'completed' } },
    })
    await vi.advanceTimersByTimeAsync(120_000)
    expect(result).toBe('pending')
    expect(events.some((event) => event.type === 'terminal')).toBe(false)
    if (output === 'cancel' || output === 'close') {
      if (output === 'cancel') await engine.cancelTurn('doc')
      else await engine.closeDocument!('doc')
      await running
      expect(result).toBe(output === 'cancel' ? 'done' : 'document_session_unavailable')
      expect(mock.revoke).toHaveBeenCalled()
      await engine.close()
      return
    }
    settle({
      output,
      summary: output,
      isError: output !== 'applied',
      mutated: output === 'applied',
    })
    await running
    expect(result).toBe(output === 'mutation_expired' ? 'enhanced_proposal_expired' : 'done')
    expect(events.at(-1)).toMatchObject({
      type: 'terminal',
      status:
        output === 'applied' || output === 'tool_failed'
          ? 'completed'
          : output === 'mutation_cancelled'
            ? 'cancelled'
            : 'failed',
    })
    await engine.close()
  },
)
