import { afterEach, expect, it, vi } from 'vitest'
import { createProductionCodexBootstrap } from '../src/main/codex-engine'

const mock = vi.hoisted(() => ({
  notify: undefined as any,
  document: undefined as any,
  revoke: vi.fn(),
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
        startThread: async () => ({ thread: { id: 'thread' } }),
        startTurn: async () => ({ turn: { id: 'turn' } }),
        interruptTurn: async () => {},
      }
    }
  },
}))
afterEach(() => {
  vi.useRealTimers()
  vi.clearAllMocks()
})

it.each(['applied', 'mutation_expired', 'mutation_cancelled', 'cancel', 'close'])(
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
        output === 'applied'
          ? 'completed'
          : output === 'mutation_cancelled'
            ? 'cancelled'
            : 'failed',
    })
    await engine.close()
  },
)
