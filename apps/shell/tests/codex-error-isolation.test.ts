import { expect, it, vi } from 'vitest'

const state = vi.hoisted(() => ({ notify: undefined as ((n: any) => void) | undefined }))
vi.mock('@wiswork/codex-bridge', () => ({
  CodexProcessManager: class {
    crashed = new Promise<void>(() => {})
    async start() {
      return {
        initialize: async () => {},
        onNotification: (listener: any) => {
          state.notify = listener
          return () => {}
        },
        startThread: async () => ({ thread: { id: 'thread-1' } }),
        startTurn: async () => ({ turn: { id: 'turn-1' } }),
        interruptTurn: async () => {},
      }
    }
    async stop() {}
  },
  startResponsesBridge: async () => ({ baseUrl: '', secret: '', close: async () => {} }),
  startDynamicMcpGateway: async () => ({
    url: '',
    secret: '',
    register: () => () => {},
    beginTurn: () => ({ capability: 'test-capability' }),
    bindTurn: () => {},
    revokeTurn: () => {},
    close: async () => {},
  }),
  createDocumentCarrierIssuer: vi.fn(),
}))
import { createProductionCodexBootstrap } from '../src/main/codex-engine'

it('keeps the process alive after a non-retryable error in one turn', async () => {
  const crashed = vi.fn()
  const engine = await createProductionCodexBootstrap({ fetchWithAuth: vi.fn() }).start({
    executablePath: '/test/codex',
    onCrash: crashed,
  })
  engine.registerDocument!({
    ownerId: 'owner',
    documentId: 'doc',
    host: 'slides',
    generation: 1,
    session: { credentials: {}, listTools: () => [], close: () => {} } as any,
  })
  const pending = engine.startTurn({
    documentId: 'doc',
    host: 'slides',
    generation: 1,
    text: 'test',
  })
  const rejected = expect(pending).rejects.toThrow('enhanced_request_rejected')
  await new Promise((resolve) => setTimeout(resolve, 0))
  state.notify!({
    method: 'error',
    params: {
      threadId: 'thread-1',
      turnId: 'turn-1',
      willRetry: false,
      error: { codexErrorInfo: 'badRequest' },
    },
  })
  await rejected
  expect(crashed).not.toHaveBeenCalled()
  const retry = engine.startTurn({
    documentId: 'doc',
    host: 'slides',
    generation: 1,
    text: 'retry',
  })
  await new Promise((resolve) => setTimeout(resolve, 0))
  state.notify!({
    method: 'turn/completed',
    params: { threadId: 'thread-1', turn: { id: 'turn-1', status: 'completed' } },
  })
  await expect(retry).resolves.toBeUndefined()
  expect(crashed).not.toHaveBeenCalled()
  await engine.close()
})
