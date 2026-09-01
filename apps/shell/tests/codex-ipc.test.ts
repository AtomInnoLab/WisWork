import { describe, expect, it, vi } from 'vitest'
import { ENHANCED_HOSTS } from '@wiswork/agent-runtime'
import { registerCodexRuntimeIpc } from '../src/main/codex-ipc'
import { ShellCodexRuntime } from '../src/main/codex-runtime'
import { CODEX_RUNTIME_CHANNELS } from '../src/shared/codex-api'

describe('Codex runtime IPC', () => {
  it('rejects sender/document substitution and forwards only authoritative bounded requests', async () => {
    const handlers = new Map<string, (...args: any[]) => any>()
    const owner = { isDestroyed: () => false }
    const attacker = { isDestroyed: () => false }
    const engine = {
      startTurn: vi.fn(async () => undefined),
      cancelTurn: vi.fn(async () => undefined),
      closeDocument: vi.fn(async () => undefined),
      close: vi.fn(async () => undefined),
    }
    const runtime = new ShellCodexRuntime({
      activeAgentRuntime: 'enhanced',
      policy: {
        globalEnabled: true,
        rawOfficeEnabled: false,
        hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as any,
      },
      isSignedIn: async () => true,
      resolveExecutable: async () => '/private/codex',
      bootstrap: { start: async () => engine },
    })
    await runtime.initialize()
    runtime.registerDocument({ owner, documentId: 'doc', host: 'docs', generation: 1 })
    registerCodexRuntimeIpc({
      ipcMain: { handle: (channel, handler) => handlers.set(channel, handler) },
      runtime,
      documentIdForOwner: (candidate) => (candidate === owner ? 'doc' : null),
    })
    await handlers.get(CODEX_RUNTIME_CHANNELS.startTurn)!(
      { sender: owner },
      {
        documentId: 'doc',
        text: 'hello',
      },
    )
    expect(engine.startTurn).toHaveBeenCalledOnce()
    await expect(
      handlers.get(CODEX_RUNTIME_CHANNELS.startTurn)!(
        { sender: attacker },
        {
          documentId: 'doc',
          text: 'attack',
        },
      ),
    ).rejects.toThrow('enhanced_untrusted_request')
    await expect(
      handlers.get(CODEX_RUNTIME_CHANNELS.startTurn)!(
        { sender: owner },
        {
          documentId: 'other',
          text: 'attack',
        },
      ),
    ).rejects.toThrow('enhanced_untrusted_request')
    await expect(
      handlers.get(CODEX_RUNTIME_CHANNELS.startTurn)!(
        { sender: owner },
        Object.defineProperty({ documentId: 'doc' }, 'text', { get: () => 'attack' }),
      ),
    ).rejects.toThrow('enhanced_invalid_request')
    await expect(
      handlers.get(CODEX_RUNTIME_CHANNELS.startTurn)!(
        { sender: owner },
        Object.assign(Object.create({ documentId: 'doc' }), { text: 'attack' }),
      ),
    ).rejects.toThrow('enhanced_invalid_request')
  })
})
