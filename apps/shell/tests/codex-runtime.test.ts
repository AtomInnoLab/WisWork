import { describe, expect, it, vi } from 'vitest'
import { ENHANCED_HOSTS, type EnhancedRolloutPolicy } from '@wiswork/agent-runtime'
import { ShellCodexRuntime } from '../src/main/codex-runtime'

const policy = (enabled = true): EnhancedRolloutPolicy => ({
  globalEnabled: enabled,
  rawOfficeEnabled: false,
  hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as Record<
    (typeof ENHANCED_HOSTS)[number],
    boolean
  >,
})

function fixture(
  mode: 'standard' | 'enhanced' = 'enhanced',
  rollout: EnhancedRolloutPolicy = policy(),
) {
  let crash = () => undefined
  const engine = {
    startTurn: vi.fn(async () => undefined),
    cancelTurn: vi.fn(async () => undefined),
    closeDocument: vi.fn(async () => undefined),
    close: vi.fn(async () => undefined),
  }
  const bootstrap = {
    start: vi.fn(async ({ onCrash }: { onCrash: () => void }) => {
      crash = onCrash
      return engine
    }),
  }
  const runtime = new ShellCodexRuntime({
    activeAgentRuntime: mode,
    policy: rollout,
    isSignedIn: vi.fn(async () => true),
    resolveExecutable: vi.fn(async () => '/private/components/codex-app-server'),
    bootstrap,
  })
  const owner = { isDestroyed: () => false }
  return { runtime, owner, engine, bootstrap, crash: () => crash() }
}

describe('Shell Codex runtime lifecycle', () => {
  it('keeps Standard as an inert default and never starts Enhanced', async () => {
    const f = fixture('standard')
    await f.runtime.initialize()
    expect(f.runtime.state).toBe('standard')
    expect(f.bootstrap.start).not.toHaveBeenCalled()
  })

  it('starts only once at startup and binds turns to owner, document, host and generation', async () => {
    const f = fixture()
    await f.runtime.initialize()
    expect(f.runtime.state).toBe('ready')
    f.runtime.registerDocument({ owner: f.owner, documentId: 'doc-1', host: 'docs', generation: 3 })
    await f.runtime.startTurn(f.owner, 'doc-1', 'Edit this document')
    expect(f.engine.startTurn).toHaveBeenCalledWith({
      documentId: 'doc-1',
      host: 'docs',
      generation: 3,
      text: 'Edit this document',
    })
    await expect(
      f.runtime.startTurn({ isDestroyed: () => false }, 'doc-1', 'attack'),
    ).rejects.toThrow('enhanced_document_unavailable')
  })

  it('fails closed without fallback when startup or a turn fails', async () => {
    const bootstrap = { start: vi.fn(async () => Promise.reject(new Error('private'))) }
    const runtime = new ShellCodexRuntime({
      activeAgentRuntime: 'enhanced',
      policy: policy(),
      isSignedIn: async () => true,
      resolveExecutable: async () => '/private/codex',
      bootstrap,
    })
    await expect(runtime.initialize()).rejects.toThrow('enhanced_start_failed')
    expect(runtime.state).toBe('failed_safe')

    const f = fixture()
    f.engine.startTurn.mockRejectedValueOnce(new Error('crash'))
    await f.runtime.initialize()
    f.runtime.registerDocument({ owner: f.owner, documentId: 'doc', host: 'docs', generation: 0 })
    await expect(f.runtime.startTurn(f.owner, 'doc', 'request')).rejects.toThrow(
      'enhanced_turn_failed',
    )
    expect(f.engine.startTurn).toHaveBeenCalledTimes(1)
  })

  it('tears down document, crash, logout and quit state without restarting', async () => {
    const f = fixture()
    await f.runtime.initialize()
    f.runtime.registerDocument({ owner: f.owner, documentId: 'doc', host: 'docs', generation: 0 })
    f.crash()
    await vi.waitFor(() => expect(f.runtime.state).toBe('failed_safe'))
    await vi.waitFor(() => expect(f.engine.close).toHaveBeenCalledOnce())
    expect(f.bootstrap.start).toHaveBeenCalledOnce()

    const second = fixture()
    await second.runtime.initialize()
    second.runtime.registerDocument({
      owner: second.owner,
      documentId: 'doc',
      host: 'docs',
      generation: 0,
    })
    await second.runtime.logout()
    expect(second.engine.closeDocument).toHaveBeenCalledWith('doc')
    expect(second.engine.close).toHaveBeenCalledOnce()
  })

  it('fails closed when a host kill switch denies document registration', async () => {
    const hosts = Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as Record<
      (typeof ENHANCED_HOSTS)[number],
      boolean
    >
    hosts.slides = false
    const f = fixture('enhanced', { globalEnabled: true, rawOfficeEnabled: false, hosts })
    await f.runtime.initialize()
    expect(() =>
      f.runtime.registerDocument({
        owner: f.owner,
        documentId: 'deck',
        host: 'slides',
        generation: 0,
      }),
    ).toThrow('enhanced_document_unavailable')
  })
})
