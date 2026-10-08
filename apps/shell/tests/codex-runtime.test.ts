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
  const diagnosticTasks = {
    begin: vi.fn(() => 'diag_000000000000000000000000'),
    finish: vi.fn(),
  }
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
  const isSignedIn = vi.fn(async () => true)
  const runtime = new ShellCodexRuntime({
    activeAgentRuntime: mode,
    policy: rollout,
    isSignedIn,
    resolveExecutable: vi.fn(async () => '/private/components/codex-app-server'),
    bootstrap,
    diagnosticTasks,
  })
  const owner = { isDestroyed: () => false }
  return { runtime, owner, engine, bootstrap, diagnosticTasks, isSignedIn, crash: () => crash() }
}

describe('Shell Codex runtime lifecycle', () => {
  it('renews a real 15-minute Office statement without changing authority or generation', async () => {
    vi.useFakeTimers()
    try {
      const f = fixture()
      await f.runtime.initialize()
      const previous = f.runtime.createOfficeSessionStatement('office-word')!
      expect(previous.expires_at - Date.now()).toBe(15 * 60_000)
      // Another document issuing authority must not invalidate this document's lease.
      f.runtime.createOfficeSessionStatement('office-excel')
      await vi.advanceTimersByTimeAsync(10 * 60_000)
      const renewed = await f.runtime.renewOfficeSessionStatement(previous)
      expect(renewed).toEqual({ ...previous, expires_at: Date.now() + 15 * 60_000 })
      expect(Object.isFrozen(renewed)).toBe(true)
      expect(f.runtime.createOfficeSessionStatement('office-word')?.session_generation).toBe(3)
      await vi.advanceTimersByTimeAsync(5 * 60_000)
      await expect(f.runtime.renewOfficeSessionStatement(previous)).resolves.toBeUndefined()
      await expect(f.runtime.renewOfficeSessionStatement(renewed!)).resolves.toBeDefined()
    } finally {
      vi.useRealTimers()
    }
  })

  it.each(['logout', 'crash', 'policy', 'host', 'raw', 'signed-out', 'expiry'])(
    'cannot renew after %s while the async login check is pending',
    async (change) => {
      vi.useFakeTimers()
      try {
        const rollout = policy()
        const f = fixture('enhanced', rollout)
        await f.runtime.initialize()
        const previous = f.runtime.createOfficeSessionStatement('office-word')!
        let finish!: (signedIn: boolean) => void
        f.isSignedIn.mockImplementationOnce(() => new Promise((resolve) => (finish = resolve)))
        const renewal = f.runtime.renewOfficeSessionStatement(previous)
        if (change === 'logout') await f.runtime.logout()
        if (change === 'crash') f.crash()
        if (change === 'policy') Object.assign(rollout, { globalEnabled: false })
        if (change === 'host') Object.assign(rollout.hosts, { 'office-word': false })
        if (change === 'raw') Object.assign(rollout, { rawOfficeEnabled: true })
        if (change === 'expiry') await vi.advanceTimersByTimeAsync(15 * 60_000)
        finish(change !== 'signed-out')
        await expect(renewal).resolves.toBeUndefined()
      } finally {
        vi.useRealTimers()
      }
    },
  )

  it('rejects forged and replacement-runtime statements before awaiting login', async () => {
    const f = fixture()
    const replacement = fixture()
    await Promise.all([f.runtime.initialize(), replacement.runtime.initialize()])
    const previous = f.runtime.createOfficeSessionStatement('office-word')!
    f.isSignedIn.mockClear()
    replacement.isSignedIn.mockClear()
    await expect(f.runtime.renewOfficeSessionStatement({ ...previous })).resolves.toBeUndefined()
    await expect(replacement.runtime.renewOfficeSessionStatement(previous)).resolves.toBeUndefined()
    expect(f.isSignedIn).not.toHaveBeenCalled()
    expect(replacement.isSignedIn).not.toHaveBeenCalled()
  })

  it('issues restart-bound host-scoped Office statements only while Enhanced is ready', async () => {
    const f = fixture()
    expect(f.runtime.createOfficeSessionStatement('office-word', 1_000)).toBeUndefined()
    await f.runtime.initialize()
    const first = f.runtime.createOfficeSessionStatement('office-word', 1_000)!
    const second = f.runtime.createOfficeSessionStatement('office-word', 1_000)!
    expect(first).toMatchObject({
      runtime_mode: 'enhanced',
      component_version: '0.147.0',
      host: 'office-word',
      raw_office: false,
      expires_at: 901_000,
      session_generation: 1,
    })
    expect(second.runtime_instance).toBe(first.runtime_instance)
    expect(second.session_generation).toBe(2)
    await f.runtime.logout()
    expect(f.runtime.createOfficeSessionStatement('office-word', 1_000)).toBeUndefined()
  })

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

  it('allows an explicit retry after a transient Enhanced startup failure', async () => {
    const engine = {
      startTurn: vi.fn(async () => undefined),
      cancelTurn: vi.fn(async () => undefined),
      closeDocument: vi.fn(async () => undefined),
      close: vi.fn(async () => undefined),
    }
    const bootstrap = {
      start: vi
        .fn()
        .mockRejectedValueOnce(new Error('transient startup failure'))
        .mockResolvedValueOnce(engine),
    }
    const runtime = new ShellCodexRuntime({
      activeAgentRuntime: 'enhanced',
      policy: policy(),
      isSignedIn: async () => true,
      resolveExecutable: async () => '/private/codex',
      bootstrap,
    })

    await expect(runtime.initialize()).rejects.toThrow('enhanced_start_failed')
    await expect(runtime.initialize()).resolves.toBeUndefined()

    expect(runtime.state).toBe('ready')
    expect(bootstrap.start).toHaveBeenCalledTimes(2)
  })

  it('preserves actionable bounded turn failures for the renderer', async () => {
    const f = fixture()
    f.engine.startTurn.mockRejectedValueOnce(new Error('enhanced_turn_timeout'))
    await f.runtime.initialize()
    f.runtime.registerDocument({
      owner: f.owner,
      documentId: 'deck',
      host: 'slides',
      generation: 0,
    })

    await expect(f.runtime.startTurn(f.owner, 'deck', 'Create slides')).rejects.toThrow(
      'enhanced_turn_timeout',
    )

    f.engine.startTurn.mockRejectedValueOnce(new Error('enhanced_service_unavailable'))
    await expect(f.runtime.startTurn(f.owner, 'deck', 'Try again')).rejects.toThrow(
      'enhanced_service_unavailable',
    )
    f.engine.startTurn.mockRejectedValueOnce(new Error('enhanced_proposal_expired'))
    await expect(f.runtime.startTurn(f.owner, 'deck', 'Expired proposal')).rejects.toThrow(
      'enhanced_proposal_expired',
    )
    expect(f.diagnosticTasks.begin).toHaveBeenNthCalledWith(1, 'slides')
    expect(f.diagnosticTasks.finish).toHaveBeenNthCalledWith(
      1,
      'diag_000000000000000000000000',
      'failed',
      'enhanced_turn_timeout',
    )
  })

  it('keeps a document busy until the engine reports terminal settlement', async () => {
    const f = fixture()
    let settle!: () => void
    f.engine.startTurn.mockImplementationOnce(
      () => new Promise<void>((resolve) => (settle = resolve)),
    )
    await f.runtime.initialize()
    f.runtime.registerDocument({ owner: f.owner, documentId: 'doc', host: 'docs', generation: 0 })
    const first = f.runtime.startTurn(f.owner, 'doc', 'first')
    await vi.waitFor(() => expect(f.engine.startTurn).toHaveBeenCalledOnce())
    await expect(f.runtime.startTurn(f.owner, 'doc', 'second')).rejects.toThrow(
      'enhanced_turn_in_progress',
    )
    await f.runtime.cancelTurn(f.owner, 'doc')
    expect(f.engine.cancelTurn).toHaveBeenCalledWith('doc')
    expect(f.diagnosticTasks.finish).toHaveBeenCalledWith(
      'diag_000000000000000000000000',
      'cancelled',
    )
    settle()
    await first
    await expect(f.runtime.startTurn(f.owner, 'doc', 'third')).resolves.toBeUndefined()
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

  it('sets the signed raw Office flag only from the independent trusted policy', async () => {
    const hosts = Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as Record<
      (typeof ENHANCED_HOSTS)[number],
      boolean
    >
    const f = fixture('enhanced', { globalEnabled: true, rawOfficeEnabled: true, hosts })
    await f.runtime.initialize()
    expect(f.runtime.createOfficeSessionStatement('office-word')).toMatchObject({
      host: 'office-word',
      raw_office: true,
    })
  })

  it('rejects late initialization and closes a stale engine after shutdown', async () => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    const engine = {
      startTurn: vi.fn(async () => undefined),
      cancelTurn: vi.fn(async () => undefined),
      closeDocument: vi.fn(async () => undefined),
      close: vi.fn(async () => undefined),
    }
    const bootstrap = { start: vi.fn(async () => (await gate, engine)) }
    const runtime = new ShellCodexRuntime({
      activeAgentRuntime: 'enhanced',
      policy: policy(),
      isSignedIn: async () => true,
      resolveExecutable: async () => '/private/codex',
      bootstrap,
    })
    const initializing = runtime.initialize()
    await vi.waitFor(() => expect(bootstrap.start).toHaveBeenCalledOnce())
    const closing = runtime.shutdown()
    release()
    await expect(initializing).rejects.toThrow('enhanced_runtime_closed')
    await closing
    expect(engine.close).toHaveBeenCalledOnce()
    expect(runtime.activeAgentRuntime).toBe('standard')
    expect(runtime.state).toBe('unavailable')
  })
})
