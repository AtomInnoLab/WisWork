import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ENHANCED_HOSTS, type EnhancedRolloutPolicy } from '@wiswork/agent-runtime'
import { ShellCodexRuntime } from '../src/main/codex-runtime'
import { createOfficeRelayClient, type RelaySocket } from '../src/main/office-relay-client'

class Socket implements RelaySocket {
  readyState = 0
  sent: Array<Record<string, any>> = []
  listeners = new Map<string, Array<(event: any) => void>>()
  addEventListener(name: string, listener: (event: any) => void) {
    this.listeners.set(name, [...(this.listeners.get(name) ?? []), listener])
  }
  send(data: string) {
    this.sent.push(JSON.parse(data))
  }
  emit(name: string, value: unknown = {}) {
    for (const listener of this.listeners.get(name) ?? []) listener(value)
  }
  close() {
    this.readyState = 3
    this.emit('close')
  }
  message(value: object) {
    this.emit('message', { data: JSON.stringify(value) })
  }
}

async function fixture(negotiated = true, callback = true, resume = false, freshness = true) {
  const account = { loggedIn: true, userId: 'paired-account' }
  const policy: EnhancedRolloutPolicy = {
    globalEnabled: true,
    rawOfficeEnabled: false,
    hosts: Object.fromEntries(
      ENHANCED_HOSTS.map((host) => [host, true]),
    ) as EnhancedRolloutPolicy['hosts'],
  }
  let crash = () => {}
  const runtime = new ShellCodexRuntime({
    activeAgentRuntime: 'enhanced',
    policy,
    isSignedIn: async () => account.loggedIn,
    resolveExecutable: async () => '/private/components/codex',
    bootstrap: {
      start: async ({ onCrash }) => {
        crash = onCrash
        return {
          startTurn: async () => undefined,
          cancelTurn: async () => undefined,
          closeDocument: async () => undefined,
          close: async () => undefined,
        }
      },
    },
  })
  await runtime.initialize()
  const socket = new Socket()
  const enhancedProxy = vi.fn(
    async (
      request: Parameters<
        NonNullable<Parameters<typeof createOfficeRelayClient>[0]['enhancedProxy']>
      >[0],
    ) => ({
      status: 200,
      body: (async function* () {
        for (const callId of ['call_once_12345678', 'call_next_12345678']) {
          await request.executeTool({
            turnId: 'turn_12345678',
            callId,
            generation: request.statement.session_generation,
            toolName: 'write_document',
            input: { approvedOnce: true },
          })
        }
        yield new TextEncoder().encode('done')
      })(),
    }),
  )
  const renew = vi.fn((previous: Parameters<ShellCodexRuntime['renewOfficeSessionStatement']>[0]) =>
    runtime.renewOfficeSessionStatement(previous),
  )
  const getValidAccountStatus = vi.fn(async () => ({ ...account }))
  const isEnhancedStatementCurrent = vi.fn(
    (statement: Parameters<ShellCodexRuntime['renewOfficeSessionStatement']>[0]) =>
      runtime.isOfficeSessionStatementCurrent(statement),
  )
  const retrieval = vi.fn(async () => new Uint8Array())
  const client = createOfficeRelayClient({
    endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
    connect: () => socket,
    getValidAccountStatus,
    getAccessToken: async () => 'token',
    proxy: async () => {
      throw new Error('standard_must_not_run')
    },
    enhancedProxy,
    enhancedStatement: () => runtime.createOfficeSessionStatement('office-word'),
    ...(callback ? { renewEnhancedStatement: renew } : {}),
    ...(freshness ? { isEnhancedStatementCurrent } : {}),
    retrievalProxy: retrieval,
    negotiateCapabilities: true,
    onPending() {},
  })
  const capabilities = ['agent.v1', ...(negotiated ? ['enhanced-lease.v1'] : [])]
  const opening = resume
    ? client.resume({
        accountId: account.userId,
        bindingId: 'binding_12345678',
        host: 'Word',
        origin: 'https://office.8-216-134-194.sslip.io',
        capabilities,
        createdAt: Date.now(),
      })
    : client.claim('123456')
  await vi.advanceTimersByTimeAsync(0)
  socket.readyState = 1
  socket.emit('open')
  await opening
  if (!resume) {
    socket.message({ version: 2, type: 'pc.negotiated', pairing_version: 2, capabilities })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities,
    })
    await client.approve('pairing_12345678')
  }
  socket.message({
    version: 2,
    type: 'pc.approved',
    session_id: 'session_12345678',
    capability: 'secret_12345678',
    expires_in: 1800,
    capabilities,
  })
  const states = () => socket.sent.filter((frame) => frame.type === 'pc.session_state')
  const initial = states()[0]?.enhanced
  const request = (capability_name = 'agent.v1') =>
    socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      capability_name,
      body: { messages: [] },
    })
  return {
    account,
    policy,
    runtime,
    socket,
    client,
    renew,
    initial,
    states,
    request,
    enhancedProxy,
    retrieval,
    getValidAccountStatus,
    isEnhancedStatementCurrent,
    crash: () => crash(),
  }
}

describe('PC Enhanced authorization leases', () => {
  beforeEach(() => {
    vi.useFakeTimers()
    vi.setSystemTime(1_700_000_000_000)
  })
  afterEach(() => {
    vi.clearAllTimers()
    vi.useRealTimers()
  })

  it.each(['crash', 'logout', 'global-policy', 'host-policy', 'raw-policy'])(
    'rechecks current runtime authority after %s during the final account await',
    async (change) => {
      const f = await fixture()
      f.request()
      let finish!: (value: { loggedIn: boolean; userId: string }) => void
      f.getValidAccountStatus.mockResolvedValueOnce({ ...f.account })
      f.getValidAccountStatus.mockImplementationOnce(
        () => new Promise((resolve) => (finish = resolve)),
      )
      await vi.advanceTimersByTimeAsync(10 * 60_000)
      expect(f.renew).toHaveBeenCalledOnce()
      if (change === 'crash') f.crash()
      if (change === 'logout') await f.runtime.logout()
      if (change === 'global-policy') Object.assign(f.policy, { globalEnabled: false })
      if (change === 'host-policy') Object.assign(f.policy.hosts, { 'office-word': false })
      if (change === 'raw-policy') Object.assign(f.policy, { rawOfficeEnabled: true })
      finish({ ...f.account })
      await vi.advanceTimersByTimeAsync(0)
      expect(f.states()).toHaveLength(1)
      expect(f.client.status()).toBe('disconnected:enhanced_authority_unavailable')
      expect(f.enhancedProxy.mock.calls[0]![0].signal.aborted).toBe(true)
    },
  )

  it.each([false, true])(
    'requires synchronous authority validation to offer or renew leases (resume=%s)',
    async (resume) => {
      const f = await fixture(resume, true, resume, false)
      if (!resume) expect(f.socket.sent[0]!.capabilities).not.toContain('enhanced-lease.v1')
      await vi.advanceTimersByTimeAsync(15 * 60_000)
      expect(f.renew).not.toHaveBeenCalled()
      expect(f.client.status()).toBe('disconnected:session_expired')
    },
  )

  it('renews a real 15-minute lease while an approved-once write and original proxy generation remain active', async () => {
    const f = await fixture()
    expect(f.initial.expires_at - Date.now()).toBe(15 * 60_000)
    f.request()
    await vi.advanceTimersByTimeAsync(0)
    const active = f.enhancedProxy.mock.calls[0]![0]
    await vi.advanceTimersByTimeAsync(10 * 60_000)
    expect(f.renew).toHaveBeenCalledOnce()
    expect(f.states()).toHaveLength(2)
    expect(f.states()[1]!.enhanced).toEqual({ ...f.initial, expires_at: Date.now() + 15 * 60_000 })
    await vi.advanceTimersByTimeAsync(6 * 60_000)
    expect(active.signal.aborted).toBe(false)
    expect(f.socket.sent.filter((frame) => frame.type === 'pc.tool_call')).toHaveLength(1)
    for (const callId of ['call_once_12345678', 'call_next_12345678']) {
      f.socket.message({
        version: 2,
        type: 'relay.tool_result',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        turn_id: 'turn_12345678',
        call_id: callId,
        generation: f.initial.session_generation,
        output: 'done',
        is_error: false,
      })
      await vi.advanceTimersByTimeAsync(0)
    }
    expect(f.socket.sent.filter((frame) => frame.type === 'pc.tool_call')).toHaveLength(2)
    expect(f.socket.sent.some((frame) => frame.type === 'pc.done')).toBe(true)
    expect(f.enhancedProxy).toHaveBeenCalledOnce()
    f.client.revoke()
  })

  it.each([false, true])(
    'does not renew or broaden an unnegotiated binding (resume=%s)',
    async (resume) => {
      const f = await fixture(false, true, resume)
      f.request()
      await vi.advanceTimersByTimeAsync(15 * 60_000)
      expect(f.renew).not.toHaveBeenCalled()
      expect(f.states()).toHaveLength(1)
      expect(f.enhancedProxy.mock.calls[0]![0].signal.aborted).toBe(true)
      if (resume) expect(f.socket.sent[0]!.capabilities).toEqual(['agent.v1'])
    },
  )

  it('offers the control capability only with a renewal callback and never as retrieval', async () => {
    const without = await fixture(false, false)
    expect(without.socket.sent[0]!.capabilities).not.toContain('enhanced-lease.v1')
    without.client.revoke()
    const f = await fixture()
    expect(f.socket.sent[0]!.capabilities).toContain('enhanced-lease.v1')
    f.request('enhanced-lease.v1')
    await vi.advanceTimersByTimeAsync(0)
    expect(f.retrieval).not.toHaveBeenCalled()
    expect(f.client.status()).toBe('disconnected:protocol_violation')
  })

  it.each(['logout', 'account', 'policy', 'revoke', 'close', 'timeout', 'expiry'])(
    'rejects a late renewal after %s and cannot revive active work',
    async (ending) => {
      const f = await fixture()
      f.request()
      let finish!: (value: any) => void
      f.renew.mockImplementationOnce(() => new Promise((resolve) => (finish = resolve)))
      await vi.advanceTimersByTimeAsync(10 * 60_000)
      expect(f.renew).toHaveBeenCalledOnce()
      if (ending === 'logout') {
        f.account.loggedIn = false
        await f.runtime.logout()
      }
      if (ending === 'account') f.account.userId = 'different-account'
      if (ending === 'policy') Object.assign(f.policy, { globalEnabled: false })
      if (ending === 'revoke') f.client.revoke()
      if (ending === 'close') f.socket.close()
      if (ending === 'timeout') await vi.advanceTimersByTimeAsync(10_001)
      if (ending === 'expiry') await vi.advanceTimersByTimeAsync(5 * 60_000)
      finish(
        ending === 'policy' ? undefined : { ...f.initial, expires_at: Date.now() + 15 * 60_000 },
      )
      await vi.advanceTimersByTimeAsync(0)
      expect(f.states()).toHaveLength(1)
      await vi.advanceTimersByTimeAsync(5 * 60_000)
      expect(f.renew).toHaveBeenCalledOnce()
      expect(f.enhancedProxy.mock.calls[0]![0].signal.aborted).toBe(true)
    },
  )

  it('bounds periodic renewal without extending the active request deadline', async () => {
    const f = await fixture()
    f.request()
    await vi.advanceTimersByTimeAsync(30 * 60_000 + 25_000)
    expect(f.renew).toHaveBeenCalledTimes(3)
    expect(f.enhancedProxy.mock.calls[0]![0].signal.aborted).toBe(true)
    f.client.revoke()
    await vi.advanceTimersByTimeAsync(60 * 60_000)
    expect(f.renew).toHaveBeenCalledTimes(3)
  })

  it('does not replay a cancelled write when a pending renewal subsequently succeeds', async () => {
    const f = await fixture()
    f.request()
    let finish!: (value: any) => void
    f.renew.mockImplementationOnce(() => new Promise((resolve) => (finish = resolve)))
    await vi.advanceTimersByTimeAsync(10 * 60_000)
    f.socket.message({
      version: 2,
      type: 'relay.cancel',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
    })
    finish(await f.runtime.renewOfficeSessionStatement(f.renew.mock.calls[0]![0]))
    await vi.advanceTimersByTimeAsync(0)
    expect(f.states()).toHaveLength(2)
    expect(f.enhancedProxy.mock.calls[0]![0].signal.aborted).toBe(true)
    expect(f.socket.sent.filter((frame) => frame.type === 'pc.tool_call')).toHaveLength(1)
    expect(f.socket.sent.some((frame) => frame.type === 'pc.done')).toBe(false)
    f.client.revoke()
  })

  it('keeps the original expiry when renewal rejects without retrying', async () => {
    const f = await fixture()
    f.request()
    f.renew.mockRejectedValueOnce(new Error('unavailable'))
    await vi.advanceTimersByTimeAsync(15 * 60_000)
    expect(f.states()).toHaveLength(1)
    expect(f.renew).toHaveBeenCalledOnce()
    expect(f.client.status()).toBe('disconnected:session_expired')
    expect(f.enhancedProxy.mock.calls[0]![0].signal.aborted).toBe(true)
  })

  it.each(['before-runtime', 'after-runtime'])(
    'fences account validation that stalls %s and settles after revoke',
    async (stage) => {
      const f = await fixture()
      let finish!: (value: { loggedIn: boolean; userId: string }) => void
      if (stage === 'after-runtime') f.getValidAccountStatus.mockResolvedValueOnce({ ...f.account })
      f.getValidAccountStatus.mockImplementationOnce(
        () => new Promise((resolve) => (finish = resolve)),
      )
      await vi.advanceTimersByTimeAsync(10 * 60_000)
      expect(f.renew).toHaveBeenCalledTimes(stage === 'after-runtime' ? 1 : 0)
      f.client.revoke()
      finish({ ...f.account })
      await vi.advanceTimersByTimeAsync(15 * 60_000)
      expect(f.states()).toHaveLength(1)
      expect(f.client.status()).toBe('disconnected:revoked')
    },
  )

  it.each([
    { raw_office: true },
    { runtime_instance: 'runtime_other_12345678' },
    { host: 'office-excel' },
    { policy_generation: 100 },
    { session_generation: 100 },
    { component_version: '1.0.0' },
    { extra: true },
    { expires_at: 1_700_000_000_000 },
    { expires_at: 1_700_000_000_000 + 15 * 60_000 },
    { expires_at: 1_700_000_000_000 + 25 * 60_000 + 1 },
  ])(
    'rejects a callback that changes authority or supplies invalid expiry: %j',
    async (changes) => {
      const f = await fixture()
      f.renew.mockImplementationOnce(async () => ({
        ...f.initial,
        expires_at: Date.now() + 15 * 60_000,
        ...changes,
      }))
      await vi.advanceTimersByTimeAsync(10 * 60_000)
      expect(f.states()).toHaveLength(1)
      expect(f.client.status()).toBe('disconnected:protocol_violation')
    },
  )

  it('does not busy-loop when a callback advances expiry by only one millisecond', async () => {
    const f = await fixture()
    f.isEnhancedStatementCurrent.mockReturnValue(true)
    f.renew.mockImplementation(async (previous) => ({
      ...previous,
      expires_at: previous.expires_at + 1,
    }))
    await vi.advanceTimersByTimeAsync(10 * 60_000 + 1_000)
    expect(f.renew).toHaveBeenCalledOnce()
    expect(f.states()).toHaveLength(2)
    f.client.revoke()
  })
})
