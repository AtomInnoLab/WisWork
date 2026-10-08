import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationStore, PresentationLifecycleStore } from '@wiswork/project-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { acquirePresentationProjectLock } from '../src/main/presentation-service'
import { describe, expect, it, vi } from 'vitest'
import {
  createOfficeRelayClient,
  createOfficePresentationGovernanceProxy,
  officeRelayEndpointFromEnv,
  type RelaySocket,
} from '../src/main/office-relay-client'
import type { OfficeRelayBinding } from '../src/main/office-relay-binding-store'
import { createOfficeLocalSearchProxy } from '../src/main/office-retrieval-proxy'

class FakeSocket implements RelaySocket {
  readyState = 0
  sent: string[] = []
  closedWith?: { code?: number; reason?: string }
  pings = 0
  listeners = new Map<string, Array<(event: any) => void>>()
  addEventListener(name: string, listener: (event: any) => void): void {
    this.listeners.set(name, [...(this.listeners.get(name) ?? []), listener])
  }
  onUnexpectedResponse(listener: (statusCode: number | undefined) => void): void {
    this.listeners.set('unexpected-response', [
      ...(this.listeners.get('unexpected-response') ?? []),
      listener,
    ])
  }
  send(data: string): void {
    this.sent.push(data)
  }
  ping(): void {
    this.pings += 1
  }
  close(code?: number, reason?: string): void {
    this.closedWith = { code, reason }
    this.readyState = 3
    this.emit('close', {})
  }
  open(): void {
    this.readyState = 1
    this.emit('open', {})
  }
  message(value: object): void {
    this.emit('message', { data: JSON.stringify(value) })
  }
  emit(name: string, event: any): void {
    for (const listener of this.listeners.get(name) ?? []) listener(event)
  }
}

function setup(loggedIn = true) {
  const socket = new FakeSocket()
  const pending = vi.fn()
  const proxy = vi.fn(async () => ({
    status: 200,
    contentType: 'text/event-stream',
    body: (async function* () {
      yield new TextEncoder().encode('hello')
    })(),
  }))
  const client = createOfficeRelayClient({
    endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
    connect: () => socket,
    getValidAccountStatus: async () => ({ loggedIn }),
    getAccessToken: async () => (loggedIn ? 'access-token' : null),
    proxy,
    onPending: pending,
  })
  return { client, socket, pending, proxy }
}

describe('Office relay PC client', () => {
  it.each(['agent.v1', 'presentation-master-backups.v1', 'presentation-package-backups.v1'])(
    'reattaches an approved v2 %s session with a fresh token after socket loss',
    async (capabilityName) => {
      const first = new FakeSocket()
      const second = new FakeSocket()
      const connect = vi.fn().mockReturnValueOnce(first).mockReturnValueOnce(second)
      const getAccessToken = vi
        .fn()
        .mockResolvedValueOnce('first-token')
        .mockResolvedValueOnce('second-token')
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect,
        getValidAccountStatus: async () => ({ loggedIn: true }),
        getAccessToken,
        proxy: async () => ({ status: 200, body: new Uint8Array() }),
        negotiateCapabilities: true,
        presentationProxy: async () => new Uint8Array(),
        onPending() {},
      })
      const claiming = client.claim('123456')
      await vi.waitFor(() => expect(connect).toHaveBeenCalledTimes(1))
      first.open()
      await claiming
      first.message({
        version: 2,
        type: 'pc.negotiated',
        pairing_version: 2,
        capabilities: [capabilityName],
      })
      first.message({
        version: 2,
        type: 'pc.claimed',
        pairing_id: 'pairing_12345678',
        host: 'PowerPoint',
        origin: 'https://office.8-216-134-194.sslip.io',
        verification_code: '123456',
        expires_in: 120,
        capabilities: [capabilityName],
      })
      await client.approve('pairing_12345678')
      first.message({
        version: 2,
        type: 'pc.approved',
        session_id: 'session_12345678',
        capability: 'secret-capability',
        expires_in: 1800,
        capabilities: [capabilityName],
      })
      expect(client.status()).toBe('paired')
      first.close()
      await vi.waitFor(() => expect(connect).toHaveBeenCalledTimes(2))
      expect(client.status()).toBe('connecting')
      expect(connect).toHaveBeenNthCalledWith(2, expect.any(String), 'second-token')
      second.open()
      await vi.waitFor(() => expect(second.sent).toHaveLength(1))
      expect(JSON.parse(second.sent[0]!)).toEqual({
        version: 2,
        type: 'pc.resume',
        session_id: 'session_12345678',
        capability: 'secret-capability',
      })
      second.message({
        version: 2,
        type: 'pc.resumed',
        session_id: 'session_12345678',
        expires_in: 1800,
        capabilities: [capabilityName],
      })
      expect(client.status()).toBe('paired')
      expect(second.pings).toBeGreaterThan(0)
      client.revoke('test_complete')
      expect(second.closedWith).toEqual({ code: 1000, reason: 'session_revoked' })
    },
  )
  it.each(['before-token', 'during-token'])(
    'refuses session reattachment across account changes %s',
    async (changeAt) => {
      const capabilityName = 'agent.v1'
      let accountId = 'original-account'
      const first = new FakeSocket()
      const second = new FakeSocket()
      const connect = vi.fn().mockReturnValueOnce(first).mockReturnValueOnce(second)
      const getAccessToken = vi
        .fn()
        .mockResolvedValueOnce('first-token')
        .mockImplementationOnce(async () => {
          if (changeAt === 'during-token') accountId = 'replacement-account'
          return 'second-token'
        })
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect,
        getValidAccountStatus: async () => ({ loggedIn: true, userId: accountId }),
        getAccessToken,
        proxy: async () => ({ status: 200, body: new Uint8Array() }),
        negotiateCapabilities: true,
        presentationProxy: async () => new Uint8Array(),
        onPending() {},
      })
      const claiming = client.claim('123456')
      await vi.waitFor(() => expect(connect).toHaveBeenCalledTimes(1))
      first.open()
      await claiming
      first.message({
        version: 2,
        type: 'pc.negotiated',
        pairing_version: 2,
        capabilities: [capabilityName],
      })
      first.message({
        version: 2,
        type: 'pc.claimed',
        pairing_id: 'pairing_12345678',
        host: 'PowerPoint',
        origin: 'https://office.8-216-134-194.sslip.io',
        verification_code: '123456',
        expires_in: 120,
        capabilities: [capabilityName],
      })
      await client.approve('pairing_12345678')
      first.message({
        version: 2,
        type: 'pc.approved',
        session_id: 'session_12345678',
        capability: 'secret-capability',
        expires_in: 1800,
        capabilities: [capabilityName],
      })
      expect(client.status()).toBe('paired')
      if (changeAt === 'before-token') accountId = 'replacement-account'
      first.close()
      await vi.waitFor(() => expect(client.status()).toBe('disconnected:auth_required'))
      expect(connect).toHaveBeenCalledTimes(1)
      expect(second.sent).toEqual([])
      client.revoke('test_complete')
    },
  )
  it('refreshes once and reconnects when a websocket handshake rejects a stale token', async () => {
    const sockets: FakeSocket[] = []
    const tokens: string[] = []
    const refreshAccessToken = vi.fn(async () => 'fresh-token')
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: (_url, token) => {
        tokens.push(token)
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'stale-token',
      refreshAccessToken,
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      onPending() {},
    })

    const claim = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(1))
    sockets[0]!.emit('unexpected-response', 401)
    await vi.waitFor(() => expect(sockets).toHaveLength(2))
    sockets[1]!.open()
    await expect(claim).resolves.toBeUndefined()
    expect(tokens).toEqual(['stale-token', 'fresh-token'])
    expect(refreshAccessToken).toHaveBeenCalledOnce()
  })

  it('actively keeps the PC websocket alive and stops heartbeats after close', async () => {
    const socket = new FakeSocket()
    let heartbeat: (() => void) | undefined
    const clearHeartbeat = vi.fn()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      onPending() {},
      scheduleHeartbeat(callback, delay) {
        expect(delay).toBe(20_000)
        heartbeat = callback
        return 17 as unknown as ReturnType<typeof setInterval>
      },
      clearHeartbeat,
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    expect(socket.pings).toBe(1)
    heartbeat?.()
    expect(socket.pings).toBe(2)
    socket.close()
    expect(clearHeartbeat).toHaveBeenCalledWith(17)
    heartbeat?.()
    expect(socket.pings).toBe(2)
  })

  it.each(['revoked', 'socket-close'])(
    'clears only its own client image provenance when %s',
    async (ending) => {
      const makeProxy = () =>
        createOfficeLocalSearchProxy({
          fetchWithAuth: vi.fn(),
          downloadImage: async () => ({ mime: 'image/png' as const, bytes: new Uint8Array(4) }),
          searchImages: async () => ({
            images: [
              {
                title: 'Cover',
                imageUrl: 'https://images.example/cover.png',
                sourceUrl: 'https://example.com/cover',
                source: 'example.com',
              },
            ],
            method: 'serpapi',
          }),
        })
      const first = makeProxy(),
        second = makeProxy()
      const socket = new FakeSocket()
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect: () => socket,
        getValidAccountStatus: async () => ({ loggedIn: true }),
        getAccessToken: async () => 'token',
        proxy: async () => ({ status: 200, body: new Uint8Array() }),
        retrievalProxy: first,
        onPending() {},
      })
      const claiming = client.claim('123456')
      await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
      socket.open()
      await claiming
      await first('image-search.v1', { query: 'private cover query', max_results: 1 })
      await expect(
        second('image-fetch.v1', { url: 'https://images.example/cover.png' }),
      ).rejects.toThrow('retrieval_invalid_request')
      await second('image-search.v1', { query: 'separate cover query', max_results: 1 })
      if (ending === 'socket-close') socket.close()
      else client.revoke()
      await expect(
        first('image-fetch.v1', { url: 'https://images.example/cover.png' }),
      ).rejects.toThrow('retrieval_invalid_request')
      await expect(
        second('image-fetch.v1', { url: 'https://images.example/cover.png' }),
      ).resolves.toBeInstanceOf(Uint8Array)
      client.revoke('test-complete')
    },
  )

  it.each([
    { delayMs: 0, imageBytes: 0, toolName: 'read_document', reply: 'small' },
    { delayMs: 360_000, imageBytes: 0, toolName: 'read_document', reply: 'small' },
    { delayMs: 0, imageBytes: 180 * 1024, toolName: 'read_document', reply: 'small' },
    { delayMs: 0, imageBytes: 200 * 1024, toolName: 'read_document', reply: 'small' },
    { delayMs: 0, imageBytes: 0, toolName: 'plan_deck', reply: 'design' },
    { delayMs: 0, imageBytes: 0, toolName: 'screenshot_slide', reply: 'design' },
    { delayMs: 0, imageBytes: 0, toolName: 'plan_deck', reply: 'boundary' },
    { delayMs: 0, imageBytes: 0, toolName: 'plan_deck', reply: 'oversized' },
    { delayMs: 0, imageBytes: 0, toolName: 'plan_deck', reply: 'wrong-call' },
    { delayMs: 0, imageBytes: 0, toolName: 'plan_deck', reply: 'extra-field' },
    { delayMs: 0, imageBytes: 0, toolName: 'plan_deck', reply: 'control' },
  ])('handles bounded $toolName $reply ($delayMs/$imageBytes)', async (testCase) => {
    const { delayMs, imageBytes, toolName, reply } = testCase
    vi.useFakeTimers()
    const socket = new FakeSocket()
    const responses: string[] = []
    const output =
      reply === 'small'
        ? '{"title":"Doc"}'
        : reply === 'boundary'
          ? 'x'.repeat(256 * 1024)
          : reply === 'oversized'
            ? 'x'.repeat(272 * 1024)
            : JSON.stringify({
                designMd: '# DESIGN.md\n' + 'a'.repeat(32 * 1024),
                ...(toolName === 'screenshot_slide'
                  ? { mime: 'image/png', bytes: 100, fingerprint: 'shot' }
                  : {}),
              })
    let runtimeReady = false
    const enhanced = {
      version: 1,
      runtime_mode: 'enhanced',
      runtime_instance: 'runtime_0123456789abcdef',
      component_version: '0.147.0',
      host: 'office-word',
      raw_office: false,
      expires_at: Date.now() + 40 * 60_000,
      policy_generation: 2,
      session_generation: 7,
    } as const
    const enhancedProxy = vi.fn(async ({ executeTool }) => ({
      status: 200,
      contentType: 'text/event-stream',
      body: (async function* () {
        const imageInput = imageBytes
          ? { _wiswork_image_base64: Buffer.alloc(imageBytes).toString('base64') }
          : {}
        if (imageBytes > 180 * 1024)
          await expect(
            executeTool({
              turnId: 'turn_12345678',
              callId: 'call_oversized',
              generation: 7,
              toolName: 'insert_web_image',
              input: imageInput,
            }),
          ).rejects.toThrow('invalid_tool_call')
        const result = await executeTool({
          turnId: 'turn_12345678',
          callId: 'call_12345678',
          generation: 7,
          toolName,
          input: imageBytes > 180 * 1024 ? {} : imageInput,
        })
        responses.push(result.output)
        if (reply === 'design') {
          const next = await executeTool({
            turnId: 'turn_12345678',
            callId: 'call_next_12345678',
            generation: 7,
            toolName: 'read_document',
            input: {},
          })
          responses.push(next.output)
        }
        yield new TextEncoder().encode('done')
      })(),
    }))
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy: async () => {
        throw new Error('standard_must_not_run')
      },
      enhancedProxy,
      enhancedStatement: () => (runtimeReady ? enhanced : undefined),
      negotiateCapabilities: true,
      onPending() {},
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
    })
    await vi.waitFor(() =>
      expect(socket.sent.map(JSON.parse).some((value) => value.type === 'pc.session_state')).toBe(
        true,
      ),
    )
    expect(socket.sent.map(JSON.parse).find((value) => value.type === 'pc.session_state')).toEqual({
      version: 2,
      type: 'pc.session_state',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      generation: 0,
      enhanced: null,
    })
    runtimeReady = true
    socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      capability_name: 'agent.v1',
      body: { messages: [] },
    })
    await vi.waitFor(() =>
      expect(
        socket.sent
          .map(JSON.parse)
          .filter((value) => value.type === 'pc.session_state')
          .at(-1),
      ).toMatchObject({ generation: 7, enhanced }),
    )
    await vi.waitFor(() =>
      expect(socket.sent.map(JSON.parse).some((value) => value.type === 'pc.tool_call')).toBe(true),
    )
    expect(
      socket.sent.map(JSON.parse).find((value) => value.type === 'pc.tool_call'),
    ).toMatchObject({
      request_id: 'request_12345678',
      turn_id: 'turn_12345678',
      call_id: 'call_12345678',
      generation: 7,
      tool_name: toolName,
    })
    await vi.advanceTimersByTimeAsync(delayMs)
    socket.message(
      reply === 'control'
        ? {
            version: 2,
            type: 'relay.cancel',
            session_id: 'session_12345678',
            request_id: 'x'.repeat(32 * 1024),
          }
        : {
            version: 2,
            type: 'relay.tool_result',
            session_id: 'session_12345678',
            request_id: 'request_12345678',
            turn_id: 'turn_12345678',
            call_id: reply === 'wrong-call' ? 'call_wrong_12345678' : 'call_12345678',
            generation: 7,
            output,
            is_error: false,
            ...(reply === 'extra-field' ? { unexpected: true } : {}),
          },
    )
    if (['oversized', 'wrong-call', 'extra-field', 'control'].includes(reply)) {
      expect(client.status()).toBe('disconnected:protocol_violation')
      expect(responses).toEqual([])
      client.revoke()
      vi.useRealTimers()
      return
    }
    expect(client.status()).toBe('paired')
    if (reply === 'design') {
      await vi.waitFor(() =>
        expect(
          socket.sent
            .map(JSON.parse)
            .some(
              (frame) => frame.type === 'pc.tool_call' && frame.call_id === 'call_next_12345678',
            ),
        ).toBe(true),
      )
      socket.message({
        version: 2,
        type: 'relay.tool_result',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        turn_id: 'turn_12345678',
        call_id: 'call_next_12345678',
        generation: 7,
        output: 'next tool succeeded',
        is_error: false,
      })
    }
    await vi.waitFor(() =>
      expect(socket.sent.map(JSON.parse).some((value) => value.type === 'pc.done')).toBe(true),
    )
    expect(enhancedProxy).toHaveBeenCalledOnce()
    expect(responses).toEqual(reply === 'design' ? [output, 'next tool succeeded'] : [output])
    client.revoke()
    vi.useRealTimers()
  })

  it.each(['account', 'token'] as const)(
    'does not connect when revoked while awaiting %s validation',
    async (phase) => {
      let releaseAccount!: (value: { loggedIn: boolean }) => void
      let releaseToken!: (value: string | null) => void
      const account = new Promise<{ loggedIn: boolean }>((resolve) => {
        releaseAccount = resolve
      })
      const token = new Promise<string | null>((resolve) => {
        releaseToken = resolve
      })
      const connect = vi.fn(() => new FakeSocket())
      const getAccessToken = vi.fn(() => token)
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect,
        getValidAccountStatus: () => account,
        getAccessToken,
        proxy: async () => ({ status: 200, body: new Uint8Array() }),
        onPending() {},
      })

      const claim = client.claim('123456')
      if (phase === 'token') {
        releaseAccount({ loggedIn: true })
        await vi.waitFor(() => expect(getAccessToken).toHaveBeenCalledOnce())
      }
      client.revoke('logout')
      releaseAccount({ loggedIn: true })
      releaseToken('access-token')

      await expect(claim).rejects.toThrow('relay_connection_failed')
      expect(connect).not.toHaveBeenCalled()
      expect(client.status()).toBe('disconnected:logout')
    },
  )

  it('accepts only secure relay endpoint configuration without embedded credentials', () => {
    expect(officeRelayEndpointFromEnv({})).toBe('wss://office.8-216-134-194.sslip.io/office-relay')
    expect(() =>
      officeRelayEndpointFromEnv({ WISWORK_OFFICE_RELAY_URL: 'ws://localhost/relay' }),
    ).toThrow('invalid_office_relay_url')
    expect(() =>
      officeRelayEndpointFromEnv({ WISWORK_OFFICE_RELAY_URL: 'wss://dev.example/office-relay' }),
    ).toThrow('invalid_office_relay_url')
    expect(() =>
      officeRelayEndpointFromEnv({
        WISWORK_OFFICE_RELAY_URL: 'wss://user:secret@example.com/relay',
      }),
    ).toThrow('invalid_office_relay_url')
  })

  it('requires sign-in and an exact six-digit code before connecting', async () => {
    await expect(setup(false).client.claim('123456')).rejects.toThrow('auth_required')
    await expect(setup().client.claim('12345')).rejects.toThrow('invalid_verification_code')
  })

  it('passes a freshly loaded access token to the authenticated socket factory on every claim', async () => {
    const sockets = [new FakeSocket(), new FakeSocket()]
    const connect = vi.fn((_url: string, _token: string) => sockets.shift()!)
    const getAccessToken = vi
      .fn()
      .mockResolvedValueOnce('token-one')
      .mockResolvedValueOnce('token-two')
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken,
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      onPending() {},
    })
    const first = client.claim('123456')
    await vi.waitFor(() => expect(connect).toHaveBeenCalledTimes(1))
    connect.mock.results[0]!.value.open()
    await first
    const second = client.claim('654321')
    await vi.waitFor(() => expect(connect).toHaveBeenCalledTimes(2))
    connect.mock.results[1]!.value.open()
    await second
    expect(connect).toHaveBeenNthCalledWith(
      1,
      'wss://office.8-216-134-194.sslip.io/office-relay',
      'token-one',
    )
    expect(connect).toHaveBeenNthCalledWith(
      2,
      'wss://office.8-216-134-194.sslip.io/office-relay',
      'token-two',
    )
  })

  it('claims, shows the relay-asserted host and code, and requires explicit approval', async () => {
    const { client, socket, pending } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 1,
      type: 'pc.claim',
      verification_code: '123456',
    })
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    expect(pending).toHaveBeenCalledWith({
      pairingId: 'pairing_12345678',
      hostLabel: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verificationCode: '123456',
    })
    await client.approve('pairing_12345678')
    expect(JSON.parse(socket.sent.at(-1)!)).toEqual({
      version: 1,
      type: 'pc.approve',
      pairing_id: 'pairing_12345678',
    })
  })

  it('keeps ordinary v2 code pairing without enhanced fields when persistence is disabled', async () => {
    const socket = new FakeSocket()
    const onBinding = vi.fn()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'access-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: false,
      onBinding,
      onPending() {},
    })

    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.negotiate',
      verification_code: '123456',
      capabilities: ['agent.v1'],
    })
    expect(JSON.parse(socket.sent[0]!)).not.toHaveProperty('features')

    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
    })
    expect(JSON.parse(socket.sent[1]!)).toEqual({
      version: 2,
      type: 'pc.claim',
      verification_code: '123456',
      capabilities: ['agent.v1'],
    })
    expect(onBinding).not.toHaveBeenCalled()
  })

  it('fails closed if the relay echoes a different human verification code', async () => {
    const { client, socket, pending } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '654321',
      expires_in: 120,
    })
    expect(pending).not.toHaveBeenCalled()
    expect(client.status()).toBe('disconnected:protocol_violation')
  })

  it('rejects approval before the exact pending pairing was locally approved', async () => {
    const { client, socket } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    expect(client.status()).toBe('disconnected:protocol_violation')
  })

  it('rejects non-object agent request bodies without invoking the proxy', async () => {
    const { client, socket, proxy } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: '{"messages":[]}',
    })
    expect(proxy).not.toHaveBeenCalled()
    expect(client.status()).toBe('disconnected:protocol_violation')
  })

  it('bounds remembered request identifiers and rejects unknown relay error codes', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        queueMicrotask(() => socket.open())
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      onPending() {},
      maxRequestIds: 0,
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: {},
    })
    expect(client.status()).toBe('disconnected:protocol_violation')

    const second = setup()
    const secondClaim = second.client.claim('123456')
    await vi.waitFor(() => expect(second.socket.listeners.has('open')).toBe(true))
    second.socket.open()
    await secondClaim
    second.socket.message({ version: 1, type: 'relay.error', code: 'attacker_supplied_status' })
    expect(second.client.status()).toBe('disconnected:protocol_violation')
  })

  it('proxies bounded requests, streams base64 chunks, and mirrors completion metadata', async () => {
    const { client, socket, proxy } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Excel',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: { model: 'x' },
    })
    await vi.waitFor(() => expect(proxy).toHaveBeenCalled())
    await vi.waitFor(() =>
      expect(socket.sent.some((raw) => JSON.parse(raw).type === 'pc.done')).toBe(true),
    )
    const frames = socket.sent.map((raw) => JSON.parse(raw))
    expect(frames).toContainEqual({
      version: 1,
      type: 'pc.chunk',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      request_id: 'request_12345678',
      sequence: 0,
      data: 'aGVsbG8=',
    })
    expect(frames).toContainEqual({
      version: 1,
      type: 'pc.start',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      request_id: 'request_12345678',
      status: 200,
      content_type: 'text/event-stream',
    })
    expect(frames).toContainEqual({
      version: 1,
      type: 'pc.done',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      request_id: 'request_12345678',
    })
  })

  it('reports Relay session expiry explicitly', async () => {
    const { client, socket } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'PowerPoint',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    expect(client.status()).toBe('paired')
    socket.message({ version: 1, type: 'relay.error', code: 'session_expired' })
    expect(client.status()).toBe('disconnected:session_expired')
  })

  it('reports a legacy pairing protocol mismatch without calling it a malformed frame', async () => {
    const { client, socket } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({ version: 1, type: 'relay.error', code: 'protocol_version_mismatch' })
    expect(client.status()).toBe('disconnected:protocol_version_mismatch')
  })

  it.each([100, 199, 204, 205, 304])(
    'returns request_failed without ending the session for non-streaming status %i',
    async (status) => {
      const socket = new FakeSocket()
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect: () => socket,
        getValidAccountStatus: async () => ({ loggedIn: true }),
        getAccessToken: async () => 'token',
        proxy: async () => ({ status, body: new Uint8Array() }),
        onPending() {},
      })
      const claiming = client.claim('123456')
      await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
      socket.open()
      await claiming
      socket.message({
        version: 1,
        type: 'pc.claimed',
        pairing_id: 'pairing_12345678',
        host: 'Word',
        origin: 'https://office.8-216-134-194.sslip.io',
        verification_code: '123456',
        expires_in: 120,
      })
      await client.approve('pairing_12345678')
      socket.message({
        version: 1,
        type: 'pc.approved',
        session_id: 'session_12345678',
        capability: 'secret-capability',
        expires_in: 1800,
      })
      socket.message({
        version: 1,
        type: 'relay.request',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        body: {},
      })
      await vi.waitFor(() =>
        expect(socket.sent.some((raw) => JSON.parse(raw).type === 'pc.error')).toBe(true),
      )
      const frames = socket.sent.map((raw) => JSON.parse(raw))
      expect(frames).not.toContainEqual(expect.objectContaining({ type: 'pc.start' }))
      expect(frames).toContainEqual(
        expect.objectContaining({ type: 'pc.error', code: 'request_failed' }),
      )
      expect(client.status()).toBe('paired')
    },
  )

  it('aborts active upstream work on relay.cancel', async () => {
    const socket = new FakeSocket()
    let aborted = false
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'access-token',
      proxy: ({ signal }) =>
        new Promise((_resolve, reject) => {
          signal.addEventListener('abort', () => {
            aborted = true
            reject(new Error('cancelled'))
          })
        }),
      onPending() {},
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: {},
    })
    await vi.waitFor(() => expect(client.status()).toBe('paired'))
    socket.message({
      version: 1,
      type: 'relay.cancel',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
    })
    await vi.waitFor(() => expect(aborted).toBe(true))
    expect(socket.sent.map((raw) => JSON.parse(raw).type)).not.toContain('pc.error')
    expect(client.status()).toBe('paired')
  })

  it('lets Relay own the 300s deadline and tolerates its late cancel after the 305s watchdog', async () => {
    vi.useFakeTimers()
    const socket = new FakeSocket()
    let aborted = false
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        queueMicrotask(() => socket.open())
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy: ({ signal }) =>
        new Promise((_resolve, reject) =>
          signal.addEventListener('abort', () => {
            aborted = true
            reject(new Error('aborted'))
          }),
        ),
      onPending() {},
    })
    await client.claim('123456')
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    socket.message({
      version: 1,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      body: {},
    })
    await vi.advanceTimersByTimeAsync(300_000)
    expect(aborted).toBe(false)
    await vi.advanceTimersByTimeAsync(5_000)
    expect(aborted).toBe(true)
    socket.message({
      version: 1,
      type: 'relay.cancel',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
    })
    expect(client.status()).toBe('paired')
    vi.useRealTimers()
  })

  it('expires a pending pairing and removes its approval prompt', async () => {
    vi.useFakeTimers()
    const socket = new FakeSocket()
    const expired = vi.fn()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      onPending() {},
      onPendingExpired: expired,
    })
    const claiming = client.claim('123456')
    await vi.advanceTimersByTimeAsync(0)
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 1,
    })
    await vi.advanceTimersByTimeAsync(1_000)
    expect(expired).toHaveBeenCalledWith('pairing_12345678')
    expect(client.listPending()).toEqual([])
    expect(client.status()).toBe('disconnected:pairing_expired')
    vi.useRealTimers()
  })

  it('closes the capability on logout or authentication loss', async () => {
    const { client, socket } = setup()
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'PowerPoint',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    client.revoke('auth_required')
    expect(socket.readyState).toBe(3)
    expect(client.status()).toBe('disconnected:auth_required')
  })

  it('does not apply the renewable 30 minute Relay idle TTL as a local hard expiry', async () => {
    vi.useFakeTimers()
    const { client, socket } = setup()
    const claiming = client.claim('123456')
    await vi.advanceTimersByTimeAsync(0)
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 1,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
    })
    await vi.advanceTimersByTimeAsync(31 * 60 * 1_000)
    expect(client.status()).toBe('paired')
    client.revoke('test_complete')
    vi.useRealTimers()
  })

  it('ignores old socket messages and proxy completions after re-pairing', async () => {
    const sockets = [new FakeSocket(), new FakeSocket()]
    let currentSocket = 0
    const completions: Array<(value: { status: number; body: Uint8Array }) => void> = []
    const signals: AbortSignal[] = []
    const proxy = vi.fn(
      ({ signal }: { signal: AbortSignal }) =>
        new Promise<{ status: number; body: Uint8Array }>((resolve) => {
          completions.push(resolve)
          signals.push(signal)
        }),
    )
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => sockets[currentSocket++]!,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy,
      onPending() {},
    })
    for (const [index, socket] of sockets.entries()) {
      const claiming = client.claim('123456')
      await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
      socket.open()
      await claiming
      socket.message({
        version: 1,
        type: 'pc.claimed',
        pairing_id: 'pairing_12345678',
        host: 'PowerPoint',
        origin: 'https://office.8-216-134-194.sslip.io',
        verification_code: '123456',
        expires_in: 120,
      })
      await client.approve('pairing_12345678')
      socket.message({
        version: 1,
        type: 'pc.approved',
        session_id: `session_1234567${index}`,
        capability: 'secret-capability',
        expires_in: 1800,
      })
      {
        socket.message({
          version: 1,
          type: 'relay.request',
          session_id: `session_1234567${index}`,
          request_id: 'request_12345678',
          body: {},
        })
        await vi.waitFor(() => expect(proxy).toHaveBeenCalledTimes(index + 1))
      }
    }
    const sent = [...sockets[1]!.sent]
    completions[0]!({ status: 200, body: new TextEncoder().encode('{}') })
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(sockets[1]!.sent).toEqual(sent)
    sockets[0]!.emit('message', { data: 'malformed late frame' })
    expect(client.status()).toBe('paired')
    sockets[1]!.message({
      version: 1,
      type: 'relay.cancel',
      session_id: 'session_12345671',
      request_id: 'request_12345678',
    })
    expect(signals[1]!.aborted).toBe(true)
    completions[1]!({ status: 200, body: new Uint8Array() })
    client.revoke('test_complete')
  })

  it.each([
    'presentation.v1',
    'presentation-attachments.v1',
    'presentation-assets.v1',
    'presentation-remote-images.v1',
    'presentation-webpages.v1',
    'presentation-asset-rights.v1',
    'presentation-animation-frame.v1',
    'presentation-pdf.v1',
    'presentation-production-pdf.v1',
  ])(
    'negotiates presentation only when provided and streams recoverable generation requests',
    async (capabilityName) => {
      const socket = new FakeSocket()
      const payload = new TextEncoder().encode(JSON.stringify({ result: 'x'.repeat(70_000) }))
      const presentationProxy = vi
        .fn<(body: unknown, signal: AbortSignal) => Promise<Uint8Array>>()
        .mockRejectedValueOnce(new Error('provider_unavailable'))
        .mockResolvedValue(payload)
      const proxy = vi.fn()
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect: () => socket,
        getValidAccountStatus: async () => ({ loggedIn: true }),
        getAccessToken: async () => 'token',
        proxy,
        presentationProxy,
        onPending() {},
      })
      const claiming = client.claim('123456')
      await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
      socket.open()
      await claiming
      const capabilities = [
        'agent.v1',
        'presentation.v1',
        'presentation-attachments.v1',
        'presentation-assets.v1',
        'presentation-remote-images.v1',
        'presentation-webpages.v1',
        'presentation-asset-rights.v1',
        'presentation-animation-frame.v1',
        'presentation-pdf.v1',
        'presentation-production-pdf.v1',
        'presentation-master-backups.v1',
        'presentation-package-backups.v1',
      ]
      expect(JSON.parse(socket.sent[0]!)).toEqual({
        version: 2,
        type: 'pc.negotiate',
        verification_code: '123456',
        capabilities,
      })
      socket.message({ version: 2, type: 'pc.negotiated', pairing_version: 2, capabilities })
      if (socket.closedWith) return { socket, client, presentationProxy, offered }
      socket.message({
        version: 2,
        type: 'pc.claimed',
        pairing_id: 'pairing_12345678',
        host: 'PowerPoint',
        origin: 'https://office.8-216-134-194.sslip.io',
        verification_code: '123456',
        expires_in: 120,
        capabilities,
      })
      await client.approve('pairing_12345678')
      socket.message({
        version: 2,
        type: 'pc.approved',
        session_id: 'session_12345678',
        capability: 'secret-capability',
        expires_in: 1800,
        capabilities,
      })
      for (const request_id of ['request_failure', 'request_success']) {
        const body =
          capabilityName === 'presentation-pdf.v1'
            ? { operation: 'export_pdf', projectId: 'deck', requestId: 'first' }
            : capabilityName === 'presentation-production-pdf.v1'
              ? {
                  operation: 'export_pdf',
                  source: 'production',
                  projectId: 'deck',
                  requestId: 'first',
                }
              : { instruction: 'Create a deck' }
        socket.message({
          version: 2,
          type: 'relay.request',
          session_id: 'session_12345678',
          request_id,
          capability_name: capabilityName,
          body,
        })
        await vi.waitFor(() =>
          expect(socket.sent.map((raw) => JSON.parse(raw))).toContainEqual(
            expect.objectContaining({
              type: request_id === 'request_failure' ? 'pc.error' : 'pc.done',
              request_id,
            }),
          ),
        )
        expect(client.status()).toBe('paired')
      }
      expect(presentationProxy).toHaveBeenCalledWith(
        capabilityName === 'presentation-pdf.v1'
          ? { operation: 'export_pdf', projectId: 'deck', requestId: 'first' }
          : capabilityName === 'presentation-production-pdf.v1'
            ? {
                operation: 'export_pdf',
                source: 'production',
                projectId: 'deck',
                requestId: 'first',
              }
            : { instruction: 'Create a deck' },
        expect.any(AbortSignal),
      )
      expect(proxy).not.toHaveBeenCalled()
      const chunks = socket.sent
        .map((raw) => JSON.parse(raw))
        .filter((item) => item.type === 'pc.chunk')
      expect(chunks.map((item) => item.sequence)).toEqual([0, 1])
      expect(Buffer.concat(chunks.map((item) => Buffer.from(item.data, 'base64')))).toEqual(
        Buffer.from(payload),
      )
      if (capabilityName === 'presentation-attachments.v1') {
        socket.message({
          version: 2,
          type: 'relay.request',
          session_id: 'session_12345678',
          request_id: 'wrong_remote_capability',
          capability_name: capabilityName,
          body: { operation: 'attachment_import_url', url: 'https://example.com/image.png' },
        })
        await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
      }
      if (capabilityName === 'presentation-webpages.v1') {
        socket.message({
          version: 2,
          type: 'relay.request',
          session_id: 'session_12345678',
          request_id: 'wrong_webpage_capability',
          capability_name: 'presentation-attachments.v1',
          body: { operation: 'attachment_import_webpage', url: 'https://example.com/page' },
        })
        await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
      }
      if (capabilityName === 'presentation.v1') {
        socket.message({
          version: 2,
          type: 'relay.request',
          session_id: 'session_12345678',
          request_id: 'wrong_pdf_capability',
          capability_name: capabilityName,
          body: { operation: 'export_pdf', projectId: 'deck', requestId: 'first' },
        })
        await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
      }
      if (capabilityName === 'presentation-assets.v1') {
        socket.message({
          version: 2,
          type: 'relay.request',
          session_id: 'session_12345678',
          request_id: 'wrong_rights_capability',
          capability_name: capabilityName,
          body: {
            operation: 'attachment_attest_license',
            attachmentId: 'a'.repeat(64),
            license: 'owned',
            evidenceAttachmentId: 'b'.repeat(64),
          },
        })
        await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
      }
      if (capabilityName === 'presentation-animation-frame.v1') {
        socket.message({
          version: 2,
          type: 'relay.request',
          session_id: 'session_12345678',
          request_id: 'wrong_animation_capability',
          capability_name: 'presentation-attachments.v1',
          body: { operation: 'attachment_extract_first_frame', attachmentId: 'a'.repeat(64) },
        })
        await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
      }
      if (capabilityName === 'presentation-pdf.v1') {
        socket.message({
          version: 2,
          type: 'relay.request',
          session_id: 'session_12345678',
          request_id: 'wrong_pdf_capability',
          capability_name: 'presentation-pdf.v1',
          body: {
            operation: 'export_pdf',
            source: 'production',
            projectId: 'deck',
            requestId: 'first',
          },
        })
        await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
      }
      if (capabilityName === 'presentation-production-pdf.v1') {
        socket.message({
          version: 2,
          type: 'relay.request',
          session_id: 'session_12345678',
          request_id: 'wrong_production_pdf_capability',
          capability_name: 'presentation-production-pdf.v1',
          body: { operation: 'export_pdf', projectId: 'deck', requestId: 'first' },
        })
        await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
      }
      client.revoke('test_complete')
    },
  )

  it('uses v2 only with a fixed retrieval proxy and dispatches negotiated web requests', async () => {
    const socket = new FakeSocket()
    const retrievalProxy = vi.fn(async () => new TextEncoder().encode('{"results":[]}'))
    const agentProxy = vi.fn()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'access-token',
      proxy: agentProxy,
      retrievalProxy,
      onPending() {},
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.negotiate',
      verification_code: '123456',
      capabilities: [
        'agent.v1',
        'web-search.v1',
        'web-fetch.v1',
        'image-search.v1',
        'image-fetch.v1',
      ],
    })
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1', 'web-search.v1'],
    })
    expect(JSON.parse(socket.sent[1]!)).toEqual({
      version: 2,
      type: 'pc.claim',
      verification_code: '123456',
      capabilities: [
        'agent.v1',
        'web-search.v1',
        'web-fetch.v1',
        'image-search.v1',
        'image-fetch.v1',
      ],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1', 'web-search.v1'],
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1', 'web-search.v1'],
    })
    socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      capability_name: 'web-search.v1',
      body: { query: 'office', max_results: 3 },
    })
    await vi.waitFor(() => expect(retrievalProxy).toHaveBeenCalled())
    expect(retrievalProxy).toHaveBeenCalledWith(
      'web-search.v1',
      { query: 'office', max_results: 3 },
      expect.any(AbortSignal),
    )
    expect(agentProxy).not.toHaveBeenCalled()
    await vi.waitFor(() =>
      expect(socket.sent.map((raw) => JSON.parse(raw))).toContainEqual(
        expect.objectContaining({ version: 2, type: 'pc.done' }),
      ),
    )
  })

  it('explicitly negotiates the Office protocol and follows a Relay-asserted v1 pairing', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'access-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      onPending() {},
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.negotiate',
      verification_code: '123456',
      capabilities: ['agent.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 1,
      capabilities: ['agent.v1'],
    })
    expect(JSON.parse(socket.sent[1]!)).toEqual({
      version: 1,
      type: 'pc.claim',
      verification_code: '123456',
    })
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    expect(client.status()).toBe('awaiting_approval')
  })

  it('does not silently downgrade when Relay skips explicit negotiation', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'access-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      onPending() {},
    })
    const claiming = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claiming
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    expect(client.status()).toBe('disconnected:protocol_violation')
  })

  it('conditionally negotiates pairing-resume.v1 and captures binding metadata after approval', async () => {
    const socket = new FakeSocket()
    const onBinding = vi.fn(async () => undefined)
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'fresh-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: true,
      onBinding,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.negotiate',
      verification_code: '123456',
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    expect(JSON.parse(socket.sent[1]!)).toEqual({
      version: 2,
      type: 'pc.claim',
      verification_code: '123456',
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    expect(JSON.parse(socket.sent.at(-1)!)).toEqual({
      version: 2,
      type: 'pc.approve',
      pairing_id: 'pairing_12345678',
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
      binding_id: 'binding_word_12345678',
    })
    await vi.waitFor(() =>
      expect(onBinding).toHaveBeenCalledWith({
        bindingId: 'binding_word_12345678',
        accountId: 'local-account',
        host: 'Word',
        origin: 'https://office.8-216-134-194.sslip.io',
        capabilities: ['agent.v1'],
        createdAt: expect.any(Number),
      }),
    )
    expect(client.status()).toBe('paired')
  })

  it('accepts only the exact enhanced short-session fallback when durable enrollment aborts', async () => {
    const sockets: FakeSocket[] = []
    const onBinding = vi.fn(async () => undefined)
    let persistenceAvailable = true
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'fresh-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: () => persistenceAvailable,
      onBinding,
      onPending() {},
    })

    const claim = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(1))
    sockets[0]!.open()
    await claim
    sockets[0]!.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    sockets[0]!.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    sockets[0]!.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: [],
    })
    expect(client.status()).toBe('paired')
    expect(onBinding).not.toHaveBeenCalled()

    client.revoke('next_claim')
    persistenceAvailable = false
    const ordinaryClaim = client.claim('654321')
    await vi.waitFor(() => expect(sockets).toHaveLength(2))
    sockets[1]!.open()
    await ordinaryClaim
    expect(JSON.parse(sockets[1]!.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.negotiate',
      verification_code: '654321',
      capabilities: ['agent.v1'],
    })
  })

  it('accepts an explicit features-empty approval when Relay disables durable pairing', async () => {
    const socket = new FakeSocket()
    const onBinding = vi.fn(async () => undefined)
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'fresh-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: true,
      onBinding,
      onPending() {},
    })

    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: [],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: [],
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: [],
    })

    expect(client.status()).toBe('paired')
    expect(onBinding).not.toHaveBeenCalled()
  })

  it.each([
    {
      label: 'omitted fallback feature marker',
      approved: {},
    },
    {
      label: 'fallback with a binding id',
      approved: { features: [], binding_id: 'binding_word_12345678' },
    },
    {
      label: 'unnegotiated fallback feature',
      approved: { features: ['future-feature.v1'] },
    },
  ])('rejects $label after enhanced enrollment', async ({ approved }) => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'fresh-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      ...approved,
    })
    expect(client.status()).toBe('disconnected:protocol_violation')
  })

  it('revokes the durable binding and reports not remembered when encrypted persistence fails', async () => {
    const sockets: FakeSocket[] = []
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: vi
        .fn()
        .mockResolvedValueOnce('pair-token')
        .mockResolvedValueOnce('revoke-token'),
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onBinding: async () => {
        throw new Error('disk_failure')
      },
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(1))
    sockets[0]!.open()
    await claim
    sockets[0]!.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    sockets[0]!.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    sockets[0]!.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
      binding_id: 'binding_word_12345678',
    })

    await vi.waitFor(() => expect(sockets).toHaveLength(2))
    expect(client.status()).not.toBe('paired')
    sockets[1]!.open()
    await vi.waitFor(() => expect(sockets[1]!.sent).toHaveLength(1))
    expect(JSON.parse(sockets[1]!.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.revoke_binding',
      binding_id: 'binding_word_12345678',
    })
    sockets[1]!.message({
      version: 2,
      type: 'pc.binding_revoked',
      binding_id: 'binding_word_12345678',
    })
    await vi.waitFor(() => expect(client.status()).toBe('disconnected:binding_not_remembered'))
  })

  it('serializes approval persistence before requests and ignores an exact duplicate approval', async () => {
    const socket = new FakeSocket()
    let releaseBinding!: () => void
    const bindingSaved = new Promise<void>((resolve) => {
      releaseBinding = resolve
    })
    const onBinding = vi.fn(() => bindingSaved)
    const proxy = vi.fn(async () => ({ status: 200, body: new Uint8Array() }))
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'token',
      proxy,
      persistentPairing: true,
      onBinding,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await client.approve('pairing_12345678')
    const approved = {
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'secret-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
      binding_id: 'binding_word_12345678',
    }
    socket.message(approved)
    await vi.waitFor(() => expect(onBinding).toHaveBeenCalledOnce())
    socket.message(approved)
    socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      capability_name: 'agent.v1',
      body: {},
    })
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(client.status()).not.toBe('disconnected:protocol_violation')
    expect(proxy).not.toHaveBeenCalled()

    releaseBinding()
    await vi.waitFor(() => expect(client.status()).toBe('paired'))
    await vi.waitFor(() => expect(proxy).toHaveBeenCalledOnce())
    expect(onBinding).toHaveBeenCalledOnce()
  })

  it('rejects a claimed feature upgrade beyond the negotiated intersection', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 2,
      capabilities: ['agent.v1'],
      features: [],
    })
    socket.message({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities: ['agent.v1'],
      features: ['pairing-resume.v1'],
    })
    await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
  })

  it('uses exact legacy v1 frames after enhanced negotiation selects pairing_version 1', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await claim
    socket.message({
      version: 2,
      type: 'pc.negotiated',
      pairing_version: 1,
      capabilities: ['agent.v1'],
      features: [],
    })
    expect(JSON.parse(socket.sent.at(-1)!)).toEqual({
      version: 1,
      type: 'pc.claim',
      verification_code: '123456',
    })
    socket.message({
      version: 1,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
    })
    expect(client.status()).toBe('awaiting_approval')
  })

  it.each(['invalid_frame', 'unknown_type'] as const)(
    'falls back once only for explicit old-schema Relay error %s',
    async (code) => {
      const sockets: FakeSocket[] = []
      const getAccessToken = vi
        .fn()
        .mockResolvedValueOnce('token-one')
        .mockResolvedValueOnce('token-two')
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect: () => {
          const socket = new FakeSocket()
          sockets.push(socket)
          return socket
        },
        getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
        getAccessToken,
        proxy: async () => ({ status: 200, body: new Uint8Array() }),
        negotiateCapabilities: true,
        persistentPairing: true,
        onPending() {},
      })
      const claim = client.claim('123456')
      await vi.waitFor(() => expect(sockets).toHaveLength(1))
      sockets[0]!.open()
      await claim
      expect(JSON.parse(sockets[0]!.sent[0]!)).toHaveProperty('features', ['pairing-resume.v1'])
      sockets[0]!.message({ version: 2, type: 'relay.error', code })
      await vi.waitFor(() => expect(sockets).toHaveLength(2))
      sockets[1]!.open()
      await vi.waitFor(() => expect(sockets[1]!.sent).toHaveLength(1))
      expect(JSON.parse(sockets[1]!.sent[0]!)).toEqual({
        version: 2,
        type: 'pc.negotiate',
        verification_code: '123456',
        capabilities: ['agent.v1'],
      })
      expect(getAccessToken).toHaveBeenNthCalledWith(1)
      expect(getAccessToken).toHaveBeenNthCalledWith(2)
    },
  )

  it.each(['invalid_code', 'auth_required', 'resume_rate_limited'] as const)(
    'does not legacy-fallback for semantic Relay error %s',
    async (code) => {
      const sockets: FakeSocket[] = []
      const client = createOfficeRelayClient({
        endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
        connect: () => {
          const socket = new FakeSocket()
          sockets.push(socket)
          return socket
        },
        getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
        getAccessToken: async () => 'token',
        proxy: async () => ({ status: 200, body: new Uint8Array() }),
        persistentPairing: true,
        onPending() {},
      })
      const claim = client.claim('123456')
      await vi.waitFor(() => expect(sockets).toHaveLength(1))
      sockets[0]!.open()
      await claim
      sockets[0]!.message({ version: 2, type: 'relay.error', code })
      await new Promise((resolve) => setTimeout(resolve, 0))
      expect(sockets).toHaveLength(1)
    },
  )

  it('keeps an invalid verification code actionable in the public status', async () => {
    const sockets: FakeSocket[] = []
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(1))
    sockets[0]!.open()
    await claim
    sockets[0]!.message({ version: 2, type: 'relay.error', code: 'invalid_code' })
    expect(client.status()).toBe('disconnected:invalid_code')
  })

  it('does not downgrade on a network close and keeps the next claim enhanced', async () => {
    const sockets: FakeSocket[] = []
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => {
        const socket = new FakeSocket()
        sockets.push(socket)
        return socket
      },
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const first = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(1))
    sockets[0]!.open()
    await first
    sockets[0]!.close()
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(sockets).toHaveLength(1)

    const second = client.claim('123456')
    await vi.waitFor(() => expect(sockets).toHaveLength(2))
    sockets[1]!.open()
    await second
    expect(JSON.parse(sockets[1]!.sent[0]!)).toHaveProperty('features', ['pairing-resume.v1'])
  })

  it('uses a fresh access token for pc.resume and accepts waiting_for_office then standard v2 approval', async () => {
    const socket = new FakeSocket()
    const getAccessToken = vi.fn(async () => 'resume-token')
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken,
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      negotiateCapabilities: true,
      persistentPairing: true,
      onPending() {},
    })
    const binding: OfficeRelayBinding = {
      bindingId: 'binding_word_12345678',
      accountId: 'local-account',
      host: 'Word',
      origin: 'https://office.8-216-134-194.sslip.io',
      capabilities: ['agent.v1'],
      createdAt: 1,
    }
    const resume = client.resume(binding)
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await resume
    expect(getAccessToken).toHaveBeenCalledOnce()
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.resume',
      binding_id: 'binding_word_12345678',
      capabilities: ['agent.v1'],
    })
    socket.message({ version: 2, type: 'pc.waiting_for_office' })
    expect(client.status()).toBe('waiting_for_office')
    socket.message({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_resume_12345678',
      capability: 'fresh-capability',
      expires_in: 1800,
      capabilities: ['agent.v1'],
    })
    expect(client.status()).toBe('paired')
  })

  it('sends an exact authenticated revocation frame and requires its exact acknowledgement', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'revoke-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const revocation = client.revokeBinding('binding_word_12345678', 'local-account')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await vi.waitFor(() => expect(socket.sent).toHaveLength(1))
    expect(JSON.parse(socket.sent[0]!)).toEqual({
      version: 2,
      type: 'pc.revoke_binding',
      binding_id: 'binding_word_12345678',
    })
    socket.message({
      version: 2,
      type: 'pc.binding_revoked',
      binding_id: 'binding_word_12345678',
    })
    await expect(revocation).resolves.toBeUndefined()
  })

  it('treats binding_unavailable as idempotent success only for an authenticated revocation', async () => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'revoke-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const revocation = client.revokeBinding('binding_word_12345678', 'local-account')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await vi.waitFor(() => expect(socket.sent).toHaveLength(1))
    socket.message({ version: 2, type: 'relay.error', code: 'binding_unavailable' })
    await expect(revocation).resolves.toBeUndefined()
    expect(client.status()).toBe('disconnected:binding_revoked')

    const claimSocket = new FakeSocket()
    const claiming = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => claimSocket,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'local-account' }),
      getAccessToken: async () => 'claim-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })
    const claim = claiming.claim('123456')
    await vi.waitFor(() => expect(claimSocket.listeners.has('open')).toBe(true))
    claimSocket.open()
    await claim
    claimSocket.message({ version: 2, type: 'relay.error', code: 'binding_unavailable' })
    expect(claiming.status()).toBe('disconnected:relay_error')
  })

  it('refuses to revoke an old-account binding with replacement-account credentials', async () => {
    const connect = vi.fn(() => new FakeSocket())
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: 'replacement-account' }),
      getAccessToken: async () => 'replacement-token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })

    await expect(client.revokeBinding('binding_word_12345678', 'old-account')).rejects.toThrow(
      'auth_required',
    )
    expect(connect).not.toHaveBeenCalled()
  })

  it('rechecks the expected revocation account after loading a token', async () => {
    let accountId = 'old-account'
    let releaseToken!: (token: string) => void
    const token = new Promise<string>((resolve) => {
      releaseToken = resolve
    })
    const connect = vi.fn(() => new FakeSocket())
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect,
      getValidAccountStatus: async () => ({ loggedIn: true, userId: accountId }),
      getAccessToken: () => token,
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      persistentPairing: true,
      onPending() {},
    })

    const revoking = client.revokeBinding('binding_word_12345678', 'old-account')
    await vi.waitFor(() => expect(client.status()).toBe('disconnected:new_revocation'))
    accountId = 'replacement-account'
    releaseToken('replacement-token')
    await expect(revoking).rejects.toThrow('auth_required')
    expect(connect).not.toHaveBeenCalled()
  })
})

async function teamPcClient(
  supportsTeamPresentation = true,
  negotiatedCapabilities?: string[],
  host = 'PowerPoint',
) {
  const socket = new FakeSocket(),
    presentationProxy = vi.fn(
      async (
        _body: unknown,
        _signal: AbortSignal,
        _context?: import('../src/main/office-relay-client').PresentationTeamContext,
      ) => new TextEncoder().encode('{}'),
    )
  const client = createOfficeRelayClient({
    endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
    connect: () => socket,
    getValidAccountStatus: async () => ({ loggedIn: true }),
    getAccessToken: async () => 'pc-token',
    proxy: vi.fn(),
    presentationProxy,
    ...(supportsTeamPresentation ? { supportsTeamPresentation: true as const } : {}),
    onPending: vi.fn(),
  })
  const claiming = client.claim('123456')
  await vi.waitFor(() => expect(socket.listeners.get('open')?.length).toBeGreaterThan(0))
  socket.open()
  await claiming
  const offered = JSON.parse(socket.sent[0]!).capabilities
  const capabilities =
    negotiatedCapabilities ??
    (supportsTeamPresentation ? ['presentation-team.v1'] : ['agent.v1', 'presentation.v1'])
  socket.message({ version: 2, type: 'pc.negotiated', pairing_version: 2, capabilities })
  if (socket.closedWith) return { socket, client, presentationProxy, offered }
  socket.message({
    version: 2,
    type: 'pc.claimed',
    pairing_id: 'pairing_12345678',
    host,
    origin: 'https://office.8-216-134-194.sslip.io',
    verification_code: '123456',
    expires_in: 120,
    capabilities,
  })
  await client.approve('pairing_12345678')
  socket.message({
    version: 2,
    type: 'pc.approved',
    session_id: 'session_12345678',
    capability: 'secret-capability',
    expires_in: 1800,
    capabilities,
  })
  return { socket, client, presentationProxy, offered }
}
it('advertises team transport only explicitly and passes verified context separately from the body', async () => {
  const old = await teamPcClient(false)
  expect(old.offered).not.toContain('presentation-team.v1')
  old.client.revoke()
  const f = await teamPcClient(),
    context = { version: 1, actorSubject: 'a'.repeat(64), pcSubject: 'b'.repeat(64) },
    body = { operation: 'team_identity' }
  expect(f.offered).toContain('presentation-team.v1')
  f.socket.message({
    version: 2,
    type: 'relay.request',
    session_id: 'session_12345678',
    request_id: 'request_12345678',
    capability_name: 'presentation-team.v1',
    team_context: context,
    body,
  })
  await vi.waitFor(() =>
    expect(f.presentationProxy).toHaveBeenCalledWith(body, expect.any(AbortSignal), context),
  )
  expect(f.socket.sent.every((value) => !value.includes('pc-token'))).toBe(true)
  f.client.revoke()
})
it.each([
  'missing',
  'bad-subject',
  'extra',
  'unknown-operation',
  'ordinary-operation',
  'feedback-read',
  'feedback-record',
  'feedback-compare',
  'bearer',
])('rejects %s team request context before proxying', async (scenario) => {
  const f = await teamPcClient(),
    context: Record<string, unknown> = {
      version: 1,
      actorSubject: 'a'.repeat(64),
      pcSubject: 'b'.repeat(64),
    },
    body: Record<string, unknown> = { operation: 'team_identity' }
  if (scenario === 'bad-subject') context.actorSubject = 'raw-actor'
  if (scenario === 'extra') context.token = 'private'
  if (scenario === 'unknown-operation') body.operation = 'team_unknown'
  if (scenario === 'ordinary-operation') body.operation = 'get_plan'
  if (scenario === 'feedback-read') body.operation = 'production_feedback_read'
  if (scenario === 'feedback-record') body.operation = 'production_feedback_record'
  if (scenario === 'feedback-compare') body.operation = 'production_feedback_compare'
  const frame: Record<string, unknown> = {
    version: 2,
    type: 'relay.request',
    session_id: 'session_12345678',
    request_id: 'request_12345678',
    capability_name: 'presentation-team.v1',
    team_context: context,
    body,
  }
  if (scenario === 'missing') delete frame.team_context
  if (scenario === 'bearer') frame.access_token = 'private-token'
  f.socket.message(frame)
  expect(f.presentationProxy).not.toHaveBeenCalled()
  expect(f.socket.closedWith).toBeDefined()
  f.client.revoke()
})
it('rejects team operations and forged context through an ordinary presentation capability', async () => {
  for (const withContext of [false, true]) {
    const f = await teamPcClient(false)
    f.socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      capability_name: 'presentation.v1',
      body: { operation: 'team_identity' },
      ...(withContext
        ? { team_context: { version: 1, actorSubject: 'a'.repeat(64), pcSubject: 'b'.repeat(64) } }
        : {}),
    })
    expect(f.presentationProxy).not.toHaveBeenCalled()
    expect(f.socket.closedWith).toBeDefined()
    f.client.revoke()
  }
})
it('refuses a forged mixed team/private negotiated session', async () => {
  const f = await teamPcClient(true, ['presentation-team.v1', 'presentation.v1'])
  expect(f.socket.closedWith).toBeDefined()
  expect(f.presentationProxy).not.toHaveBeenCalled()
  f.client.revoke()
})

it.each([
  'master_backup_begin',
  'master_backup_chunk',
  'master_backup_finish',
  'master_backup_status',
  'master_backup_read',
  'master_backup_list',
])('routes only negotiated PowerPoint master backup operation %s', async (operation) => {
  const f = await teamPcClient(false, ['agent.v1', 'presentation-master-backups.v1'])
  expect(f.offered).toContain('presentation-master-backups.v1')
  expect(f.client.status()).toBe('paired')
  const body = { operation, documentId: 'document', changeId: 'change', key: 'snapshot' }
  f.socket.message({
    version: 2,
    type: 'relay.request',
    session_id: 'session_12345678',
    request_id: 'master_1',
    capability_name: 'presentation-master-backups.v1',
    body,
  })
  await vi.waitFor(() =>
    expect(f.presentationProxy).toHaveBeenCalledWith(body, expect.any(AbortSignal)),
  )
  await vi.waitFor(() =>
    expect(f.socket.sent.map((x) => JSON.parse(x)).some((x) => x.type === 'pc.done')).toBe(true),
  )
  f.client.revoke('complete')
})
it.each([
  ['presentation.v1', 'master_backup_status', 'PowerPoint', undefined],
  ['agent.v1', 'master_backup_status', 'PowerPoint', undefined],
  ['presentation-master-backups.v1', 'production_status', 'PowerPoint', undefined],
  ['presentation-master-backups.v1', 'master_backup_delete', 'PowerPoint', undefined],
  ['presentation-master-backups.v1', 'master_backup_status', 'Word', undefined],
  ['presentation-master-backups.v1', 'master_backup_status', 'Excel', undefined],
  ['presentation-master-backups.v1', 'master_backup_status', 'PowerPoint', {}],
])(
  'rejects master capability routing mismatch %s/%s/%s before proxy',
  async (capability_name, operation, host, context) => {
    const f = await teamPcClient(
      false,
      ['agent.v1', 'presentation.v1', 'presentation-master-backups.v1'],
      host as string,
    )
    f.socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'master_bad',
      capability_name,
      body: { operation },
      ...(context ? { team_context: context } : {}),
    })
    await vi.waitFor(() => expect(f.client.status()).toBe('disconnected:protocol_violation'))
    expect(f.presentationProxy).not.toHaveBeenCalled()
  },
)
it.each([false, true])(
  'offers the exact optional governance catalog (enabled=%s) without changing the disabled bound',
  async (enabled) => {
    const socket = new FakeSocket()
    const client = createOfficeRelayClient({
      endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
      connect: () => socket,
      getValidAccountStatus: async () => ({ loggedIn: true }),
      getAccessToken: async () => 'token',
      proxy: async () => ({ status: 200, body: new Uint8Array() }),
      retrievalProxy: async () => new Uint8Array(),
      ...(enabled
        ? {
            presentationProxy: async () => new Uint8Array(),
            supportsTeamPresentation: true as const,
            presentationGovernanceProxy: async () => new Uint8Array(),
          }
        : {}),
      onPending() {},
    })
    const pending = client.claim('123456')
    await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
    socket.open()
    await pending
    const offered = JSON.parse(socket.sent[0]!).capabilities
    expect(offered).toHaveLength(enabled ? 18 : 5)
    expect(new Set(offered).size).toBe(offered.length)
    expect(offered).toContain('web-fetch.v1')
    expect(offered.includes('presentation-governance.v1')).toBe(enabled)
    expect(offered.includes('presentation-master-backups.v1')).toBe(enabled)
    expect(offered.includes('presentation-package-backups.v1')).toBe(enabled)
    client.revoke('complete')
  },
)
it.each([
  'package_backup_begin',
  'package_backup_chunk',
  'package_backup_finish',
  'package_backup_status',
  'package_backup_read',
  'package_backup_list',
])('routes only negotiated PowerPoint package backup operation %s', async (operation) => {
  const f = await teamPcClient(false, ['agent.v1', 'presentation-package-backups.v1'])
  expect(f.offered).toContain('presentation-package-backups.v1')
  expect(f.client.status()).toBe('paired')
  const body = { operation, documentId: 'document', changeId: 'change', key: 'snapshot' }
  f.socket.message({
    version: 2,
    type: 'relay.request',
    session_id: 'session_12345678',
    request_id: 'master_1',
    capability_name: 'presentation-package-backups.v1',
    body,
  })
  await vi.waitFor(() =>
    expect(f.presentationProxy).toHaveBeenCalledWith(body, expect.any(AbortSignal)),
  )
  await vi.waitFor(() =>
    expect(f.socket.sent.map((x) => JSON.parse(x)).some((x) => x.type === 'pc.done')).toBe(true),
  )
  f.client.revoke('complete')
})
it.each([
  ['presentation.v1', 'package_backup_status', 'PowerPoint', undefined],
  ['agent.v1', 'package_backup_status', 'PowerPoint', undefined],
  ['presentation-package-backups.v1', 'production_status', 'PowerPoint', undefined],
  ['presentation-package-backups.v1', 'package_backup_delete', 'PowerPoint', undefined],
  ['presentation-package-backups.v1', 'package_backup_status', 'Word', undefined],
  ['presentation-package-backups.v1', 'package_backup_status', 'Excel', undefined],
  ['presentation-package-backups.v1', 'package_backup_status', 'PowerPoint', {}],
])(
  'rejects package capability routing mismatch %s/%s/%s before proxy',
  async (capability_name, operation, host, context) => {
    const f = await teamPcClient(
      false,
      ['agent.v1', 'presentation.v1', 'presentation-package-backups.v1'],
      host as string,
    )
    f.socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'master_bad',
      capability_name,
      body: { operation },
      ...(context ? { team_context: context } : {}),
    })
    await vi.waitFor(() => expect(f.client.status()).toBe('disconnected:protocol_violation'))
    expect(f.presentationProxy).not.toHaveBeenCalled()
  },
)

it.each([
  ['presentation-master-backups.v1', 'package_backup_status'],
  ['presentation-package-backups.v1', 'master_backup_status'],
])('rejects cross-namespace %s', async (capability_name, operation) => {
  const f = await teamPcClient(false, [
    'agent.v1',
    'presentation-master-backups.v1',
    'presentation-package-backups.v1',
  ])
  f.socket.message({
    version: 2,
    type: 'relay.request',
    session_id: 'session_12345678',
    request_id: 'cross_bad',
    capability_name,
    body: { operation },
  })
  await vi.waitFor(() => expect(f.client.status()).toBe('disconnected:protocol_violation'))
  expect(f.presentationProxy).not.toHaveBeenCalled()
})

async function governancePcClient(
  settings: {
    enabled?: boolean
    capabilities?: string[]
    host?: string
    handler?: (body: unknown, signal: AbortSignal) => Promise<Uint8Array>
  } = {},
) {
  const socket = new FakeSocket(),
    presentationProxy = vi.fn(async () => new TextEncoder().encode('{}')),
    governanceProxy = vi.fn(
      settings.handler ?? (async () => new TextEncoder().encode('{"lifecycle":null}')),
    )
  const client = createOfficeRelayClient({
    endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
    connect: () => socket,
    getValidAccountStatus: async () => ({ loggedIn: true }),
    getAccessToken: async () => 'pc-token',
    proxy: vi.fn(),
    presentationProxy,
    ...(settings.enabled === false ? {} : { presentationGovernanceProxy: governanceProxy }),
    onPending: vi.fn(),
  })
  const claiming = client.claim('123456')
  await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
  socket.open()
  await claiming
  const offered = JSON.parse(socket.sent[0]!).capabilities,
    capabilities = settings.capabilities ?? [
      'agent.v1',
      'presentation.v1',
      'presentation-governance.v1',
    ]
  socket.message({ version: 2, type: 'pc.negotiated', pairing_version: 2, capabilities })
  if (socket.closedWith) return { socket, client, presentationProxy, governanceProxy, offered }
  socket.message({
    version: 2,
    type: 'pc.claimed',
    pairing_id: 'pairing_12345678',
    host: settings.host ?? 'PowerPoint',
    origin: 'https://office.8-216-134-194.sslip.io',
    verification_code: '123456',
    expires_in: 120,
    capabilities,
  })
  if (socket.closedWith) return { socket, client, presentationProxy, governanceProxy, offered }
  await client.approve('pairing_12345678')
  socket.message({
    version: 2,
    type: 'pc.approved',
    session_id: 'session_12345678',
    capability: 'secret-capability',
    expires_in: 1800,
    capabilities,
  })
  return { socket, client, presentationProxy, governanceProxy, offered }
}
function governanceFrame(
  operation: string,
  capability_name = 'presentation-governance.v1',
  extra: Record<string, unknown> = {},
) {
  return {
    version: 2,
    type: 'relay.request',
    session_id: 'session_12345678',
    request_id: 'governance_12345678',
    capability_name,
    body: { operation, documentId: 'doc', projectId: 'project' },
    ...extra,
  }
}
it('offers governance only with a real dedicated handler and never routes it through ordinary presentation', async () => {
  const old = await governancePcClient({
    enabled: false,
    capabilities: ['agent.v1', 'presentation.v1'],
  })
  expect(old.offered).not.toContain('presentation-governance.v1')
  old.client.revoke('done')
  const f = await governancePcClient()
  expect(f.offered).toContain('presentation-governance.v1')
  f.socket.message(governanceFrame('project_lifecycle_read'))
  await vi.waitFor(() =>
    expect(f.governanceProxy).toHaveBeenCalledWith(
      { operation: 'project_lifecycle_read', documentId: 'doc', projectId: 'project' },
      expect.any(AbortSignal),
    ),
  )
  expect(f.presentationProxy).not.toHaveBeenCalled()
  await vi.waitFor(() =>
    expect(f.socket.sent.map((x) => JSON.parse(x)).some((x) => x.type === 'pc.done')).toBe(true),
  )
  f.client.revoke('done')
})
it.each([
  'project_deletion_preview',
  'project_deletion_confirm',
  'project_deletion_resume',
  'project_lifecycle_initialize',
  'project_lifecycle_read',
  'project_lifecycle_set_policy',
  'project_lifecycle_export_audit',
])('routes whitelisted governance operation %s', async (operation) => {
  const f = await governancePcClient()
  f.socket.message(governanceFrame(operation))
  await vi.waitFor(() => expect(f.governanceProxy).toHaveBeenCalledTimes(1))
  f.client.revoke('done')
})
it.each([
  ['presentation.v1', 'project_deletion_confirm', {}],
  ['agent.v1', 'project_deletion_resume', {}],
  ['presentation-governance.v1', 'unknown', {}],
  ['presentation-governance.v1', 'compile', {}],
  [
    'presentation-governance.v1',
    'project_lifecycle_read',
    { team_context: { version: 1, actorSubject: 'a'.repeat(64), pcSubject: 'b'.repeat(64) } },
  ],
  [
    'presentation-governance.v1',
    'project_lifecycle_read',
    {
      body: {
        operation: 'project_lifecycle_read',
        documentId: 'doc',
        projectId: 'project',
        access_token: 'private',
      },
    },
  ],
  [
    'presentation-governance.v1',
    'project_lifecycle_read',
    {
      body: {
        operation: 'project_lifecycle_read',
        documentId: 'doc',
        projectId: 'project',
        team_context: {},
      },
    },
  ],
])('rejects governance family/context misuse %s %s', async (capability, operation, extra) => {
  const f = await governancePcClient()
  f.socket.message(governanceFrame(operation, capability, extra))
  await vi.waitFor(() => expect(f.client.status()).toBe('disconnected:protocol_violation'))
  expect(f.governanceProxy).not.toHaveBeenCalled()
  expect(f.presentationProxy).not.toHaveBeenCalled()
})
it.each(['Word', 'Excel'])('refuses governance pairing for primary host %s', async (host) => {
  const f = await governancePcClient({ host })
  expect(f.client.status()).toBe('disconnected:protocol_violation')
  expect(f.governanceProxy).not.toHaveBeenCalled()
})
it('refuses unnegotiated governance capability despite installed handler', async () => {
  const f = await governancePcClient({ capabilities: ['agent.v1', 'presentation.v1'] })
  f.socket.message(governanceFrame('project_deletion_preview'))
  await vi.waitFor(() => expect(f.client.status()).toBe('disconnected:protocol_violation'))
  expect(f.governanceProxy).not.toHaveBeenCalled()
})
it('refuses negotiated governance when no handler is installed', async () => {
  const f = await governancePcClient({ enabled: false })
  expect(f.client.status()).toBe('disconnected:protocol_violation')
  expect(f.governanceProxy).not.toHaveBeenCalled()
})

it.each([
  'project_lifecycle_initialize',
  'project_lifecycle_read',
  'project_lifecycle_set_policy',
  'project_lifecycle_export_audit',
])('preserves original presentation channel lifecycle operation %s', async (operation) => {
  const f = await governancePcClient()
  f.socket.message(governanceFrame(operation, 'presentation.v1'))
  await vi.waitFor(() => expect(f.presentationProxy).toHaveBeenCalledOnce())
  expect(f.governanceProxy).not.toHaveBeenCalled()
  f.client.revoke()
})
it('real governance factory uses the existing project lock and actual userData store', async () => {
  const root = mkdtempSync(join(tmpdir(), 'governance-relay-factory-'))
  let release: (() => void) | undefined
  try {
    const plan = benchmarkPlan(),
      scope = { projectId: plan.projectId, documentId: 'synthetic-doc' }
    new PresentationStore(root).savePlan(scope.projectId, scope.documentId, 0, plan)
    const proxy = createOfficePresentationGovernanceProxy({ userDataPath: root })
    release = await acquirePresentationProjectLock(root, scope.projectId)
    let settled = false
    const pending = proxy(
      { operation: 'project_lifecycle_initialize', ...scope },
      new AbortController().signal,
    ).then((bytes) => {
      settled = true
      return JSON.parse(Buffer.from(bytes).toString())
    })
    await new Promise((resolve) => setTimeout(resolve, 20))
    expect(settled).toBe(false)
    expect(new PresentationLifecycleStore(root).readControl(scope)).toBeUndefined()
    release()
    release = undefined
    expect(await pending).toMatchObject({ lifecycle: { state: 'active', revision: 0 } })
  } finally {
    release?.()
    rmSync(root, { recursive: true, force: true })
  }
})

it.each([
  'agent.v1',
  'presentation.v1',
  'presentation-attachments.v1',
  'presentation-assets.v1',
  'presentation-remote-images.v1',
  'presentation-webpages.v1',
  'presentation-asset-rights.v1',
  'presentation-animation-frame.v1',
  'presentation-pdf.v1',
  'presentation-production-pdf.v1',
  'presentation-master-backups.v1',
  'presentation-package-backups.v1',
])('refuses deletion on negotiated ordinary capability %s', async (capability) => {
  const f = await governancePcClient({ capabilities: [capability] })
  f.socket.message(governanceFrame('project_deletion_confirm', capability))
  await vi.waitFor(() => expect(f.client.status()).toBe('disconnected:protocol_violation'))
  expect(f.presentationProxy).not.toHaveBeenCalled()
  expect(f.governanceProxy).not.toHaveBeenCalled()
})
it('rejects project deletion over legacy agent protocol', async () => {
  const { client, socket, proxy } = setup()
  const claimed = client.claim('123456')
  await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
  socket.open()
  await claimed
  socket.message({
    version: 1,
    type: 'pc.claimed',
    pairing_id: 'pairing_12345678',
    host: 'PowerPoint',
    origin: 'https://office.8-216-134-194.sslip.io',
    verification_code: '123456',
    expires_in: 120,
  })
  await client.approve('pairing_12345678')
  socket.message({
    version: 1,
    type: 'pc.approved',
    session_id: 'session_12345678',
    capability: 'secret-capability',
    expires_in: 1800,
  })
  socket.message({
    version: 1,
    type: 'relay.request',
    session_id: 'session_12345678',
    request_id: 'legacy-delete',
    body: { operation: 'project_deletion_confirm' },
  })
  await vi.waitFor(() => expect(client.status()).toBe('disconnected:protocol_violation'))
  expect(proxy).not.toHaveBeenCalled()
})
it('bounds governance input to 32KiB before invoking its handler', async () => {
  const f = await governancePcClient()
  f.socket.message(
    governanceFrame('project_lifecycle_read', 'presentation-governance.v1', {
      body: { operation: 'project_lifecycle_read', padding: 'x'.repeat(32 * 1024) },
    }),
  )
  await vi.waitFor(() => expect(f.client.status()).toBe('disconnected:protocol_violation'))
  expect(f.governanceProxy).not.toHaveBeenCalled()
})
it.each([0, 1])(
  'preserves exact maximum governance wire response and rejects extra byte %s',
  async (extra) => {
    const f = await governancePcClient({
      handler: async () => new Uint8Array(2 * 1024 * 1024 + 14 + extra),
    })
    f.socket.message(governanceFrame('project_lifecycle_read'))
    await vi.waitFor(() =>
      expect(
        f.socket.sent.some((s) => JSON.parse(s).type === (extra ? 'pc.error' : 'pc.done')),
      ).toBe(true),
    )
    const frames = f.socket.sent.map((s) => JSON.parse(s))
    expect(frames.some((s) => s.type === 'pc.done')).toBe(!extra)
    f.client.revoke()
  },
)

it('offers the known 18-capability catalog while negotiating a 17-capability PowerPoint subset', async () => {
  const socket = new FakeSocket()
  const client = createOfficeRelayClient({
    endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
    connect: () => socket,
    getValidAccountStatus: async () => ({ loggedIn: true }),
    getAccessToken: async () => 'token',
    proxy: async () => ({ status: 200, body: new Uint8Array() }),
    retrievalProxy: async () => new Uint8Array(),
    presentationProxy: async () => new Uint8Array(),
    presentationGovernanceProxy: async () => new TextEncoder().encode('{}'),
    supportsTeamPresentation: true,
    onPending() {},
  })
  const claiming = client.claim('123456')
  await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
  socket.open()
  await claiming
  const offered = JSON.parse(socket.sent[0]!).capabilities as string[]
  expect(offered).toHaveLength(18)
  expect(new Set(offered).size).toBe(18)
  expect(offered).toContain('presentation-team.v1')
  expect(offered).toContain('presentation-governance.v1')
  const primary = offered.filter((value) => value !== 'presentation-team.v1')
  expect(primary).toHaveLength(17)
  socket.message({ version: 2, type: 'pc.negotiated', pairing_version: 2, capabilities: primary })
  socket.message({
    version: 2,
    type: 'pc.claimed',
    pairing_id: 'pairing_12345678',
    host: 'PowerPoint',
    origin: 'https://office.8-216-134-194.sslip.io',
    verification_code: '123456',
    expires_in: 120,
    capabilities: primary,
  })
  expect(client.status()).not.toContain('protocol_violation')
  await client.approve('pairing_12345678')
  socket.message({
    version: 2,
    type: 'pc.approved',
    session_id: 'session_12345678',
    capability: 'secret-capability',
    expires_in: 1800,
    capabilities: primary,
  })
  expect(client.status()).not.toContain('protocol_violation')
  client.revoke()
})
