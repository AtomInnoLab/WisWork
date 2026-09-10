import { describe, expect, it, vi } from 'vitest'
import {
  OFFICE_RELAY_URL,
  createOfficeRelaySession,
  officeTransportMode,
  type RelayWebSocket,
} from '../src/relay/session.js'
import type { OfficeDiagnosticEvent } from '../src/diagnostics/office-diagnostics.js'

class FakeSocket implements RelayWebSocket {
  static readonly OPEN = 1
  readonly OPEN = 1
  readyState = 0
  sent: string[] = []
  onopen: (() => void) | null = null
  onmessage: ((event: { data: unknown }) => void) | null = null
  onclose: (() => void) | null = null
  onerror: (() => void) | null = null
  send(value: string) {
    this.sent.push(value)
  }
  close() {
    this.readyState = 3
    this.onclose?.()
  }
  open() {
    this.readyState = 1
    this.onopen?.()
  }
  receive(value: unknown) {
    this.onmessage?.({ data: value })
  }
}

const frame = (socket: FakeSocket, index: number) => JSON.parse(socket.sent[index]!)
const flushFrames = async () => {
  for (let turn = 0; turn < 4; turn += 1) await Promise.resolve()
}

async function connectedEnhancedSession(lifetimeMs = 60_000) {
  const socket = new FakeSocket()
  const session = createOfficeRelaySession({
    createSocket: () => socket,
    persistentPairing: false,
    capabilities: ['agent.v1'],
    randomUUID: () => 'request_12345678',
  })
  const toolHandler = vi.fn(async () => ({ output: 'ok' }))
  session.setToolHandler?.(toolHandler)
  const connecting = session.connect('powerpoint')
  socket.open()
  socket.receive(
    JSON.stringify({
      version: 2,
      type: 'office.created',
      pairing_id: 'pair_12345678',
      verification_code: '123456',
      expires_in: 120,
    }),
  )
  socket.receive(
    JSON.stringify({
      version: 2,
      type: 'office.approved',
      session_id: 'session_12345678',
      capability: 'capability_12345678',
      expires_in: 1800,
      capabilities: ['agent.v1'],
    }),
  )
  await connecting
  socket.receive(
    JSON.stringify({
      version: 2,
      type: 'relay.session_state',
      session_id: 'session_12345678',
      generation: 1,
      enhanced: {
        version: 1,
        runtime_mode: 'enhanced',
        runtime_instance: 'runtime_0123456789abcdef',
        component_version: '0.147.0',
        host: 'office-powerpoint',
        raw_office: false,
        expires_at: Date.now() + lifetimeMs,
        policy_generation: 1,
        session_generation: 1,
      },
    }),
  )
  await flushFrames()
  return { socket, session, toolHandler }
}

describe('Office cloud relay session', () => {
  it('keeps a multi-step agent request past five minutes and cancels at its own total deadline', async () => {
    vi.useFakeTimers()
    try {
      const { session, socket } = await connectedEnhancedSession(40 * 60_000)
      const response = session.authenticatedFetch('/v1/office/messages', {
        method: 'POST',
        body: '{}',
      })
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.start',
          session_id: 'session_12345678',
          request_id: 'request_12345678',
          status: 200,
          content_type: 'text/event-stream',
        }),
      )
      await flushFrames()
      const body = (await response).text()
      const rejected = expect(body).rejects.toThrow('relay_timeout')
      await vi.advanceTimersByTimeAsync(360_000)
      expect(
        socket.sent.map((value) => JSON.parse(value)).filter((f) => f.type === 'office.cancel'),
      ).toHaveLength(0)
      expect(session.snapshot().status).toBe('connected')
      await vi.advanceTimersByTimeAsync(30 * 60_000 + 10_000 - 360_000)
      await rejected
      expect(
        socket.sent.map((value) => JSON.parse(value)).filter((f) => f.type === 'office.cancel'),
      ).toHaveLength(1)
      expect(session.snapshot().status).toBe('connected')
      session.disconnect()
    } finally {
      vi.useRealTimers()
    }
  })
  const diagnostic: OfficeDiagnosticEvent = {
    event_id: '00000000-0000-4000-8000-000000000001',
    trace_id: '00000000-0000-4000-8000-000000000002',
    timestamp_ms: 1_777_000_000_000,
    host: 'word',
    platform: 'mac',
    build: 'build-123',
    tool: 'write_document',
    phase: 'write',
    outcome: 'failed',
    error_code: 'office_write_failed',
    office_error_code: 'InvalidArgument',
    office_error_location: 'Body.insertText',
    duration_ms: 25,
    requirement_sets: { WordApi: true },
  }

  it('uses the fixed secure relay by default and loopback only as explicit rollback', () => {
    expect(OFFICE_RELAY_URL).toBe('wss://office.8-216-134-194.sslip.io/office-relay')
    expect(officeTransportMode({})).toBe('relay')
    expect(officeTransportMode({ VITE_WISWORK_OFFICE_TRANSPORT: 'loopback' })).toBe('loopback')
    expect(() => officeTransportMode({ VITE_WISWORK_OFFICE_TRANSPORT: 'http' })).toThrow(
      'invalid_office_transport',
    )
  })

  it('pairs on one WSS socket and exposes only the six-digit verification code', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: (url) => {
        expect(url).toBe(OFFICE_RELAY_URL)
        return socket
      },
    })
    const connecting = session.connect('word')
    expect(session.snapshot()).toEqual({ status: 'connecting' })
    socket.open()
    expect(frame(socket, 0)).toEqual({ version: 1, type: 'office.create', host: 'Word' })
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'office.created',
        pairing_id: 'pair_1',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    await flushFrames()
    expect(session.snapshot()).toEqual({ status: 'pending', verificationCode: '123456' })
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'office.approved',
        session_id: 'session_1',
        capability: 'cap_1',
        expires_in: 1800,
      }),
    )
    await connecting
    expect(session.snapshot()).toEqual({ status: 'connected' })
    socket.receive(JSON.stringify({ version: 1, type: 'relay.error', code: 'session_expired' }))
    await flushFrames()
    expect(session.snapshot()).toEqual({ status: 'expired' })
  })

  it('advertises only Relay-v2 capabilities and blocks unnegotiated requests', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1', 'web-search.v1', 'web-fetch.v1'],
      randomUUID: () => 'request_12345678',
    })
    const connecting = session.connect('word')
    socket.open()
    expect(frame(socket, 0)).toEqual({
      version: 2,
      type: 'office.create',
      host: 'Word',
      capabilities: ['agent.v1', 'web-search.v1', 'web-fetch.v1'],
    })
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_12345678',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_12345678',
        capability: 'capability_12345678',
        expires_in: 1800,
        capabilities: ['agent.v1', 'web-search.v1'],
      }),
    )
    await connecting
    expect(session.snapshot()).toEqual({
      status: 'connected',
      capabilities: ['agent.v1', 'web-search.v1'],
    })
    await expect(
      session.capabilityFetch('web-fetch.v1', { url: 'https://example.com' }),
    ).rejects.toThrow('relay_capability_unavailable')
    const pending = session.capabilityFetch('web-search.v1', { query: 'office', max_results: 3 })
    expect(frame(socket, 1)).toEqual({
      version: 2,
      type: 'office.request',
      session_id: 'session_12345678',
      capability: 'capability_12345678',
      request_id: 'request_12345678',
      capability_name: 'web-search.v1',
      body: { query: 'office', max_results: 3 },
    })
    session.disconnect()
    await expect(pending).rejects.toThrow('relay_disconnected')
  })

  it('routes exact session-bound Enhanced tool subframes through the existing agent capability', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1'],
      randomUUID: () => 'request_12345678',
    })
    const calls: string[] = []
    session.setToolHandler?.(async (call) => {
      calls.push(call.toolName)
      return { output: JSON.stringify({ title: 'Document' }) }
    })
    const connecting = session.connect('word')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_12345678',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_12345678',
        capability: 'capability_12345678',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connecting
    const enhanced = {
      version: 1,
      runtime_mode: 'enhanced',
      runtime_instance: 'runtime_0123456789abcdef',
      component_version: '0.147.0',
      host: 'office-word',
      raw_office: false,
      expires_at: Date.now() + 60_000,
      policy_generation: 2,
      session_generation: 4,
    }
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.session_state',
        session_id: 'session_12345678',
        generation: 4,
        enhanced,
      }),
    )
    await flushFrames()
    expect(session.snapshot()).toMatchObject({ status: 'connected', enhanced })
    const modeUpdate = vi.fn()
    const unsubscribeModeUpdate = session.subscribe(modeUpdate)
    const pending = session.capabilityFetch('agent.v1', { messages: [] })
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.tool_call',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        turn_id: 'turn_12345678',
        call_id: 'call_12345678',
        generation: 4,
        tool_name: 'read_document',
        input: {},
      }),
    )
    await flushFrames()
    expect(modeUpdate).toHaveBeenCalled()
    expect(session.snapshot()).toMatchObject({ status: 'connected', enhanced })
    unsubscribeModeUpdate()
    expect(calls).toEqual(['read_document'])
    expect(frame(socket, 2)).toEqual({
      version: 2,
      type: 'office.tool_result',
      session_id: 'session_12345678',
      capability: 'capability_12345678',
      request_id: 'request_12345678',
      turn_id: 'turn_12345678',
      call_id: 'call_12345678',
      generation: 4,
      output: JSON.stringify({ title: 'Document' }),
      is_error: false,
    })
    session.disconnect()
    await expect(pending).rejects.toThrow('relay_disconnected')
  })

  it('allows concurrent Enhanced read tool subframes in the same request', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1'],
      randomUUID: () => 'request_12345678',
    })
    const resolvers = new Map<string, (result: { output: string }) => void>()
    session.setToolHandler?.(
      (call) =>
        new Promise((resolve) => {
          resolvers.set(call.callId, resolve)
        }),
    )
    const connecting = session.connect('powerpoint')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_12345678',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_12345678',
        capability: 'capability_12345678',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connecting
    const enhanced = {
      version: 1,
      runtime_mode: 'enhanced',
      runtime_instance: 'runtime_0123456789abcdef',
      component_version: '0.147.0',
      host: 'office-powerpoint',
      raw_office: false,
      expires_at: Date.now() + 60_000,
      policy_generation: 2,
      session_generation: 4,
    }
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.session_state',
        session_id: 'session_12345678',
        generation: 4,
        enhanced,
      }),
    )
    await flushFrames()
    const pending = session.capabilityFetch('agent.v1', { messages: [] }).catch((error) => error)
    for (const [callId, toolName] of [
      ['call_state_12345678', 'get_presentation_state'],
      ['call_shapes_12345678', 'list_slide_shapes'],
    ] as const)
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.tool_call',
          session_id: 'session_12345678',
          request_id: 'request_12345678',
          turn_id: 'turn_12345678',
          call_id: callId,
          generation: 4,
          tool_name: toolName,
          input: {},
        }),
      )
    await flushFrames()
    await flushFrames()
    expect(session.snapshot()).toMatchObject({ status: 'connected', enhanced })
    expect([...resolvers.keys()]).toEqual(['call_state_12345678', 'call_shapes_12345678'])
    resolvers.get('call_state_12345678')?.({ output: '{"slideCount":8}' })
    resolvers.get('call_shapes_12345678')?.({ output: '{"shapes":[]}' })
    await flushFrames()
    expect([frame(socket, 2), frame(socket, 3)]).toMatchObject([
      { type: 'office.tool_result', call_id: 'call_state_12345678', is_error: false },
      { type: 'office.tool_result', call_id: 'call_shapes_12345678', is_error: false },
    ])
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.tool_call',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        turn_id: 'turn_12345678',
        call_id: 'call_state_12345678',
        generation: 4,
        tool_name: 'get_presentation_state',
        input: {},
      }),
    )
    await flushFrames()
    expect(session.snapshot()).toEqual({ status: 'offline' })
    await expect(pending).resolves.toMatchObject({ message: 'relay_disconnected' })
  })

  it('cancels in-flight Enhanced tools when the runtime generation changes', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1'],
      randomUUID: () => 'request_12345678',
    })
    let signal: AbortSignal | undefined
    let resolveTool: ((result: { output: string }) => void) | undefined
    let toolCalls = 0
    session.setToolHandler?.(
      (call) =>
        new Promise((resolve) => {
          toolCalls += 1
          signal = call.signal
          resolveTool = resolve
        }),
    )
    const connecting = session.connect('powerpoint')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_12345678',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_12345678',
        capability: 'capability_12345678',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connecting
    const statement = (generation: number) => ({
      version: 1,
      runtime_mode: 'enhanced',
      runtime_instance: 'runtime_0123456789abcdef',
      component_version: '0.147.0',
      host: 'office-powerpoint',
      raw_office: false,
      expires_at: Date.now() + 60_000,
      policy_generation: 2,
      session_generation: generation,
    })
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.session_state',
        session_id: 'session_12345678',
        generation: 4,
        enhanced: statement(4),
      }),
    )
    await flushFrames()
    const pending = session.capabilityFetch('agent.v1', { messages: [] }).catch((error) => error)
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.tool_call',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        turn_id: 'turn_12345678',
        call_id: 'call_state_12345678',
        generation: 4,
        tool_name: 'get_presentation_state',
        input: {},
      }),
    )
    await flushFrames()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.session_state',
        session_id: 'session_12345678',
        generation: 5,
        enhanced: statement(5),
      }),
    )
    await flushFrames()
    expect(signal?.aborted).toBe(true)
    resolveTool?.({ output: '{"slideCount":8}' })
    await flushFrames()
    expect(socket.sent).toHaveLength(2)
    expect(session.snapshot()).toMatchObject({
      status: 'connected',
      enhanced: { session_generation: 5 },
    })
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.tool_call',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        turn_id: 'turn_12345678',
        call_id: 'call_state_12345678',
        generation: 5,
        tool_name: 'get_presentation_state',
        input: {},
      }),
    )
    await flushFrames()
    expect(toolCalls).toBe(1)
    expect(session.snapshot()).toEqual({ status: 'offline' })
    await expect(pending).resolves.toMatchObject({ message: 'relay_disconnected' })
  })

  it('revokes the session on replayed or generation-drifted tool subframes', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1'],
      randomUUID: () => 'request_12345678',
    })
    session.setToolHandler?.(async () => new Promise(() => undefined))
    const connecting = session.connect('excel')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_12345678',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_12345678',
        capability: 'capability_12345678',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connecting
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.session_state',
        session_id: 'session_12345678',
        generation: 9,
        enhanced: {
          version: 1,
          runtime_mode: 'enhanced',
          runtime_instance: 'runtime_0123456789abcdef',
          component_version: '0.147.0',
          host: 'office-excel',
          raw_office: false,
          expires_at: Date.now() + 60_000,
          policy_generation: 2,
          session_generation: 9,
        },
      }),
    )
    await flushFrames()
    void session.capabilityFetch('agent.v1', { messages: [] }).catch(() => undefined)
    await Promise.resolve()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.tool_call',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        turn_id: 'turn_12345678',
        call_id: 'call_12345678',
        generation: 8,
        tool_name: 'read_workbook',
        input: {},
      }),
    )
    await flushFrames()
    expect(session.snapshot()).toEqual({ status: 'offline' })
  })

  it.each([true, false])(
    'revokes an expired Enhanced statement with active request: %s',
    async (active) => {
      vi.useFakeTimers()
      vi.setSystemTime(1_000)
      try {
        const { socket, session } = await connectedEnhancedSession()
        const pending = active
          ? session.capabilityFetch('agent.v1', { messages: [] }).catch((error) => error.message)
          : undefined
        await vi.advanceTimersByTimeAsync(60_000)
        expect(session.snapshot()).toEqual({ status: 'offline' })
        if (active) await expect(pending).resolves.toBe('relay_disconnected')
        session.disconnect()
        expect(
          socket.sent
            .map((value) => JSON.parse(value))
            .filter((value) => value.type === 'office.cancel'),
        ).toEqual(
          active
            ? [
                {
                  version: 2,
                  type: 'office.cancel',
                  session_id: 'session_12345678',
                  capability: 'capability_12345678',
                  request_id: 'request_12345678',
                },
              ]
            : [],
        )
      } finally {
        vi.useRealTimers()
      }
    },
  )

  it.each([
    { size: 100 * 1024, accepted: true },
    { size: 256 * 1024 + 1, accepted: false },
    { size: 272 * 1024, accepted: false },
  ])('bounds Enhanced plan input at the tool limit: $size bytes', async ({ size, accepted }) => {
    const { socket, session, toolHandler } = await connectedEnhancedSession()
    const pending = session
      .capabilityFetch('agent.v1', { messages: [] })
      .catch((error) => error.message)
    const input = { design_contract: { design_md: 'x'.repeat(size) } }
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.tool_call',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        turn_id: 'turn_12345678',
        call_id: 'call_12345678',
        generation: 1,
        tool_name: 'plan_deck',
        input,
      }),
    )
    await flushFrames()
    expect(session.snapshot().status).toBe(accepted ? 'connected' : 'offline')
    expect(toolHandler).toHaveBeenCalledTimes(accepted ? 1 : 0)
    if (accepted)
      expect(toolHandler).toHaveBeenCalledWith(
        expect.objectContaining({ toolName: 'plan_deck', input }),
      )
    session.disconnect()
    await expect(pending).resolves.toBe('relay_disconnected')
  })

  it.each(['relay.chunk', 'relay.session_state'])(
    'keeps oversized non-tool frames fail-closed: %s',
    async (type) => {
      const { socket, session, toolHandler } = await connectedEnhancedSession()
      const pending = session.capabilityFetch('agent.v1', { messages: [] })
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.start',
          session_id: 'session_12345678',
          request_id: 'request_12345678',
          status: 200,
          content_type: 'text/event-stream',
        }),
      )
      const response = await pending
      const body = response.text().catch((error: Error) => error.message)
      socket.receive(
        JSON.stringify({
          version: 2,
          type,
          session_id: 'session_12345678',
          request_id: 'request_12345678',
          sequence: 0,
          data: 'x'.repeat(100 * 1024),
        }),
      )
      await flushFrames()
      expect(session.snapshot().status).toBe('offline')
      expect(toolHandler).not.toHaveBeenCalled()
      await expect(body).resolves.toBe('relay_disconnected')
    },
  )

  it.each([
    { reason: 'disconnect', active: true },
    { reason: 'disconnect', active: false },
    { reason: 'protocol', active: true },
    { reason: 'protocol', active: false },
    { reason: 'cancel_send_failure', active: true },
  ])('cancels once before revoking: $reason, active: $active', async ({ reason, active }) => {
    const { socket, session } = await connectedEnhancedSession()
    const pending = active
      ? session.capabilityFetch('agent.v1', { messages: [] }).catch((error) => error.message)
      : undefined
    const send = vi.spyOn(socket, 'send')
    const close = vi.spyOn(socket, 'close')
    if (reason === 'cancel_send_failure')
      send.mockImplementation(() => {
        throw new Error('socket_failed')
      })
    if (reason === 'protocol') {
      socket.receive('{')
      await flushFrames()
    } else session.disconnect()
    session.disconnect()
    expect(session.snapshot().status).toBe('offline')
    expect(socket.readyState).toBe(3)
    if (active) await expect(pending).resolves.toBe('relay_disconnected')
    expect(send.mock.calls.map(([value]) => JSON.parse(value))).toEqual(
      active
        ? [
            {
              version: 2,
              type: 'office.cancel',
              session_id: 'session_12345678',
              capability: 'capability_12345678',
              request_id: 'request_12345678',
            },
          ]
        : [],
    )
    if (active)
      expect(send.mock.invocationCallOrder[0]).toBeLessThan(close.mock.invocationCallOrder[0]!)
  })

  it('sends bounded diagnostics only over an approved v2 session', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1'],
    })
    const connecting = session.connect('word')
    await expect(session.sendDiagnostic(diagnostic)).rejects.toThrow('diagnostic_unavailable')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_12345678',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_12345678',
        capability: 'capability_12345678',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connecting
    const accepted = session.sendDiagnostic(diagnostic)
    expect(frame(socket, 1)).toEqual({
      version: 2,
      type: 'office.diagnostic',
      session_id: 'session_12345678',
      capability: 'capability_12345678',
      ...diagnostic,
    })
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.diagnostic.accepted',
        event_id: diagnostic.event_id,
      }),
    )
    await expect(accepted).resolves.toBeUndefined()
    expect(session.snapshot().status).toBe('connected')
    const staged = session.sendDiagnostic({
      ...diagnostic,
      event_id: '00000000-0000-4000-8000-000000000003',
      error_code: 'office_recovery_failed:word_body_shape',
      verification_stage: 'body_shape',
    })
    expect(frame(socket, 2).error_code).toBe('office_recovery_failed')
    expect(frame(socket, 2)).not.toHaveProperty('verification_stage')
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.diagnostic.accepted',
        event_id: '00000000-0000-4000-8000-000000000003',
      }),
    )
    await expect(staged).resolves.toBeUndefined()
    const concurrent = session.sendDiagnostic({
      ...diagnostic,
      event_id: '00000000-0000-4000-8000-000000000004',
      error_code: 'office_concurrent_change',
    })
    expect(frame(socket, 3).error_code).toBe('office_verify_failed')
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.diagnostic.accepted',
        event_id: '00000000-0000-4000-8000-000000000004',
      }),
    )
    await expect(concurrent).resolves.toBeUndefined()
    const invalidInput = session.sendDiagnostic({
      ...diagnostic,
      event_id: '00000000-0000-4000-8000-000000000005',
      error_code: 'invalid_tool_input',
    })
    expect(frame(socket, 4).error_code).toBe('agent_run_failed')
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.diagnostic.accepted',
        event_id: '00000000-0000-4000-8000-000000000005',
      }),
    )
    await expect(invalidInput).resolves.toBeUndefined()
    for (const [index, errorCode] of [
      'image_fetch_unavailable',
      'image_limit',
      'image_mime_unsupported',
      'invalid_image',
    ].entries()) {
      const eventId = `00000000-0000-4000-8000-00000000000${index + 6}`
      const upload = session.sendDiagnostic({
        ...diagnostic,
        event_id: eventId,
        error_code: errorCode,
      })
      expect(frame(socket, index + 5).error_code).toBe(
        errorCode === 'image_fetch_unavailable' ? 'network_error' : 'agent_run_failed',
      )
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'office.diagnostic.accepted',
          event_id: eventId,
        }),
      )
      await expect(upload).resolves.toBeUndefined()
    }
    await expect(
      session.sendDiagnostic({ ...diagnostic, tool: 'x'.repeat(5_000) }),
    ).rejects.toThrow('diagnostic_too_large')
    expect(socket.sent).toHaveLength(9)
  })

  it('keeps Agent streaming usable after a nonfatal diagnostic limit response', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1'],
      randomUUID: () => 'request_12345678',
    })
    const connecting = session.connect('excel')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_12345678',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_12345678',
        capability: 'capability_12345678',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connecting
    const rejected = session.sendDiagnostic({ ...diagnostic, host: 'excel' })
    socket.receive(
      JSON.stringify({ version: 2, type: 'relay.error', code: 'diagnostic_rate_limited' }),
    )
    await expect(rejected).rejects.toThrow('diagnostic_rate_limited')
    const hostMismatch = session.sendDiagnostic({
      ...diagnostic,
      event_id: '00000000-0000-4000-8000-000000000003',
      host: 'excel',
    })
    socket.receive(
      JSON.stringify({ version: 2, type: 'relay.error', code: 'diagnostic_host_mismatch' }),
    )
    await expect(hostMismatch).rejects.toThrow('diagnostic_host_mismatch')
    expect(session.snapshot().status).toBe('connected')
    const pending = session.authenticatedFetch('/v1/office/messages', {
      method: 'POST',
      body: '{"model":"fixed"}',
    })
    expect(frame(socket, 3).type).toBe('office.request')
    session.disconnect()
    await expect(pending).rejects.toThrow('relay_disconnected')
  })

  it('does not send diagnostics over a v1 rollback session', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: () => socket })
    const connecting = session.connect('word')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'office.created',
        pairing_id: 'pair',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'office.approved',
        session_id: 'session',
        capability: 'capability',
        expires_in: 1800,
      }),
    )
    await connecting
    await expect(session.sendDiagnostic(diagnostic)).rejects.toThrow('diagnostic_unavailable')
    expect(socket.sent).toHaveLength(1)
  })

  it('streams bounded SSE events, sends cancel, and completes once', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: () => socket })
    const connecting = session.connect('excel')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'office.created',
        pairing_id: 'p',
        verification_code: '654321',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'office.approved',
        session_id: 'session',
        capability: 'cap',
        expires_in: 1800,
      }),
    )
    await connecting
    const controller = new AbortController()
    const responsePending = session.authenticatedFetch('/v1/office/messages', {
      method: 'POST',
      body: '{"model":"fixed"}',
      signal: controller.signal,
    })
    const request = frame(socket, 1)
    expect(request).toMatchObject({
      version: 1,
      type: 'office.request',
      session_id: 'session',
      capability: 'cap',
      body: { model: 'fixed' },
    })
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'relay.start',
        session_id: 'session',
        request_id: request.request_id,
        status: 200,
        content_type: 'text/event-stream',
      }),
    )
    const response = await responsePending
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'relay.chunk',
        session_id: 'session',
        request_id: request.request_id,
        sequence: 0,
        data: btoa('data: [DONE]\n'),
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'relay.done',
        session_id: 'session',
        request_id: request.request_id,
      }),
    )
    await expect(response.text()).resolves.toBe('data: [DONE]\n')
    controller.abort()
    controller.abort()
    expect(
      socket.sent
        .map((value) => JSON.parse(value))
        .filter((value) => value.type === 'office.cancel'),
    ).toHaveLength(0)
  })

  it('cancels an active request, revokes on disconnect, and can pair again', async () => {
    const made: FakeSocket[] = []
    const session = createOfficeRelaySession({
      createSocket: () => {
        const socket = new FakeSocket()
        made.push(socket)
        return socket
      },
    })
    const first = session.connect('powerpoint')
    made[0]!.open()
    made[0]!.receive(
      JSON.stringify({
        version: 1,
        type: 'office.created',
        pairing_id: 'p1',
        verification_code: '111111',
        expires_in: 120,
      }),
    )
    made[0]!.receive(
      JSON.stringify({
        version: 1,
        type: 'office.approved',
        session_id: 'session1',
        capability: 'c1',
        expires_in: 1800,
      }),
    )
    await first
    const controller = new AbortController()
    const cancelled = session.authenticatedFetch('/v1/office/messages', {
      method: 'POST',
      body: '{}',
      signal: controller.signal,
    })
    const requestFrame = frame(made[0]!, 1)
    controller.abort()
    await expect(cancelled).rejects.toThrow('relay_cancelled')
    expect(
      made[0]!.sent
        .map((value) => JSON.parse(value))
        .filter((value) => value.type === 'office.cancel'),
    ).toHaveLength(1)
    made[0]!.receive(
      JSON.stringify({
        version: 1,
        type: 'relay.error',
        session_id: 'session1',
        request_id: requestFrame.request_id,
        code: 'request_timeout',
      }),
    )
    expect(session.snapshot()).toEqual({ status: 'connected' })
    made[0]!.close()
    expect(session.snapshot()).toEqual({ status: 'offline' })

    const second = session.connect('powerpoint')
    made[1]!.open()
    made[1]!.receive(
      JSON.stringify({
        version: 1,
        type: 'office.created',
        pairing_id: 'p2',
        verification_code: '222222',
        expires_in: 120,
      }),
    )
    made[1]!.receive(
      JSON.stringify({
        version: 1,
        type: 'office.approved',
        session_id: 'session2',
        capability: 'c2',
        expires_in: 1800,
      }),
    )
    await second
    expect(session.snapshot()).toEqual({ status: 'connected' })
  })

  it('fails closed on binary, malformed, extra-key, oversized, and unknown frames', async () => {
    for (const hostile of [
      new Uint8Array([1]),
      '{',
      JSON.stringify({ version: 1, type: 'office.created', extra: true }),
      'x'.repeat(16 * 1024 + 1),
    ]) {
      const socket = new FakeSocket()
      const session = createOfficeRelaySession({ createSocket: () => socket })
      const connecting = session.connect('word')
      socket.open()
      socket.receive(hostile)
      await connecting
      expect(session.snapshot()).toEqual({ status: 'offline' })
      expect(socket.readyState).toBe(3)
    }
  })

  it('rejects oversized requests before send and chunks before relay.start', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: () => socket })
    const connecting = session.connect('word')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'office.created',
        pairing_id: 'p',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'office.approved',
        session_id: 'sid',
        capability: 'cap',
        expires_in: 1800,
      }),
    )
    await connecting
    await expect(
      session.authenticatedFetch('/v1/office/messages', {
        method: 'POST',
        body: 'x'.repeat(256 * 1024 + 1),
      }),
    ).rejects.toThrow('relay_request_too_large')
    expect(socket.sent).toHaveLength(1)

    const response = session.authenticatedFetch('/v1/office/messages', {
      method: 'POST',
      body: '{}',
    })
    const request = frame(socket, 1)
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'relay.chunk',
        session_id: 'sid',
        request_id: request.request_id,
        sequence: 0,
        data: btoa('data'),
      }),
    )
    await expect(response).rejects.toThrow('relay_disconnected')
    expect(session.snapshot()).toEqual({ status: 'offline' })
  })

  it.each(['{', '[]', 'null', '"text"'])(
    'rejects non-object request JSON %s before send',
    async (body) => {
      const socket = new FakeSocket()
      const session = createOfficeRelaySession({ createSocket: () => socket })
      const connecting = session.connect('word')
      socket.open()
      socket.receive(
        JSON.stringify({
          version: 1,
          type: 'office.created',
          pairing_id: 'p',
          verification_code: '123456',
          expires_in: 120,
        }),
      )
      socket.receive(
        JSON.stringify({
          version: 1,
          type: 'office.approved',
          session_id: 'sid',
          capability: 'cap',
          expires_in: 1800,
        }),
      )
      await connecting
      await expect(
        session.authenticatedFetch('/v1/office/messages', { method: 'POST', body }),
      ).rejects.toThrow('relay_invalid_request')
      expect(socket.sent).toHaveLength(1)
    },
  )

  it.each([204, 205, 304])(
    'fails closed when relay.start has non-streaming status %i',
    async (status) => {
      const socket = new FakeSocket()
      const session = createOfficeRelaySession({ createSocket: () => socket })
      const connecting = session.connect('word')
      socket.open()
      socket.receive(
        JSON.stringify({
          version: 1,
          type: 'office.created',
          pairing_id: 'p',
          verification_code: '123456',
          expires_in: 120,
        }),
      )
      socket.receive(
        JSON.stringify({
          version: 1,
          type: 'office.approved',
          session_id: 'sid',
          capability: 'cap',
          expires_in: 1800,
        }),
      )
      await connecting
      const response = session.authenticatedFetch('/v1/office/messages', {
        method: 'POST',
        body: '{}',
      })
      const request = frame(socket, 1)
      socket.receive(
        JSON.stringify({
          version: 1,
          type: 'relay.start',
          session_id: 'sid',
          request_id: request.request_id,
          status,
          content_type: 'text/event-stream',
        }),
      )
      await expect(response).rejects.toThrow('relay_disconnected')
      expect(session.snapshot()).toEqual({ status: 'offline' })
    },
  )

  it('cleans up when request send throws and enforces pairing frame order', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({ createSocket: () => socket })
    const connecting = session.connect('word')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'office.approved',
        session_id: 'sid',
        capability: 'cap',
        expires_in: 1800,
      }),
    )
    await connecting
    expect(session.snapshot()).toEqual({ status: 'offline' })

    const throwing = new FakeSocket()
    const connected = createOfficeRelaySession({ createSocket: () => throwing })
    const pairing = connected.connect('word')
    throwing.open()
    throwing.receive(
      JSON.stringify({
        version: 1,
        type: 'office.created',
        pairing_id: 'p',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    throwing.receive(
      JSON.stringify({
        version: 1,
        type: 'office.approved',
        session_id: 'sid',
        capability: 'cap',
        expires_in: 1800,
      }),
    )
    await pairing
    throwing.send = () => {
      throw new Error('socket failed')
    }
    await expect(
      connected.authenticatedFetch('/v1/office/messages', { method: 'POST', body: '{}' }),
    ).rejects.toThrow('relay_disconnected')
    expect(connected.snapshot()).toEqual({ status: 'offline' })
  })
})
