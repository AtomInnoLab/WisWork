import { describe, expect, it, vi } from 'vitest'
import {
  OFFICE_RELAY_URL,
  createOfficeRelaySession,
  officeTransportMode,
  type RelayWebSocket,
  type OfficeRelayCapability,
} from '../src/relay/session.js'
import { type OfficeDiagnosticEvent } from '../src/diagnostics/office-diagnostics.js'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
import { type StructuredProposalController } from '../src/agent/proposal-controller.js'

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

async function connectedEnhancedSession(
  lifetimeMs = 60_000,
  capabilities: OfficeRelayCapability[] = ['agent.v1'],
  approvedCapabilities = capabilities,
) {
  const socket = new FakeSocket()
  let requestSequence = 0
  const session = createOfficeRelaySession({
    createSocket: () => socket,
    persistentPairing: false,
    capabilities,
    randomUUID: () =>
      capabilities.includes('design-document.v1')
        ? `request_design_${++requestSequence}`
        : 'request_12345678',
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
      capabilities: approvedCapabilities,
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
  it('renews a negotiated 15-minute Enhanced lease without cancelling its active tool or semantic proposal', async () => {
    vi.useFakeTimers()
    const { socket, session } = await connectedEnhancedSession(15 * 60_000, [
      'agent.v1',
      'enhanced-lease.v1',
    ])
    const runtime = createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        request: vi.fn(),
        documentId: async () => 'document-1',
        lastProject: () => undefined,
        rememberProject: async () => undefined,
      },
    })
    const proposals = runtime.proposals as StructuredProposalController
    const execute = vi.fn()
    // Matches App's raw_office=false effect, which also runs for a new lease object.
    const unsubscribe = session.subscribe(() => runtime.disableElevatedOffice())
    try {
      const original = session.snapshot().enhanced!
      let activeSignal: AbortSignal | undefined
      session.setToolHandler?.(async (call) => {
        activeSignal = call.signal
        const proposal = proposals.propose({
          operation: 'edit_slide_text',
          title: 'Update slide title',
          preview: { text: 'Renewed title' },
          impact: { host: 'powerpoint', targets: ['slide_1'], count: 1 },
          fingerprint: 'slide_1_revision',
          validate: () => !call.signal.aborted,
          execute,
        })
        call.signal.addEventListener('abort', () => proposals.logout(), { once: true })
        const decision = await proposals.waitForDecision(proposal.id)
        return { output: decision.status }
      })
      const pending = session.capabilityFetch('agent.v1', { messages: [] }).catch((error) => error)
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.tool_call',
          session_id: 'session_12345678',
          request_id: 'request_12345678',
          turn_id: 'turn_12345678',
          call_id: 'call_12345678',
          generation: 1,
          tool_name: 'read_document',
          input: {},
        }),
      )
      await flushFrames()
      const proposalId = proposals.pending()!.id
      await vi.advanceTimersByTimeAsync(10 * 60_000)
      const renewed = { ...original, expires_at: Date.now() + 15 * 60_000 }
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.session_state',
          session_id: 'session_12345678',
          generation: 1,
          enhanced: renewed,
        }),
      )
      await flushFrames()
      expect(session.snapshot()).toMatchObject({ status: 'connected', enhanced: renewed })
      await vi.advanceTimersByTimeAsync(6 * 60_000)
      expect(Date.now()).toBeGreaterThan(original.expires_at)
      expect(activeSignal?.aborted).toBe(false)
      expect(proposals.pending()?.id).toBe(proposalId)
      await proposals.confirm(proposalId)
      expect(execute).toHaveBeenCalledOnce()
      await expect(proposals.confirm(proposalId)).rejects.toThrow('proposal_missing')
      await flushFrames()
      expect(socket.sent.map((value) => JSON.parse(value))).toContainEqual(
        expect.objectContaining({
          type: 'office.tool_result',
          call_id: 'call_12345678',
          output: 'confirmed',
        }),
      )
      // The active request and its replay protection must both survive renewal.
      await expect(session.capabilityFetch('agent.v1', {})).rejects.toThrow('relay_busy')
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.tool_call',
          session_id: 'session_12345678',
          request_id: 'request_12345678',
          turn_id: 'turn_12345678',
          call_id: 'call_12345678',
          generation: 1,
          tool_name: 'read_document',
          input: {},
        }),
      )
      await flushFrames()
      expect(session.snapshot()).toEqual({ status: 'offline' })
      expect(execute).toHaveBeenCalledOnce()
      await expect(pending).resolves.toMatchObject({ message: 'relay_disconnected' })
    } finally {
      unsubscribe()
      runtime.dispose()
      session.disconnect()
      vi.useRealTimers()
    }
  })

  it.each([
    ['replayed expiry', {}],
    ['decreasing expiry', { expires_at: 1 }],
    ['overlong lease', { expires_at: Number.MAX_SAFE_INTEGER }],
    ['changed runtime', { runtime_instance: 'runtime_different_0123456789' }],
    ['changed component', { component_version: '0.148.0' }],
    ['changed host', { host: 'office-word' }],
    ['changed permissions', { raw_office: true }],
    ['changed policy', { policy_generation: 2 }],
    ['changed session', { session_generation: 2 }],
    ['changed mode', { runtime_mode: 'standard' }],
    ['changed version', { version: 2 }],
    ['unknown field', { unexpected: true }],
  ])('rejects a same-generation renewal with %s', async (_label, changed) => {
    const { socket, session } = await connectedEnhancedSession(15 * 60_000, [
      'agent.v1',
      'enhanced-lease.v1',
    ])
    try {
      const original = session.snapshot().enhanced!
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.session_state',
          session_id: 'session_12345678',
          generation: 1,
          enhanced: {
            ...original,
            expires_at:
              _label === 'replayed expiry' ? original.expires_at : original.expires_at + 1,
            ...changed,
          },
        }),
      )
      await flushFrames()
      expect(session.snapshot()).toEqual({ status: 'offline' })
    } finally {
      session.disconnect()
    }
  })

  it('rejects an offered but unnegotiated renewal capability', async () => {
    const { socket, session } = await connectedEnhancedSession(
      15 * 60_000,
      ['agent.v1', 'enhanced-lease.v1'],
      ['agent.v1'],
    )
    const original = session.snapshot().enhanced!
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.session_state',
        session_id: 'session_12345678',
        generation: 1,
        enhanced: { ...original, expires_at: original.expires_at + 1 },
      }),
    )
    await flushFrames()
    expect(session.snapshot()).toEqual({ status: 'offline' })
  })

  it.each(['expired', 'disconnected'] as const)('does not revive an %s lease', async (reason) => {
    vi.useFakeTimers()
    const { socket, session } = await connectedEnhancedSession(15 * 60_000, [
      'agent.v1',
      'enhanced-lease.v1',
    ])
    try {
      const original = session.snapshot().enhanced!
      if (reason === 'expired') vi.setSystemTime(original.expires_at)
      else session.disconnect()
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.session_state',
          session_id: 'session_12345678',
          generation: 1,
          enhanced: { ...original, expires_at: Date.now() + 15 * 60_000 },
        }),
      )
      await flushFrames()
      expect(session.snapshot()).toEqual({ status: 'offline' })
    } finally {
      session.disconnect()
      vi.useRealTimers()
    }
  })

  it('bounds renewal clock skew to 30 seconds and expires at the renewed deadline', async () => {
    vi.useFakeTimers()
    const { socket, session } = await connectedEnhancedSession(15 * 60_000, [
      'agent.v1',
      'enhanced-lease.v1',
    ])
    try {
      const original = session.snapshot().enhanced!
      await vi.advanceTimersByTimeAsync(10 * 60_000)
      const renewed = { ...original, expires_at: Date.now() + 15 * 60_000 + 30_000 }
      socket.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.session_state',
          session_id: 'session_12345678',
          generation: 1,
          enhanced: renewed,
        }),
      )
      await flushFrames()
      expect(session.snapshot().enhanced).toEqual(renewed)
      await vi.advanceTimersByTimeAsync(15 * 60_000 + 29_999)
      expect(session.snapshot().status).toBe('connected')
      await vi.advanceTimersByTimeAsync(1)
      expect(session.snapshot()).toEqual({ status: 'offline' })
    } finally {
      session.disconnect()
      vi.useRealTimers()
    }
  })

  it('never sends the negotiated lease control capability as a callable request', async () => {
    const { socket, session } = await connectedEnhancedSession(15 * 60_000, [
      'agent.v1',
      'enhanced-lease.v1',
    ])
    try {
      const sent = socket.sent.length
      await expect(session.capabilityFetch('enhanced-lease.v1', {})).rejects.toThrow(
        'relay_capability_unavailable',
      )
      expect(socket.sent).toHaveLength(sent)
      expect(session.snapshot().status).toBe('connected')
    } finally {
      session.disconnect()
    }
  })

  it.each([
    ['request_timeout', 'relay_request_timeout'],
    ['auth_required', 'relay_auth_required'],
    ['upstream_error', 'relay_upstream_error'],
    ['unknown_private_code', 'relay_error'],
  ])('retains only a known request failure category: %s', async (code, expected) => {
    const { socket, session } = await connectedEnhancedSession()
    const pending = session
      .capabilityFetch('agent.v1', { messages: [] })
      .catch((error) => error.message)
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.error',
        session_id: 'session_12345678',
        request_id: 'request_12345678',
        code,
      }),
    )
    await expect(pending).resolves.toBe(expected)
    session.disconnect()
  })
  it('lets a new model turn preempt a background DESIGN.md read without breaking pairing', async () => {
    const { session, socket } = await connectedEnhancedSession(60_000, [
      'agent.v1',
      'design-document.v1',
    ])
    try {
      const read = session
        .capabilityFetch('design-document.v1', { action: 'read', documentId: 'document_12345678' })
        .catch((error: Error) => error.message)
      const stop = new AbortController()
      const run = session
        .capabilityFetch('agent.v1', { messages: [] }, stop.signal)
        .catch((error: Error) => error.message)
      await flushFrames()
      const frames = socket.sent.map((raw) => JSON.parse(raw))
      expect(
        frames.filter((item) => item.type === 'office.request').map((item) => item.capability_name),
      ).toEqual(['design-document.v1', 'agent.v1'])
      expect(await read).toBe('relay_cancelled')
      expect(session.snapshot().status).toBe('connected')
      stop.abort()
      await run
    } finally {
      session.disconnect()
    }
  })
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
        if (active) await expect(pending).resolves.toBe('relay_session_expired')
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
    expect(
      send.mock.calls
        .map(([value]) => JSON.parse(value))
        .filter((frame) => frame.type !== 'office.leave'),
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
      'design_contract_review_required',
      'design_contract_prototype_required',
      'design_contract_production_incomplete',
      'design_contract_verification_failed',
      'design_contract_visual_review_failed',
      'design_contract_invalid_status',
      'design_contract_review_not_pending',
      'design_contract_acceptance_mismatch',
      'design_contract_screenshot_required',
      'session_expired',
      'office_screenshot_unavailable',
      'transport_stream_budget_exceeded',
    ].entries()) {
      const eventId = `00000000-0000-4000-8000-${String(index + 6).padStart(12, '0')}`
      const upload = session.sendDiagnostic({
        ...diagnostic,
        event_id: eventId,
        error_code: errorCode,
      })
      expect(frame(socket, index + 5).error_code).toBe(
        errorCode === 'image_fetch_unavailable'
          ? 'network_error'
          : errorCode === 'session_expired'
            ? 'auth_required'
            : errorCode === 'office_screenshot_unavailable'
              ? 'office_read_failed'
              : 'agent_run_failed',
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
    expect(socket.sent).toHaveLength(21)
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

describe('durable PPT regression coverage', () => {
  it('reports an old PC attempting to claim the pending v2 pairing immediately', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1', 'presentation.v1'],
    })
    const connecting = session.connect('powerpoint')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_1',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    expect(session.snapshot().status).toBe('pending')
    socket.receive(
      JSON.stringify({ version: 2, type: 'office.pc_incompatible', pairing_id: 'pair_1' }),
    )
    await connecting
    expect(session.snapshot()).toEqual({ status: 'pc_incompatible' })
  })

  it('classifies only the exact legacy invalid_frame during a v2 handshake as incompatible', async () => {
    for (const legacyFrame of [
      { version: 1, type: 'relay.error', code: 'invalid_frame' },
      { version: 1, type: 'relay.error', code: 'invalid_frame', extra: true },
      { version: 1, type: 'relay.error', code: 'other' },
    ]) {
      const socket = new FakeSocket()
      const session = createOfficeRelaySession({
        createSocket: () => socket,
        capabilities: ['agent.v1'],
      })
      const connecting = session.connect('word')
      socket.open()
      expect(frame(socket, 0).version).toBe(2)
      socket.receive(JSON.stringify(legacyFrame))
      await connecting
      expect(session.snapshot().status).toBe(
        Object.keys(legacyFrame).length === 3 && legacyFrame.code === 'invalid_frame'
          ? 'incompatible'
          : 'offline',
      )
      expect(socket.sent).toHaveLength(1)
    }
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1'],
    })
    const connecting = session.connect('word')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_1',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    expect(session.snapshot().status).toBe('pending')
    socket.receive(JSON.stringify({ version: 1, type: 'relay.error', code: 'invalid_frame' }))
    await connecting
    expect(session.snapshot().status).toBe('offline')
  })

  it.each([
    'presentation.v1',
    'presentation-master-backups.v1',
    'presentation-package-backups.v1',
  ] as const)(
    'reattaches a v2 PowerPoint %s session after socket loss and leaves it on explicit disconnect',
    async (capabilityName) => {
      const first = new FakeSocket(),
        resumed = new FakeSocket()
      const sockets = [first, resumed]
      const session = createOfficeRelaySession({
        createSocket: () => sockets.shift()!,
        capabilities: ['agent.v1', capabilityName],
        randomUUID: () => 'request_after_resume',
      })
      const paired = session.connect('powerpoint')
      first.open()
      first.receive(
        JSON.stringify({
          version: 2,
          type: 'office.created',
          pairing_id: 'pair_1',
          verification_code: '123456',
          expires_in: 120,
        }),
      )
      first.receive(
        JSON.stringify({
          version: 2,
          type: 'office.approved',
          session_id: 'session_1',
          capability: 'cap_1',
          expires_in: 1800,
          capabilities: ['agent.v1', capabilityName],
        }),
      )
      await paired
      expect(session.diagnosticSessionId()).toBe('session_1')
      first.close()
      expect(session.snapshot().status).toBe('offline')
      expect(session.diagnosticSessionId()).toBe('session_1')
      const reconnecting = session.connect('powerpoint')
      resumed.open()
      expect(frame(resumed, 0)).toEqual({
        version: 2,
        type: 'office.resume',
        session_id: 'session_1',
        capability: 'cap_1',
        host: 'PowerPoint',
      })
      resumed.receive(
        JSON.stringify({
          version: 2,
          type: 'office.resumed',
          session_id: 'session_1',
          expires_in: 120,
          pc_online: true,
          capabilities: ['agent.v1', capabilityName],
        }),
      )
      await reconnecting
      expect(session.diagnosticSessionId()).toBe('session_1')
      expect(session.snapshot()).toEqual({
        status: 'connected',
        capabilities: ['agent.v1', capabilityName],
      })
      resumed.receive(JSON.stringify({ version: 2, type: 'office.pc_offline' }))
      expect(session.snapshot().status).toBe('waiting_for_pc')
      resumed.receive(JSON.stringify({ version: 2, type: 'office.pc_online' }))
      expect(session.snapshot().status).toBe('connected')
      const pending = session.capabilityFetch(capabilityName, {
        operation:
          capabilityName === 'presentation-package-backups.v1'
            ? 'package_backup_status'
            : capabilityName === 'presentation-master-backups.v1'
              ? 'master_backup_status'
              : 'status',
      })
      expect(frame(resumed, 1)).toMatchObject({
        type: 'office.request',
        session_id: 'session_1',
        capability: 'cap_1',
        request_id: 'request_after_resume',
      })
      resumed.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.start',
          session_id: 'session_1',
          request_id: 'request_after_resume',
          status: 200,
          content_type: 'application/json',
        }),
      )
      resumed.receive(
        JSON.stringify({
          version: 2,
          type: 'relay.done',
          session_id: 'session_1',
          request_id: 'request_after_resume',
        }),
      )
      expect(await (await pending).text()).toBe('')
      session.disconnect()
      expect(session.diagnosticSessionId()).toBeUndefined()
      expect(frame(resumed, 2)).toEqual({
        version: 2,
        type: 'office.leave',
        session_id: 'session_1',
        capability: 'cap_1',
      })
    },
  )

  it('aborts an active request when the paired PC explicitly revokes the session', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1'],
      randomUUID: () => 'request_revoked',
    })
    const paired = session.connect('powerpoint')
    socket.open()
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_1',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_1',
        capability: 'cap_1',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await paired
    const pending = session.capabilityFetch('agent.v1', { messages: [] })
    expect(frame(socket, 1).request_id).toBe('request_revoked')
    socket.receive(JSON.stringify({ version: 2, type: 'relay.error', code: 'session_revoked' }))
    await expect(pending).rejects.toThrow()
    expect(session.snapshot().status).toBe('offline')
    expect(socket.readyState).toBe(3)
  })

  it('falls back to a new pairing if the Relay cannot resume the previous session', async () => {
    const first = new FakeSocket(),
      resume = new FakeSocket(),
      fallback = new FakeSocket()
    const sockets = [first, resume, fallback]
    const session = createOfficeRelaySession({
      createSocket: () => sockets.shift()!,
      capabilities: ['agent.v1'],
    })
    const paired = session.connect('powerpoint')
    first.open()
    first.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_1',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    first.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_1',
        capability: 'cap_1',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await paired
    first.close()
    const reconnecting = session.connect('powerpoint')
    resume.open()
    resume.receive(JSON.stringify({ version: 1, type: 'relay.error', code: 'invalid_frame' }))
    expect(session.diagnosticSessionId()).toBeUndefined()
    fallback.open()
    expect(frame(fallback, 0)).toEqual({
      version: 2,
      type: 'office.create',
      host: 'PowerPoint',
      capabilities: ['agent.v1'],
    })
    fallback.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_2',
        verification_code: '654321',
        expires_in: 120,
      }),
    )
    expect(session.snapshot()).toEqual({ status: 'pending', verificationCode: '654321' })
    fallback.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_2',
        capability: 'cap_2',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await reconnecting
    expect(session.snapshot().status).toBe('connected')
    expect(session.diagnosticSessionId()).toBe('session_2')
    session.disconnect()
  })

  it('retains resumable credentials after a connected socket error', async () => {
    const first = new FakeSocket(),
      second = new FakeSocket()
    const sockets = [first, second]
    const session = createOfficeRelaySession({
      createSocket: () => sockets.shift()!,
      capabilities: ['agent.v1'],
    })
    const paired = session.connect('powerpoint')
    first.open()
    first.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_1',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    first.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_1',
        capability: 'cap_1',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await paired
    first.onerror?.()
    expect(session.snapshot().status).toBe('offline')
    const reconnecting = session.connect('powerpoint')
    second.open()
    expect(frame(second, 0).type).toBe('office.resume')
    second.receive(
      JSON.stringify({
        version: 2,
        type: 'office.resumed',
        session_id: 'session_1',
        expires_in: 120,
        capabilities: ['agent.v1'],
      }),
    )
    await reconnecting
    session.disconnect()
  })

  it('retries when the Relay still sees the old socket during resume', async () => {
    const first = new FakeSocket(),
      early = new FakeSocket(),
      retry = new FakeSocket()
    const sockets = [first, early, retry]
    const session = createOfficeRelaySession({
      createSocket: () => sockets.shift()!,
      capabilities: ['agent.v1'],
    })
    const paired = session.connect('powerpoint')
    first.open()
    first.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair_1',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    first.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session_1',
        capability: 'cap_1',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await paired
    first.close()
    const reconnecting = session.connect('powerpoint')
    early.open()
    early.receive(JSON.stringify({ version: 2, type: 'relay.error', code: 'session_active' }))
    await new Promise((resolve) => setTimeout(resolve, 130))
    retry.open()
    expect(frame(retry, 0).type).toBe('office.resume')
    retry.receive(
      JSON.stringify({
        version: 2,
        type: 'office.resumed',
        session_id: 'session_1',
        expires_in: 120,
        pc_online: true,
        capabilities: ['agent.v1'],
      }),
    )
    await reconnecting
    expect(session.snapshot().status).toBe('connected')
    session.disconnect()
  })

  it('ignores old socket frames after re-pairing and completes the current request', async () => {
    const oldSocket = new FakeSocket()
    const socket = new FakeSocket()
    const sockets = [oldSocket, socket]
    const session = createOfficeRelaySession({
      createSocket: () => sockets.shift()!,
      randomUUID: () => 'request_current',
    })
    for (const current of [oldSocket, socket]) {
      const connected = session.connect('powerpoint')
      current.open()
      current.receive(
        JSON.stringify({
          version: 1,
          type: 'office.created',
          pairing_id: 'pair_current',
          verification_code: '123456',
          expires_in: 120,
        }),
      )
      current.receive(
        JSON.stringify({
          version: 1,
          type: 'office.approved',
          session_id: 'session_current',
          capability: 'cap_current',
          expires_in: 1800,
        }),
      )
      await connected
    }
    oldSocket.receive('malformed late frame')
    expect(session.snapshot().status).toBe('connected')
    const response = session.authenticatedFetch('/v1/office/messages', {
      method: 'POST',
      body: '{}',
    })
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'relay.start',
        session_id: 'session_current',
        request_id: 'request_current',
        status: 200,
        content_type: 'application/json',
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'relay.chunk',
        session_id: 'session_current',
        request_id: 'request_current',
        sequence: 0,
        data: 'e30=',
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 1,
        type: 'relay.done',
        session_id: 'session_current',
        request_id: 'request_current',
      }),
    )
    expect(await (await response).json()).toEqual({})
    session.disconnect()
  })

  it.each([
    'web-search.v1',
    'presentation.v1',
    'presentation-attachments.v1',
    'presentation-assets.v1',
    'presentation-pdf.v1',
    'presentation-production-pdf.v1',
  ] as const)('negotiates %s and blocks unnegotiated requests', async (capability) => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1', capability, 'web-fetch.v1'],
      randomUUID: () => 'request_12345678',
    })
    const connecting = session.connect('word')
    socket.open()
    expect(frame(socket, 0)).toEqual({
      version: 2,
      type: 'office.create',
      host: 'Word',
      capabilities: ['agent.v1', capability, 'web-fetch.v1'],
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
        capabilities: ['agent.v1', capability],
      }),
    )
    await connecting
    expect(session.snapshot()).toEqual({
      status: 'connected',
      capabilities: ['agent.v1', capability],
    })
    await expect(
      session.capabilityFetch('web-fetch.v1', { url: 'https://example.com' }),
    ).rejects.toThrow('relay_capability_unavailable')
    const pending = session.capabilityFetch(capability, { query: 'office', max_results: 3 })
    expect(frame(socket, 1)).toEqual({
      version: 2,
      type: 'office.request',
      session_id: 'session_12345678',
      capability: 'capability_12345678',
      request_id: 'request_12345678',
      capability_name: capability,
      body: { query: 'office', max_results: 3 },
    })
    session.disconnect()
    await expect(pending).rejects.toThrow('relay_disconnected')
  })

  it('keeps Agent usable when an older PC offers no presentation capability', async () => {
    const socket = new FakeSocket()
    const session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['agent.v1', 'presentation.v1', 'presentation-assets.v1'],
      randomUUID: () => 'request_12345678',
    })
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
    expect(session.snapshot()).toEqual({ status: 'connected', capabilities: ['agent.v1'] })
    await expect(
      session.capabilityFetch('presentation.v1', { operation: 'status' }),
    ).rejects.toThrow('relay_capability_unavailable')
    expect(socket.sent).toHaveLength(1)
    const pending = session.capabilityFetch('agent.v1', { messages: [] })
    expect(frame(socket, 1).capability_name).toBe('agent.v1')
    session.disconnect()
    await expect(pending).rejects.toThrow('relay_disconnected')
  })

  it('sends a private actor token only for a negotiated PowerPoint team request', async () => {
    const f = await teamOfficeSession(async () => 'private-actor-token')
    const request = f.session.capabilityFetch('presentation-team.v1', {
      operation: 'team_identity',
    })
    await vi.waitFor(() =>
      expect(f.socket.sent.some((x) => JSON.parse(x).type === 'office.request')).toBe(true),
    )
    expect(
      f.socket.sent.map((x) => JSON.parse(x)).find((x) => x.type === 'office.request'),
    ).toMatchObject({
      capability_name: 'presentation-team.v1',
      access_token: 'private-actor-token',
      body: { operation: 'team_identity' },
    })
    f.session.disconnect()
    await expect(request).rejects.toThrow()
  })

  it('does not send an anonymous team request or a stale late-token request', async () => {
    const absent = await teamOfficeSession(async () => null)
    await expect(
      absent.session.capabilityFetch('presentation-team.v1', { operation: 'team_identity' }),
    ).rejects.toThrow('relay_team_auth_unavailable')
    expect(absent.socket.sent.some((x) => JSON.parse(x).type === 'office.request')).toBe(false)
    absent.session.disconnect()
    let complete: (token: string) => void = () => {}
    const late = await teamOfficeSession(
      () =>
        new Promise((resolve) => {
          complete = resolve
        }),
    )
    const request = late.session.capabilityFetch('presentation-team.v1', {
      operation: 'team_identity',
    })
    late.session.disconnect()
    complete('private-token')
    await expect(request).rejects.toThrow()
    expect(late.socket.sent.some((x) => JSON.parse(x).type === 'office.request')).toBe(false)
  })

  it('cancels an unresolved team token without sending and rejects caller credentials', async () => {
    let complete: (value: string) => void = () => {}
    const provider = vi.fn(
      () =>
        new Promise<string>((resolve) => {
          complete = resolve
        }),
    )
    const f = await teamOfficeSession(provider)
    const controller = new AbortController()
    const request = f.session.capabilityFetch(
      'presentation-team.v1',
      { operation: 'team_identity' },
      controller.signal,
    )
    await vi.waitFor(() => expect(provider).toHaveBeenCalledOnce())
    controller.abort()
    await expect(request).rejects.toThrow()
    complete('late-private-token')
    await expect(
      f.session.capabilityFetch('agent.v1', { access_token: 'forged' }),
    ).rejects.toThrow()
    await expect(
      f.session.capabilityFetch('presentation-team.v1', { team_context: {} }),
    ).rejects.toThrow()
    expect(f.socket.sent.some((value) => JSON.parse(value).type === 'office.request')).toBe(false)
    f.session.disconnect()
  })

  it('refuses a team capability combined with any private workspace capability before pairing', async () => {
    const createSocket = vi.fn(() => new FakeSocket())
    for (const other of ['agent.v1', 'presentation.v1', 'presentation-attachments.v1'] as const) {
      const session = createOfficeRelaySession({
        createSocket,
        capabilities: ['presentation-team.v1', other],
        getTeamAccessToken: async () => 'token',
      })
      await expect(session.connect('powerpoint')).rejects.toThrow('relay_invalid_capabilities')
      expect(session.snapshot().status).toBe('offline')
    }
    expect(createSocket).not.toHaveBeenCalled()
  })

  it('negotiates master backups and preserves bounded streamed response and missing-cap failure', async () => {
    const socket = new FakeSocket(),
      session = createOfficeRelaySession({
        createSocket: () => socket,
        capabilities: ['agent.v1', 'presentation-master-backups.v1'],
      })
    const pending = session.connect('powerpoint')
    socket.open()
    expect(frame(socket, 0).capabilities).toEqual(['agent.v1', 'presentation-master-backups.v1'])
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session',
        capability: 'secret',
        expires_in: 1800,
        capabilities: ['agent.v1', 'presentation-master-backups.v1'],
      }),
    )
    await pending
    const responsePending = session.capabilityFetch('presentation-master-backups.v1', {
      operation: 'master_backup_status',
      documentId: 'doc',
      changeId: 'change',
      key: 'snapshot',
    })
    const request = frame(socket, 1)
    expect(request).toMatchObject({
      capability_name: 'presentation-master-backups.v1',
      body: { operation: 'master_backup_status' },
    })
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.start',
        session_id: 'session',
        request_id: request.request_id,
        status: 200,
        content_type: 'application/json',
      }),
    )
    const response = await responsePending
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.chunk',
        session_id: 'session',
        request_id: request.request_id,
        sequence: 0,
        data: btoa('{}'),
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.done',
        session_id: 'session',
        request_id: request.request_id,
      }),
    )
    expect(await response.text()).toBe('{}')
    session.disconnect()
    const oldSocket = new FakeSocket(),
      old = createOfficeRelaySession({
        createSocket: () => oldSocket,
        capabilities: ['agent.v1', 'presentation-master-backups.v1'],
      })
    const connected = old.connect('powerpoint')
    oldSocket.open()
    oldSocket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    oldSocket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session',
        capability: 'secret',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connected
    await expect(
      old.capabilityFetch('presentation-master-backups.v1', { operation: 'master_backup_status' }),
    ).rejects.toThrow('relay_capability_unavailable')
    expect(oldSocket.sent).toHaveLength(1)
    old.disconnect()
  })

  it('filters master backup capability from non-PowerPoint handshakes', async () => {
    const socket = new FakeSocket(),
      session = createOfficeRelaySession({
        createSocket: () => socket,
        capabilities: ['agent.v1', 'presentation-master-backups.v1'],
      })
    const pending = session.connect('word')
    socket.open()
    expect(frame(socket, 0).capabilities).toEqual(['agent.v1'])
    session.disconnect()
    await pending
  })

  it('negotiates package backups and preserves bounded streamed response and missing-cap failure', async () => {
    const socket = new FakeSocket(),
      session = createOfficeRelaySession({
        createSocket: () => socket,
        capabilities: ['agent.v1', 'presentation-package-backups.v1'],
      })
    const pending = session.connect('powerpoint')
    socket.open()
    expect(frame(socket, 0).capabilities).toEqual(['agent.v1', 'presentation-package-backups.v1'])
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session',
        capability: 'secret',
        expires_in: 1800,
        capabilities: ['agent.v1', 'presentation-package-backups.v1'],
      }),
    )
    await pending
    const responsePending = session.capabilityFetch('presentation-package-backups.v1', {
      operation: 'package_backup_status',
      documentId: 'doc',
      changeId: 'change',
      key: 'snapshot',
    })
    const request = frame(socket, 1)
    expect(request).toMatchObject({
      capability_name: 'presentation-package-backups.v1',
      body: { operation: 'package_backup_status' },
    })
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.start',
        session_id: 'session',
        request_id: request.request_id,
        status: 200,
        content_type: 'application/json',
      }),
    )
    const response = await responsePending
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.chunk',
        session_id: 'session',
        request_id: request.request_id,
        sequence: 0,
        data: btoa('{}'),
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.done',
        session_id: 'session',
        request_id: request.request_id,
      }),
    )
    expect(await response.text()).toBe('{}')
    session.disconnect()
    const oldSocket = new FakeSocket(),
      old = createOfficeRelaySession({
        createSocket: () => oldSocket,
        capabilities: ['agent.v1', 'presentation-package-backups.v1'],
      })
    const connected = old.connect('powerpoint')
    oldSocket.open()
    oldSocket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    oldSocket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session',
        capability: 'secret',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connected
    await expect(
      old.capabilityFetch('presentation-package-backups.v1', {
        operation: 'package_backup_status',
      }),
    ).rejects.toThrow('relay_capability_unavailable')
    expect(oldSocket.sent).toHaveLength(1)
    old.disconnect()
  })

  it('filters package backup capability from non-PowerPoint handshakes', async () => {
    const socket = new FakeSocket(),
      session = createOfficeRelaySession({
        createSocket: () => socket,
        capabilities: ['agent.v1', 'presentation-package-backups.v1'],
      })
    const pending = session.connect('word')
    socket.open()
    expect(frame(socket, 0).capabilities).toEqual(['agent.v1'])
    session.disconnect()
    await pending
  })

  it('filters governance from non-PowerPoint handshakes', async () => {
    const socket = new FakeSocket(),
      session = createOfficeRelaySession({
        createSocket: () => socket,
        capabilities: ['agent.v1', 'presentation-governance.v1'],
      })
    const pending = session.connect('word')
    socket.open()
    expect(frame(socket, 0).capabilities).toEqual(['agent.v1'])
    session.disconnect()
    await pending
  })

  it('negotiates governance and preserves bounded streamed response and missing-cap failure', async () => {
    const socket = new FakeSocket(),
      session = createOfficeRelaySession({
        createSocket: () => socket,
        capabilities: ['agent.v1', 'presentation-governance.v1'],
      })
    const pending = session.connect('powerpoint')
    socket.open()
    expect(frame(socket, 0).capabilities).toEqual(['agent.v1', 'presentation-governance.v1'])
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session',
        capability: 'secret',
        expires_in: 1800,
        capabilities: ['agent.v1', 'presentation-governance.v1'],
      }),
    )
    await pending
    for (const operation of ['unknown', 'package_backup_status', 'team_project_read']) {
      await expect(
        session.capabilityFetch('presentation-governance.v1', { operation }),
      ).rejects.toThrow('relay_invalid_request')
    }
    await expect(
      session.capabilityFetch('agent.v1', { operation: 'project_deletion_confirm' }),
    ).rejects.toThrow('relay_invalid_request')
    await expect(
      session.capabilityFetch('presentation-governance.v1', {
        operation: 'project_lifecycle_read',
        team_context: {},
      }),
    ).rejects.toThrow('relay_invalid_request')
    await expect(
      session.capabilityFetch('presentation-governance.v1', {
        operation: 'project_lifecycle_read',
        access_token: 'private',
      }),
    ).rejects.toThrow('relay_invalid_request')
    const responsePending = session.capabilityFetch('presentation-governance.v1', {
      operation: 'project_lifecycle_read',
      documentId: 'doc',
      changeId: 'change',
      key: 'snapshot',
    })
    const request = frame(socket, 1)
    expect(request).toMatchObject({
      capability_name: 'presentation-governance.v1',
      body: { operation: 'project_lifecycle_read' },
    })
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.start',
        session_id: 'session',
        request_id: request.request_id,
        status: 200,
        content_type: 'application/json',
      }),
    )
    const response = await responsePending
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.chunk',
        session_id: 'session',
        request_id: request.request_id,
        sequence: 0,
        data: btoa('{}'),
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.done',
        session_id: 'session',
        request_id: request.request_id,
      }),
    )
    expect(await response.text()).toBe('{}')
    session.disconnect()
    const oldSocket = new FakeSocket(),
      old = createOfficeRelaySession({
        createSocket: () => oldSocket,
        capabilities: ['agent.v1', 'presentation-governance.v1'],
      })
    const connected = old.connect('powerpoint')
    oldSocket.open()
    oldSocket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    oldSocket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session',
        capability: 'secret',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connected
    await expect(
      old.capabilityFetch('presentation-governance.v1', { operation: 'project_lifecycle_read' }),
    ).rejects.toThrow('relay_capability_unavailable')
    expect(oldSocket.sent).toHaveLength(1)
    old.disconnect()
  })

  it('requests the complete sixteen-capability primary set without mixing TeamOnly', async () => {
    const capabilities = [
      'agent.v1',
      'web-search.v1',
      'web-fetch.v1',
      'image-search.v1',
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
      'presentation-governance.v1',
    ] as const
    const socket = new FakeSocket(),
      session = createOfficeRelaySession({ createSocket: () => socket, capabilities })
    const pending = session.connect('powerpoint')
    socket.open()
    expect(frame(socket, 0).capabilities).toEqual(capabilities)
    expect(frame(socket, 0).capabilities).toHaveLength(16)
    session.disconnect()
    await pending
  })

  it('preserves existing presentation lifecycle compatibility and preserves bounded streamed response and missing-cap failure', async () => {
    const socket = new FakeSocket(),
      session = createOfficeRelaySession({
        createSocket: () => socket,
        capabilities: ['agent.v1', 'presentation.v1'],
      })
    const pending = session.connect('powerpoint')
    socket.open()
    expect(frame(socket, 0).capabilities).toEqual(['agent.v1', 'presentation.v1'])
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session',
        capability: 'secret',
        expires_in: 1800,
        capabilities: ['agent.v1', 'presentation.v1'],
      }),
    )
    await pending
    const responsePending = session.capabilityFetch('presentation.v1', {
      operation: 'project_lifecycle_read',
      documentId: 'doc',
      changeId: 'change',
      key: 'snapshot',
    })
    const request = frame(socket, 1)
    expect(request).toMatchObject({
      capability_name: 'presentation.v1',
      body: { operation: 'project_lifecycle_read' },
    })
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.start',
        session_id: 'session',
        request_id: request.request_id,
        status: 200,
        content_type: 'application/json',
      }),
    )
    const response = await responsePending
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.chunk',
        session_id: 'session',
        request_id: request.request_id,
        sequence: 0,
        data: btoa('{}'),
      }),
    )
    socket.receive(
      JSON.stringify({
        version: 2,
        type: 'relay.done',
        session_id: 'session',
        request_id: request.request_id,
      }),
    )
    expect(await response.text()).toBe('{}')
    session.disconnect()
    const oldSocket = new FakeSocket(),
      old = createOfficeRelaySession({
        createSocket: () => oldSocket,
        capabilities: ['agent.v1', 'presentation.v1'],
      })
    const connected = old.connect('powerpoint')
    oldSocket.open()
    oldSocket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.created',
        pairing_id: 'pair',
        verification_code: '123456',
        expires_in: 120,
      }),
    )
    oldSocket.receive(
      JSON.stringify({
        version: 2,
        type: 'office.approved',
        session_id: 'session',
        capability: 'secret',
        expires_in: 1800,
        capabilities: ['agent.v1'],
      }),
    )
    await connected
    await expect(
      old.capabilityFetch('presentation.v1', { operation: 'project_lifecycle_read' }),
    ).rejects.toThrow('relay_capability_unavailable')
    expect(oldSocket.sent).toHaveLength(1)
    old.disconnect()
  })
})

async function teamOfficeSession(getTeamAccessToken: () => Promise<string | null>) {
  const socket = new FakeSocket(),
    session = createOfficeRelaySession({
      createSocket: () => socket,
      capabilities: ['presentation-team.v1'],
      getTeamAccessToken,
      randomUUID: () => 'team_request',
    })
  const connecting = session.connect('powerpoint')
  socket.open()
  socket.receive(
    JSON.stringify({
      version: 2,
      type: 'office.created',
      pairing_id: 'pair_team',
      verification_code: '123456',
      expires_in: 120,
    }),
  )
  socket.receive(
    JSON.stringify({
      version: 2,
      type: 'office.approved',
      session_id: 'session_team',
      capability: 'cap_team',
      expires_in: 1800,
      capabilities: ['presentation-team.v1'],
    }),
  )
  await connecting
  return { session, socket }
}

it('keeps team pairing separate from the saved document binding', async () => {
  const socket = new FakeSocket()
  const load = vi.fn()
  const session = createOfficeRelaySession({
    createSocket: () => socket,
    capabilities: ['presentation-team.v1'],
    getTeamAccessToken: async () => 'private-actor-token',
    persistentPairing: true,
    bindingStore: { load } as never,
  })
  const connected = session.connect('powerpoint')
  socket.open()
  expect(frame(socket, 0)).toMatchObject({
    type: 'office.create',
    capabilities: ['presentation-team.v1'],
  })
  expect(frame(socket, 0)).not.toHaveProperty('features')
  expect(load).not.toHaveBeenCalled()
  session.disconnect()
  await connected
})
it('filters PowerPoint-only capabilities before persistent Word enrollment', async () => {
  const socket = new FakeSocket()
  const load = vi.fn(async () => undefined)
  const createEnrollment = vi.fn(async () => {
    throw new Error('key unavailable')
  })
  const session = createOfficeRelaySession({
    createSocket: () => socket,
    capabilities: ['agent.v1', 'presentation-governance.v1', 'presentation-master-backups.v1'],
    persistentPairing: true,
    bindingStore: { load, createEnrollment, forget: vi.fn(async () => undefined) } as never,
  })
  const connected = session.connect('word')
  await vi.waitFor(() => expect(createEnrollment).toHaveBeenCalledWith('word', ['agent.v1']))
  expect(load).toHaveBeenCalledWith('word', ['agent.v1'])
  socket.open()
  expect(frame(socket, 0).capabilities).toEqual(['agent.v1'])
  session.disconnect()
  await connected
})
