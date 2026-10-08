import { afterEach, expect, it, vi } from 'vitest'
import { createOfficeRelayClient, type RelaySocket } from '../src/main/office-relay-client'
import { ENHANCED_HOSTS } from '@wiswork/agent-runtime'
import { createOfficeCodexProxy } from '../src/main/office-codex-proxy'
import { createShellEnhancedPolicyAuthority } from '../src/main/enhanced-policy-authority'

class Socket implements RelaySocket {
  readyState = 0
  sent: Record<string, any>[] = []
  listeners = new Map<string, Array<(event: any) => void>>()
  addEventListener(name: string, callback: (event: any) => void) {
    this.listeners.set(name, [...(this.listeners.get(name) ?? []), callback])
  }
  send(data: string) {
    this.sent.push(JSON.parse(data))
  }
  close() {
    this.readyState = 3
    this.emit('close', {})
  }
  emit(name: string, event: any) {
    for (const callback of this.listeners.get(name) ?? []) callback(event)
  }
  message(frame: object) {
    this.emit('message', { data: JSON.stringify(frame) })
  }
}

const cleanups: Array<() => void> = []
afterEach(() => {
  for (const cleanup of cleanups.splice(0)) cleanup()
  vi.useRealTimers()
})

async function fixture(runOfficeTurn?: (input: any) => Promise<void>) {
  const socket = new Socket()
  const turns: any[] = []
  const proxy = runOfficeTurn
    ? createOfficeCodexProxy({
        runtime: { runOfficeTurn } as any,
        rollout: {
          globalEnabled: true,
          rawOfficeEnabled: false,
          hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as any,
        },
        policyAuthority: createShellEnhancedPolicyAuthority(() => 2),
      })
    : undefined
  const client = createOfficeRelayClient({
    endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
    connect: () => socket,
    getValidAccountStatus: async () => ({ loggedIn: true }),
    getAccessToken: async () => 'test-token',
    proxy: async () => ({ status: 200, body: new Uint8Array() }),
    enhancedProxy: async (request) => {
      let finish!: () => void
      const finished = new Promise<void>((resolve) => (finish = resolve))
      turns.push({ ...request, finish })
      if (proxy) return proxy(request)
      return {
        status: 200,
        body: (async function* () {
          await finished
          yield new Uint8Array()
        })(),
      }
    },
    enhancedStatement: () => ({
      version: 1,
      runtime_mode: 'enhanced',
      runtime_instance: 'runtime_12345678',
      component_version: '0.147.0',
      host: 'office-powerpoint',
      raw_office: false,
      expires_at: Date.now() + 60_000,
      policy_generation: 2,
      session_generation: 7,
    }),
    negotiateCapabilities: true,
    onPending() {},
  })
  cleanups.push(() => {
    client.revoke('test-complete')
    for (const turn of turns) turn.finish()
  })
  const claiming = client.claim('123456')
  await vi.waitFor(() => expect(socket.listeners.has('open')).toBe(true))
  socket.readyState = 1
  socket.emit('open', {})
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
    host: 'PowerPoint',
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
    capability: 'test-capability',
    expires_in: 1800,
    capabilities: ['agent.v1'],
  })
  const start = async (requestId = 'request_12345678') => {
    const count = turns.length
    socket.message({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: requestId,
      capability_name: 'agent.v1',
      body: {
        system: '',
        messages: [],
        tools: [
          {
            name: 'list_slide_shapes',
            description: 'Read slide shapes',
            input_schema: { type: 'object' },
          },
        ],
      },
    })
    await vi.waitFor(() => expect(turns).toHaveLength(count + 1))
    return turns.at(-1)!
  }
  const call = (id: string) => ({
    turnId: 'turn_12345678',
    callId: id,
    generation: 7,
    toolName: 'list_slide_shapes',
    input: { slide_index: 0 },
  })
  const reply = (id: string, requestId = 'request_12345678', turnId = 'turn_12345678') =>
    socket.message({
      version: 2,
      type: 'relay.tool_result',
      session_id: 'session_12345678',
      request_id: requestId,
      turn_id: turnId,
      call_id: id,
      generation: 7,
      output: id,
      is_error: false,
    })
  return {
    socket,
    client,
    start,
    call,
    reply,
    calls: () => socket.sent.filter((f) => f.type === 'pc.tool_call'),
  }
}

it('serializes concurrent remote reads and preserves each result identity', async () => {
  const f = await fixture(),
    turn = await f.start()
  const first = turn.executeTool(f.call('call_first'))
  const second = turn.executeTool(f.call('call_second'))
  const results = Promise.all([first, second])
  expect(f.calls().map((f) => f.call_id)).toEqual(['call_first'])
  f.reply('call_first')
  await vi.waitFor(() =>
    expect(f.calls().map((f) => f.call_id)).toEqual(['call_first', 'call_second']),
  )
  f.reply('call_second')
  await expect(results).resolves.toEqual([
    { output: 'call_first', isError: false },
    { output: 'call_second', isError: false },
  ])
})

it('cancels a queued read without interrupting the in-flight read or dispatching the cancelled call', async () => {
  const f = await fixture(),
    turn = await f.start(),
    abort = new AbortController()
  const first = turn.executeTool(f.call('call_first'))
  const second = turn
    .executeTool(f.call('call_second'), abort.signal)
    .catch((e: Error) => e.message)
  abort.abort()
  await expect(second).resolves.toBe('tool_cancelled')
  expect(turn.signal.aborted).toBe(false)
  f.reply('call_first')
  await expect(first).resolves.toMatchObject({ output: 'call_first' })
  expect(f.calls()).toHaveLength(1)
})

it('bounds queued calls instead of retaining unbounded input', async () => {
  const f = await fixture(),
    turn = await f.start()
  const calls = Array.from({ length: 8 }, (_, index) =>
    turn.executeTool(f.call(`call_number_${index}`)).catch((e: Error) => e.message),
  )
  await expect(turn.executeTool(f.call('call_overflow'))).rejects.toThrow('tool_queue_full')
  f.client.revoke()
  expect(await Promise.all(calls)).toEqual(Array(8).fill('tool_cancelled'))
})

it('ends an aborted in-flight request, ignores its known late result and allows a fresh turn', async () => {
  const f = await fixture(),
    turn = await f.start(),
    abort = new AbortController()
  const first = turn
    .executeTool({ ...f.call('call_first'), toolName: 'edit_slide_text' }, abort.signal)
    .catch((e: Error) => e.message)
  const queued = turn.executeTool(f.call('call_second')).catch((e: Error) => e.message)
  abort.abort()
  await expect(first).resolves.toBe('tool_cancelled')
  await expect(queued).resolves.toBe('tool_cancelled')
  expect(turn.signal.aborted).toBe(true)
  expect(f.socket.sent).toContainEqual(
    expect.objectContaining({ type: 'pc.error', code: 'cancelled' }),
  )
  const next = await f.start('request_replacement')
  const read = next.executeTool(f.call('call_next'))
  f.reply('call_first')
  expect(f.client.status()).toBe('paired')
  f.reply('call_next', 'request_replacement')
  await expect(read).resolves.toMatchObject({ output: 'call_next' })
  turn.finish()
  await new Promise(setImmediate)
  expect(next.signal.aborted).toBe(false)
  expect(f.calls().map((f) => f.call_id)).toEqual(['call_first', 'call_next'])
})

it('does not finish a request successfully while a dispatched write remains unresolved', async () => {
  const f = await fixture(),
    turn = await f.start()
  const write = turn
    .executeTool({ ...f.call('call_first'), toolName: 'edit_slide_text' })
    .catch((e: Error) => e.message)
  turn.finish()
  await expect(write).resolves.toBe('tool_cancelled')
  expect(f.socket.sent.some((frame) => frame.type === 'pc.done')).toBe(false)
  expect(f.socket.sent).toContainEqual(
    expect.objectContaining({ type: 'pc.error', code: 'upstream_error' }),
  )
  const next = await f.start('request_replacement')
  const read = next.executeTool(f.call('call_next'))
  f.reply('call_first')
  f.reply('call_next', 'request_replacement')
  await expect(read).resolves.toMatchObject({ output: 'call_next' })
})

it('runs a concurrent model read batch through the real proxy, tool router and Relay queue', async () => {
  let outputs: string[] = []
  const f = await fixture(async ({ toolSession, onEvent }) => {
    const results = await Promise.all(
      [0, 1].map((slide_index) =>
        toolSession.callTool(toolSession.credentials, {
          id: `read_${slide_index}`,
          name: 'list_slide_shapes',
          input: { slide_index },
        }),
      ),
    )
    expect(results.every((result: any) => !result.isError)).toBe(true)
    outputs = results.map((result: any) => result.output)
    onEvent({ type: 'terminal', status: 'completed' })
  })
  await f.start()
  expect(f.calls()).toHaveLength(1)
  const first = f.calls()[0]
  f.reply(first.call_id, first.request_id, first.turn_id)
  await vi.waitFor(() => expect(f.calls()).toHaveLength(2))
  const second = f.calls()[1]
  f.reply(second.call_id, second.request_id, second.turn_id)
  await vi.waitFor(() => expect(f.socket.sent.some((frame) => frame.type === 'pc.done')).toBe(true))
  expect(outputs).toEqual([first.call_id, second.call_id])
})

it('releases a timed-out real router call and keeps the next Relay request usable', async () => {
  vi.useFakeTimers()
  const outputs: any[] = []
  const f = await fixture(async ({ toolSession, onEvent }) => {
    outputs.push(
      await toolSession.callTool(toolSession.credentials, {
        id: 'read_timeout',
        name: 'list_slide_shapes',
        input: { slide_index: 0 },
      }),
    )
    onEvent({ type: 'terminal', status: 'completed' })
  })
  await f.start()
  const first = f.calls()[0]
  await vi.advanceTimersByTimeAsync(30_000)
  expect(outputs).toMatchObject([{ output: 'tool_timeout', isError: true }])
  expect(f.socket.sent).toContainEqual(
    expect.objectContaining({ type: 'pc.error', code: 'cancelled' }),
  )
  await f.start('request_replacement')
  const next = f.calls()[1]
  f.reply(first.call_id, first.request_id, first.turn_id)
  expect(f.client.status()).toBe('paired')
  f.reply(next.call_id, next.request_id, next.turn_id)
  await vi.advanceTimersByTimeAsync(0)
  expect(outputs[1]).toMatchObject({ output: next.call_id, isError: false })
})
