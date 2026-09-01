import { PassThrough } from 'node:stream'
import { describe, expect, it } from 'vitest'
import { CodexAppServerClient } from '../src/app-server-client.js'
import { JsonRpcClient } from '../src/json-rpc.js'

function createClient() {
  const fromServer = new PassThrough()
  const toServer = new PassThrough()
  const writes: Array<Record<string, unknown>> = []
  toServer.setEncoding('utf8')
  toServer.on('data', (line: string) => writes.push(JSON.parse(line) as Record<string, unknown>))
  const rpc = new JsonRpcClient({ input: fromServer, output: toServer, requestTimeoutMs: 100 })
  const client = new CodexAppServerClient({
    rpc,
    cwd: '/isolated/empty',
    developerInstructions: 'Fixed host policy.',
  })
  return { client, fromServer, writes }
}

async function initialize(fixture: ReturnType<typeof createClient>): Promise<void> {
  const pending = fixture.client.initialize()
  expect(fixture.writes[0]).toEqual({
    jsonrpc: '2.0',
    id: 1,
    method: 'initialize',
    params: {
      clientInfo: { name: 'wiswork', version: '0.1.0' },
      capabilities: null,
    },
  })
  fixture.fromServer.write(
    '{"jsonrpc":"2.0","id":1,"result":{"userAgent":"codex-cli/0.147.0","codexHome":"/isolated/home","platformFamily":"unix","platformOs":"linux"}}\n',
  )
  await pending
  expect(fixture.writes[1]).toEqual({
    jsonrpc: '2.0',
    method: 'initialized',
    params: {},
  })
}

describe('Codex app-server client', () => {
  it('rejects oversized fixed policy and turn inputs before writing protocol data', async () => {
    const fixture = createClient()
    const oversizedRpc = new JsonRpcClient({
      input: new PassThrough(),
      output: new PassThrough(),
      requestTimeoutMs: 100,
    })
    let oversizedClient: CodexAppServerClient | undefined
    expect(
      () =>
        (oversizedClient = new CodexAppServerClient({
          rpc: oversizedRpc,
          cwd: '/isolated/empty',
          developerInstructions: 'x'.repeat(65_537),
        })),
    ).toThrow('invalid_app_server_client_options')
    await oversizedClient?.shutdown()
    await oversizedRpc.close()

    await initialize(fixture)
    await expect(fixture.client.startTurn('t'.repeat(257), 'prompt')).rejects.toMatchObject({
      code: 'app_server_invalid_argument',
    })
    await expect(fixture.client.startTurn('thread-1', 'x'.repeat(1_000_001))).rejects.toMatchObject(
      {
        code: 'app_server_invalid_argument',
      },
    )
    expect(fixture.writes).toHaveLength(2)
    await fixture.client.shutdown()
  })

  it('requires the pinned 0.147.0 server identity', async () => {
    const fixture = createClient()
    const pending = fixture.client.initialize()
    fixture.fromServer.write(
      '{"jsonrpc":"2.0","id":1,"result":{"userAgent":"codex-cli/0.148.0","codexHome":"/isolated/home","platformFamily":"unix","platformOs":"linux"}}\n',
    )
    await expect(pending).rejects.toMatchObject({ code: 'app_server_protocol_error' })
  })
  it('initializes exactly once before any lifecycle method', async () => {
    const fixture = createClient()
    await expect(fixture.client.startThread()).rejects.toMatchObject({
      code: 'app_server_not_initialized',
    })

    await initialize(fixture)
    await expect(fixture.client.initialize()).rejects.toMatchObject({
      code: 'app_server_already_initialized',
    })
    await fixture.client.shutdown()
  })

  it('starts an ephemeral fixed-policy thread and validates its id', async () => {
    const fixture = createClient()
    await initialize(fixture)
    const pending = fixture.client.startThread()

    expect(fixture.writes[2]).toEqual({
      jsonrpc: '2.0',
      id: 2,
      method: 'thread/start',
      params: {
        model: 'gpt-5.6-sol',
        modelProvider: 'wiswork',
        cwd: '/isolated/empty',
        approvalPolicy: 'never',
        sandbox: 'read-only',
        developerInstructions: 'Fixed host policy.',
        ephemeral: true,
      },
    })
    fixture.fromServer.write(
      '{"jsonrpc":"2.0","id":2,"result":{"thread":{"id":"thread-1"},"model":"gpt-5.6-sol"}}\n',
    )
    await expect(pending).resolves.toMatchObject({ thread: { id: 'thread-1' } })
    await fixture.client.shutdown()
  })

  it('starts and interrupts turns with the observed 0.147.0 wire shape', async () => {
    const fixture = createClient()
    await initialize(fixture)
    const turn = fixture.client.startTurn('thread-1', 'private prompt')
    expect(fixture.writes[2]).toEqual({
      jsonrpc: '2.0',
      id: 2,
      method: 'turn/start',
      params: {
        threadId: 'thread-1',
        input: [{ type: 'text', text: 'private prompt', text_elements: [] }],
        effort: 'medium',
      },
    })
    fixture.fromServer.write(
      '{"jsonrpc":"2.0","id":2,"result":{"turn":{"id":"turn-1","status":"inProgress"}}}\n',
    )
    await expect(turn).resolves.toMatchObject({ turn: { id: 'turn-1' } })

    const interrupt = fixture.client.interruptTurn('thread-1', 'turn-1')
    expect(fixture.writes[3]).toEqual({
      jsonrpc: '2.0',
      id: 3,
      method: 'turn/interrupt',
      params: { threadId: 'thread-1', turnId: 'turn-1' },
    })
    fixture.fromServer.write('{"jsonrpc":"2.0","id":3,"result":{}}\n')
    await expect(interrupt).resolves.toEqual({})
    await fixture.client.shutdown()
  })

  it('forwards only version-known typed notifications and safely unsubscribes', async () => {
    const fixture = createClient()
    await initialize(fixture)
    const seen: unknown[] = []
    const unsubscribe = fixture.client.onNotification((notification) => seen.push(notification))
    fixture.fromServer.write(
      '{"jsonrpc":"2.0","method":"item/agentMessage/delta","params":{"threadId":"t","turnId":"u","itemId":"i","delta":"hello"}}\n',
    )
    expect(seen).toEqual([
      {
        method: 'item/agentMessage/delta',
        params: { threadId: 't', turnId: 'u', itemId: 'i', delta: 'hello' },
      },
    ])
    unsubscribe()
    fixture.fromServer.write(
      '{"jsonrpc":"2.0","method":"turn/completed","params":{"threadId":"t","turn":{"id":"u"}}}\n',
    )
    expect(seen).toHaveLength(1)
    await fixture.client.shutdown()
  })

  it('fails closed on malformed ids and unknown notification methods', async () => {
    const malformed = createClient()
    await initialize(malformed)
    const pending = malformed.client.startTurn('thread-1', 'private prompt')
    malformed.fromServer.write('{"jsonrpc":"2.0","id":2,"result":{"turn":{"id":""}}}\n')
    await expect(pending).rejects.toMatchObject({ code: 'app_server_protocol_error' })
    await malformed.client.shutdown()

    const unknown = createClient()
    await initialize(unknown)
    unknown.fromServer.write(
      '{"jsonrpc":"2.0","method":"future/private","params":{"prompt":"private prompt"}}\n',
    )
    await expect(unknown.client.startThread()).rejects.toMatchObject({
      code: 'app_server_protocol_error',
    })
  })

  it.each([
    'item/commandExecution/outputDelta',
    'item/fileChange/outputDelta',
    'process/outputDelta',
    'turn/plan/updated',
  ])('fails closed on disabled capability notification %s', async (method) => {
    const fixture = createClient()
    await initialize(fixture)
    fixture.fromServer.write(
      `${JSON.stringify({ jsonrpc: '2.0', method, params: { private: 'secret' } })}\n`,
    )
    await expect(fixture.client.startThread()).rejects.toMatchObject({
      code: 'app_server_protocol_error',
    })
  })

  it('shutdown is idempotent and prevents new requests', async () => {
    const fixture = createClient()
    await initialize(fixture)
    const first = fixture.client.shutdown()
    const second = fixture.client.shutdown()
    expect(second).toBe(first)
    await first
    await expect(fixture.client.startThread()).rejects.toMatchObject({
      code: 'app_server_closed',
    })
  })
})
