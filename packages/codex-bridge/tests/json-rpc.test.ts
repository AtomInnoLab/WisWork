import { PassThrough, Writable } from 'node:stream'
import { describe, expect, it, vi } from 'vitest'
import { JsonRpcClient, JsonRpcError } from '../src/json-rpc.js'

function createTransport(
  options: Omit<Partial<ConstructorParameters<typeof JsonRpcClient>[0]>, 'input' | 'output'> = {},
) {
  const fromServer = new PassThrough()
  const toServer = new PassThrough()
  const writes: Array<Record<string, unknown>> = []
  toServer.setEncoding('utf8')
  toServer.on('data', (line: string) => writes.push(JSON.parse(line) as Record<string, unknown>))
  const client = new JsonRpcClient({
    input: fromServer,
    output: toServer,
    requestTimeoutMs: 1_000,
    ...options,
  })
  return { client, fromServer, toServer, writes }
}

async function tick(): Promise<void> {
  await new Promise((resolve) => setImmediate(resolve))
}

describe('JSONL JSON-RPC client', () => {
  it('correlates out-of-order responses using monotonically increasing ids', async () => {
    const { client, fromServer, writes } = createTransport()
    const first = client.request<{ value: string }>('first', { private: 'prompt' })
    const second = client.request<{ value: string }>('second', {})

    expect(writes).toEqual([
      { jsonrpc: '2.0', id: 1, method: 'first', params: { private: 'prompt' } },
      { jsonrpc: '2.0', id: 2, method: 'second', params: {} },
    ])
    fromServer.write('{"jsonrpc":"2.0","id":2,"result":{"value":"b"}}\n')
    fromServer.write('{"jsonrpc":"2.0","id":1,"result":{"value":"a"}}\n')

    await expect(first).resolves.toEqual({ value: 'a' })
    await expect(second).resolves.toEqual({ value: 'b' })
    await client.close()
  })

  it('emits typed notifications and supports safe unsubscribe', async () => {
    const { client, fromServer } = createTransport()
    const seen: unknown[] = []
    const throwing = client.subscribe(() => {
      throw new Error('must not escape')
    })
    const unsubscribe = client.subscribe((notification) => seen.push(notification))

    fromServer.write('{"jsonrpc":"2.0","method":"turn/completed","params":{"turnId":"t"}}\n')
    await tick()
    expect(seen).toEqual([{ method: 'turn/completed', params: { turnId: 't' } }])
    throwing()
    unsubscribe()
    fromServer.write('{"jsonrpc":"2.0","method":"ignored","params":{}}\n')
    await tick()
    expect(seen).toHaveLength(1)
    await client.close()
  })

  it.each([
    ['malformed JSON', 'not private prompt\n'],
    ['invalid envelope', '{"jsonrpc":"1.0","id":1,"result":{}}\n'],
    ['unknown response', '{"jsonrpc":"2.0","id":99,"result":{}}\n'],
    [
      'ambiguous response',
      '{"jsonrpc":"2.0","id":1,"result":{},"error":{"code":-1,"message":"private prompt"}}\n',
    ],
  ])('fails closed on %s with redacted diagnostics', async (_name, line) => {
    const diagnostic = vi.fn()
    const { client, fromServer } = createTransport({ diagnostics: diagnostic })
    const pending = client.request('test', {})
    fromServer.write(line)

    await expect(pending).rejects.toMatchObject({ code: 'rpc_protocol_error' })
    expect(diagnostic).toHaveBeenCalledWith({ code: 'rpc_protocol_error' })
    expect(JSON.stringify(diagnostic.mock.calls)).not.toContain('private prompt')
  })

  it('treats a duplicate response as an unknown response and closes', async () => {
    const diagnostic = vi.fn()
    const { client, fromServer } = createTransport({ diagnostics: diagnostic })
    const request = client.request('once', {})
    fromServer.write('{"jsonrpc":"2.0","id":1,"result":{}}\n')
    await expect(request).resolves.toEqual({})
    fromServer.write('{"jsonrpc":"2.0","id":1,"result":{}}\n')
    await tick()

    await expect(client.request('later', {})).rejects.toMatchObject({ code: 'rpc_protocol_error' })
    expect(diagnostic).toHaveBeenCalledWith({ code: 'rpc_protocol_error' })
  })

  it('maps server errors without retaining their message or data', async () => {
    const { client, fromServer } = createTransport()
    const request = client.request('test', {})
    fromServer.write(
      '{"jsonrpc":"2.0","id":1,"error":{"code":-32000,"message":"private prompt","data":{"secret":"token"}}}\n',
    )

    const error = await request.catch((reason: unknown) => reason)
    expect(error).toBeInstanceOf(JsonRpcError)
    expect(error).toMatchObject({ code: 'rpc_remote_error', remoteCode: -32000 })
    expect(String(error)).not.toContain('private prompt')
    expect(JSON.stringify(error)).not.toContain('token')
    await client.close()
  })

  it('times out one request without resolving it later', async () => {
    vi.useFakeTimers()
    try {
      const { client, fromServer } = createTransport({ requestTimeoutMs: 20 })
      const request = client.request('slow', {})
      const rejection = expect(request).rejects.toMatchObject({ code: 'rpc_request_timeout' })
      await vi.advanceTimersByTimeAsync(20)
      await rejection

      fromServer.write('{"jsonrpc":"2.0","id":1,"result":{"private":"prompt"}}\n')
      await expect(client.request('closed', {})).rejects.toMatchObject({
        code: 'rpc_protocol_error',
      })
    } finally {
      vi.useRealTimers()
    }
  })

  it('supports AbortSignal and removes its listener after settlement', async () => {
    const { client } = createTransport()
    const controller = new AbortController()
    const request = client.request('slow', {}, { signal: controller.signal })
    controller.abort('private prompt')

    await expect(request).rejects.toMatchObject({ code: 'rpc_request_aborted' })
    await client.close()
  })

  it('bounds individual lines and the aggregate receive buffer', async () => {
    const tooLong = createTransport({ maxLineBytes: 64, maxBufferBytes: 128 })
    const pendingLine = tooLong.client.request('test', {})
    tooLong.fromServer.write('x'.repeat(65))
    await expect(pendingLine).rejects.toMatchObject({ code: 'rpc_line_limit_exceeded' })

    const tooMuch = createTransport({ maxLineBytes: 64, maxBufferBytes: 32 })
    const pendingBuffer = tooMuch.client.request('test', {})
    tooMuch.fromServer.write('x'.repeat(33))
    await expect(pendingBuffer).rejects.toMatchObject({ code: 'rpc_buffer_limit_exceeded' })
  })

  it('rejects oversized or non-serializable outbound messages before writing', async () => {
    const oversized = createTransport({ maxLineBytes: 64, maxBufferBytes: 128 })
    await expect(
      oversized.client.request('test', { value: 'x'.repeat(100) }),
    ).rejects.toMatchObject({ code: 'rpc_request_limit_exceeded' })
    expect(oversized.writes).toHaveLength(0)
    await oversized.client.close()

    const cyclic = createTransport()
    const params: Record<string, unknown> = {}
    params.self = params
    await expect(cyclic.client.request('test', params)).rejects.toMatchObject({
      code: 'rpc_invalid_request',
    })
    expect(cyclic.writes).toHaveLength(0)
    await cyclic.client.close()

    const accessor = createTransport()
    const accessorParams = Object.defineProperty({}, 'secret', { get: () => 'private' })
    await expect(accessor.client.request('test', accessorParams)).rejects.toMatchObject({
      code: 'rpc_invalid_request',
    })
    expect(accessor.writes).toHaveLength(0)
    await accessor.client.close()
  })

  it('rejects every pending request on input crash and close is idempotent', async () => {
    const { client, fromServer } = createTransport()
    const first = client.request('first', {})
    const second = client.request('second', {})
    fromServer.destroy(new Error('private crash details'))

    await expect(first).rejects.toMatchObject({ code: 'rpc_transport_closed' })
    await expect(second).rejects.toMatchObject({ code: 'rpc_transport_closed' })
    await expect(client.close()).resolves.toBeUndefined()
    await expect(client.close()).resolves.toBeUndefined()
  })

  it('bounds pending concurrency and queued backpressure', async () => {
    const pending = createTransport({ maxPendingRequests: 1 })
    const first = pending.client.request('first', {})
    await expect(pending.client.request('second', {})).rejects.toMatchObject({
      code: 'rpc_pending_limit_exceeded',
    })
    pending.fromServer.write('{"jsonrpc":"2.0","id":1,"result":{}}\n')
    await first
    await pending.client.close()

    const fromServer = new PassThrough()
    const blocked = new Writable({
      write(_chunk, _encoding, _callback) {
        // Intentionally retain the callback to simulate a blocked child stdin.
      },
    })
    const client = new JsonRpcClient({
      input: fromServer,
      output: blocked,
      maxQueuedWriteBytes: 80,
    })
    const queued = client.request('first', {})
    await expect(client.request('second', {})).rejects.toMatchObject({
      code: 'rpc_write_queue_limit_exceeded',
    })
    fromServer.destroy()
    await expect(queued).rejects.toMatchObject({ code: 'rpc_transport_closed' })
    await client.close()
  })

  it('rejects one giant inbound chunk against the budget', async () => {
    const fixture = createTransport({ maxLineBytes: 64, maxBufferBytes: 64 })
    const pending = fixture.client.request('test', {})
    fixture.fromServer.write(Buffer.alloc(1_000_000, 0x78))
    await expect(pending).rejects.toMatchObject({ code: 'rpc_line_limit_exceeded' })
  })

  it('contains late stream errors after a protocol-driven close', async () => {
    const { client, fromServer, toServer } = createTransport()
    const pending = client.request('test', {})
    fromServer.write('malformed\n')
    await expect(pending).rejects.toMatchObject({ code: 'rpc_protocol_error' })

    expect(() => fromServer.emit('error', new Error('private late input error'))).not.toThrow()
    expect(() => toServer.emit('error', new Error('private late output error'))).not.toThrow()
  })
})
