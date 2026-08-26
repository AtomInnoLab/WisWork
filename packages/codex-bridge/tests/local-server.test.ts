import http from 'node:http'
import { once } from 'node:events'
import { describe, expect, it, vi } from 'vitest'
import { startResponsesBridge, type MessagesRequest, type ResponsesBridge } from '../src/index.js'

const messageStart =
  'event: message_start\ndata: {"type":"message_start","message":{"id":"msg_1","model":"openai/gpt-5.6-sol","usage":{"input_tokens":2}}}\n\n'
const textStart =
  'event: content_block_start\ndata: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":""}}\n\n'
const textDelta =
  'event: content_block_delta\ndata: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"Hello"}}\n\n'
const textStop = 'event: content_block_stop\ndata: {"type":"content_block_stop","index":0}\n\n'
const messageDelta =
  'event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":1}}\n\n'
const messageStop = 'event: message_stop\ndata: {"type":"message_stop"}\n\n'
const upstreamSse = messageStart + textStart + textDelta + textStop + messageDelta + messageStop

function chunkedResponse(source = upstreamSse, chunkSize = source.length): Response {
  const bytes = new TextEncoder().encode(source)
  return new Response(
    new ReadableStream<Uint8Array>({
      start(controller) {
        for (let offset = 0; offset < bytes.length; offset += chunkSize) {
          controller.enqueue(bytes.slice(offset, offset + chunkSize))
        }
        controller.close()
      },
    }),
    { status: 200, headers: { 'content-type': 'text/event-stream' } },
  )
}

interface RawResponse {
  status: number
  headers: http.IncomingHttpHeaders
  body: string
}

function rawRequest(
  url: string,
  options: {
    method?: string
    token?: string
    contentType?: string
    body?: string
  } = {},
): Promise<RawResponse> {
  const target = new URL(url)
  const body = options.body ?? ''
  return new Promise((resolve, reject) => {
    const request = http.request(
      {
        hostname: target.hostname,
        port: target.port,
        path: target.pathname,
        method: options.method ?? 'POST',
        headers: {
          ...(options.token ? { authorization: `Bearer ${options.token}` } : {}),
          ...(options.contentType ? { 'content-type': options.contentType } : {}),
          'content-length': Buffer.byteLength(body),
        },
      },
      (response) => {
        const chunks: Buffer[] = []
        response.on('data', (chunk: Buffer) => chunks.push(chunk))
        response.on('end', () => {
          resolve({
            status: response.statusCode ?? 0,
            headers: response.headers,
            body: Buffer.concat(chunks).toString('utf8'),
          })
        })
      },
    )
    request.on('error', reject)
    request.end(body)
  })
}

const validBody = (input = 'Hello') =>
  JSON.stringify({ model: 'gpt-5.6-sol', input, stream: true, max_output_tokens: 123 })

async function start(
  fetchWithAuth: (request: MessagesRequest, signal: AbortSignal) => Promise<Response> = async () =>
    chunkedResponse(),
  options: { maxBodyBytes?: number } = {},
): Promise<ResponsesBridge> {
  return startResponsesBridge({ fetchWithAuth, ...options })
}

describe('authenticated loopback Responses server', () => {
  it('binds a random OS-assigned loopback port and creates a fresh high-entropy secret', async () => {
    const first = await startResponsesBridge({
      fetchWithAuth: async () => chunkedResponse(),
      ...({ host: '0.0.0.0', port: 1 } as Record<string, unknown>),
    })
    const second = await start()
    try {
      const firstUrl = new URL(first.baseUrl)
      const secondUrl = new URL(second.baseUrl)
      expect(firstUrl.hostname).toBe('127.0.0.1')
      expect(Number(firstUrl.port)).toBeGreaterThan(0)
      expect(secondUrl.port).not.toBe(firstUrl.port)
      expect(first.responsesUrl).toBe(`${first.baseUrl}/v1/responses`)
      expect(first.secret).not.toBe(second.secret)
      expect(first.secret).toMatch(/^[A-Za-z0-9_-]{43}$/)
    } finally {
      await Promise.all([first.close(), second.close()])
    }
  })

  it.each([
    ['missing token', undefined],
    ['wrong token of equal length', 'x'.repeat(43)],
    ['wrong token of different length', 'short'],
  ])('rejects %s without invoking upstream', async (_label, token) => {
    const fetchWithAuth = vi.fn(async () => chunkedResponse())
    const bridge = await start(fetchWithAuth)
    try {
      const response = await rawRequest(bridge.responsesUrl, {
        token,
        contentType: 'application/json',
        body: validBody(),
      })
      expect(response.status).toBe(401)
      expect(response.body).toBe('{"error":{"code":"unauthorized","message":"Unauthorized"}}')
      expect(fetchWithAuth).not.toHaveBeenCalled()
    } finally {
      await bridge.close()
    }
  })

  it.each([
    ['wrong path', '/v1/chat/completions', 'POST', 'application/json', 404, 'not_found'],
    ['wrong method', '/v1/responses', 'GET', 'application/json', 405, 'method_not_allowed'],
    ['missing media type', '/v1/responses', 'POST', undefined, 415, 'unsupported_media_type'],
    ['wrong media type', '/v1/responses', 'POST', 'text/plain', 415, 'unsupported_media_type'],
  ])(
    'rejects %s without invoking upstream',
    async (_label, path, method, contentType, status, code) => {
      const fetchWithAuth = vi.fn(async () => chunkedResponse())
      const bridge = await start(fetchWithAuth)
      try {
        const response = await rawRequest(`${bridge.baseUrl}${path}`, {
          method,
          token: bridge.secret,
          contentType,
          body: validBody(),
        })
        expect(response.status).toBe(status)
        expect(JSON.parse(response.body).error.code).toBe(code)
        expect(fetchWithAuth).not.toHaveBeenCalled()
      } finally {
        await bridge.close()
      }
    },
  )

  it.each([
    ['malformed JSON', '{"prompt":"secret"'],
    ['non-object JSON', '"secret prompt"'],
    ['unsupported field', JSON.stringify({ model: 'gpt-5.6-sol', input: 'secret', user: 'x' })],
    ['endpoint/model override', JSON.stringify({ model: 'attacker/model', input: 'secret' })],
  ])('fails closed for %s with a redacted 400', async (_label, body) => {
    const fetchWithAuth = vi.fn(async () => chunkedResponse())
    const bridge = await start(fetchWithAuth)
    try {
      const response = await rawRequest(bridge.responsesUrl, {
        token: bridge.secret,
        contentType: 'application/json; charset=utf-8',
        body,
      })
      expect(response.status).toBe(400)
      expect(response.body).not.toContain('secret')
      expect(response.body).not.toContain('attacker')
      expect(fetchWithAuth).not.toHaveBeenCalled()
    } finally {
      await bridge.close()
    }
  })

  it('caps raw bytes before parsing and never invokes upstream', async () => {
    const fetchWithAuth = vi.fn(async () => chunkedResponse())
    const bridge = await start(fetchWithAuth, { maxBodyBytes: 64 })
    try {
      const response = await rawRequest(bridge.responsesUrl, {
        token: bridge.secret,
        contentType: 'application/json',
        body: `${' '.repeat(65)}{not parsed}`,
      })
      expect(response.status).toBe(413)
      expect(response.body).toContain('request_too_large')
      expect(fetchWithAuth).not.toHaveBeenCalled()
    } finally {
      await bridge.close()
    }
  })

  it('caps a chunked raw body before parsing and never invokes upstream', async () => {
    const fetchWithAuth = vi.fn(async () => chunkedResponse())
    const bridge = await start(fetchWithAuth, { maxBodyBytes: 64 })
    try {
      const target = new URL(bridge.responsesUrl)
      const result = new Promise<RawResponse>((resolve, reject) => {
        const request = http.request(
          {
            hostname: target.hostname,
            port: target.port,
            path: target.pathname,
            method: 'POST',
            headers: {
              authorization: `Bearer ${bridge.secret}`,
              'content-type': 'application/json',
            },
          },
          (response) => {
            const chunks: Buffer[] = []
            response.on('data', (chunk: Buffer) => chunks.push(chunk))
            response.on('end', () =>
              resolve({
                status: response.statusCode ?? 0,
                headers: response.headers,
                body: Buffer.concat(chunks).toString('utf8'),
              }),
            )
          },
        )
        request.on('error', reject)
        request.write(' '.repeat(40))
        request.end(' '.repeat(40))
      })
      expect((await result).status).toBe(413)
      expect(fetchWithAuth).not.toHaveBeenCalled()
    } finally {
      await bridge.close()
    }
  })

  it('converts the request, fixes the upstream model, and streams Responses SSE across chunks', async () => {
    const calls: MessagesRequest[] = []
    const fetchWithAuth = vi.fn(async (request: MessagesRequest) => {
      calls.push(request)
      return chunkedResponse(upstreamSse, 1)
    })
    const bridge = await start(fetchWithAuth)
    try {
      const response = await rawRequest(bridge.responsesUrl, {
        token: bridge.secret,
        contentType: 'application/json',
        body: validBody('actual prompt'),
      })
      expect(response.status).toBe(200)
      expect(response.headers['content-type']).toContain('text/event-stream')
      expect(calls).toEqual([
        {
          model: 'openai/gpt-5.6-sol',
          messages: [{ role: 'user', content: [{ type: 'text', text: 'actual prompt' }] }],
          max_tokens: 123,
          stream: true,
        },
      ])
      expect(response.body).toContain('event: response.output_text.delta')
      expect(response.body).toContain('"delta":"Hello"')
      expect(response.body).toContain('event: response.completed')
      expect(response.body.endsWith('data: [DONE]\n\n')).toBe(true)
      expect(fetchWithAuth).toHaveBeenCalledTimes(1)
    } finally {
      await bridge.close()
    }
  })

  it.each([
    [401, 401, 'auth_required'],
    [403, 502, 'upstream_error'],
    [500, 502, 'upstream_error'],
  ])('redacts upstream HTTP %i without reading its body', async (upstreamStatus, status, code) => {
    const response = new Response(
      new ReadableStream({
        pull(controller) {
          controller.enqueue(new TextEncoder().encode('secret upstream token and prompt'))
          controller.close()
        },
      }),
      { status: upstreamStatus },
    )
    const fetchWithAuth = vi.fn(async () => response)
    const bridge = await start(fetchWithAuth)
    try {
      const result = await rawRequest(bridge.responsesUrl, {
        token: bridge.secret,
        contentType: 'application/json',
        body: validBody(),
      })
      expect(result.status).toBe(status)
      expect(result.body).toContain(code)
      expect(result.body).not.toContain('secret')
      expect(fetchWithAuth).toHaveBeenCalledTimes(1)
    } finally {
      await bridge.close()
    }
  })

  it('maps a thrown auth error without exposing exception text', async () => {
    const bridge = await start(async () => {
      throw Object.assign(new Error('secret token in exception'), { code: 'auth_required' })
    })
    try {
      const response = await rawRequest(bridge.responsesUrl, {
        token: bridge.secret,
        contentType: 'application/json',
        body: validBody(),
      })
      expect(response.status).toBe(401)
      expect(response.body).toContain('auth_required')
      expect(response.body).not.toContain('secret')
    } finally {
      await bridge.close()
    }
  })

  it('aborts upstream when the client disconnects', async () => {
    let upstreamSignal: AbortSignal | undefined
    let cancelObserved = false
    const bridge = await start(async (_request, signal) => {
      upstreamSignal = signal
      return new Response(
        new ReadableStream<Uint8Array>({
          start(controller) {
            controller.enqueue(new TextEncoder().encode(messageStart))
          },
          cancel() {
            cancelObserved = true
          },
        }),
      )
    })
    try {
      const target = new URL(bridge.responsesUrl)
      const request = http.request({
        hostname: target.hostname,
        port: target.port,
        path: target.pathname,
        method: 'POST',
        headers: {
          authorization: `Bearer ${bridge.secret}`,
          'content-type': 'application/json',
        },
      })
      request.end(validBody())
      const [response] = (await once(request, 'response')) as [http.IncomingMessage]
      await once(response, 'data')
      response.destroy()
      await vi.waitFor(() => expect(upstreamSignal?.aborted).toBe(true))
      await vi.waitFor(() => expect(cancelObserved).toBe(true))
    } finally {
      await bridge.close()
    }
  })

  it('aborts active upstream work on close and close is idempotent', async () => {
    let upstreamSignal: AbortSignal | undefined
    const bridge = await start(async (_request, signal) => {
      upstreamSignal = signal
      return new Promise<Response>((_resolve, reject) => {
        signal.addEventListener('abort', () => reject(signal.reason), { once: true })
      })
    })
    const pending = rawRequest(bridge.responsesUrl, {
      token: bridge.secret,
      contentType: 'application/json',
      body: validBody(),
    }).catch(() => undefined)
    await vi.waitFor(() => expect(upstreamSignal).toBeDefined())
    await expect(Promise.all([bridge.close(), bridge.close(), bridge.close()])).resolves.toEqual([
      undefined,
      undefined,
      undefined,
    ])
    expect(upstreamSignal?.aborted).toBe(true)
    await pending
  })

  it('supports concurrent requests and cleans them up before close resolves', async () => {
    const fetchWithAuth = vi.fn(async () => chunkedResponse())
    const bridge = await start(fetchWithAuth)
    const responses = await Promise.all(
      Array.from({ length: 8 }, () =>
        rawRequest(bridge.responsesUrl, {
          token: bridge.secret,
          contentType: 'application/json',
          body: validBody(),
        }),
      ),
    )
    expect(responses.every((response) => response.status === 200)).toBe(true)
    expect(fetchWithAuth).toHaveBeenCalledTimes(8)
    await bridge.close()
  })
})
