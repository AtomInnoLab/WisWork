import { randomBytes, timingSafeEqual } from 'node:crypto'
import http, { type IncomingMessage, type ServerResponse } from 'node:http'
import type { AddressInfo, Socket } from 'node:net'
import { TextDecoder } from 'node:util'
import type { MessagesRequest, PreparedResponsesTurn } from './types.js'

const HOST = '127.0.0.1'
const PATH = '/v1/responses'
const MAX_BODY = 8_000_000
const MAX_TURNS = 8
const MAX_TURN_MS = 10 * 60_000
const MAX_IDLE_MS = 30_000
const MAX_STREAM_BYTES = 32_000_000
const MAX_STREAM_FRAMES = 100_000

export interface ResponsesBridgeOptions {
  /** The sole authenticated WisUsage transport. It receives an already-fixed request. */
  readonly fetchWithAuth: (request: MessagesRequest, signal: AbortSignal) => Promise<Response>
  /** Host-owned, one-use Task 2 carrier closure. Never derive authority from request metadata. */
  readonly prepareTurn: (input: unknown) => PreparedResponsesTurn
  readonly maxBodyBytes?: number
  readonly maxActiveTurns?: number
  readonly maxTurnDurationMs?: number
  readonly maxStreamIdleMs?: number
}

export interface ResponsesBridge {
  readonly baseUrl: string
  readonly responsesUrl: string
  readonly secret: string
  close(): Promise<void>
}

class BodyLimitError extends Error {}
class RequestEndedError extends Error {}

function downward(value: number | undefined, maximum: number, code: string): number {
  const result = value ?? maximum
  if (!Number.isSafeInteger(result) || result <= 0 || result > maximum) throw new TypeError(code)
  return result
}

function authorized(header: string | undefined, expected: Buffer): boolean {
  if (typeof header !== 'string' || !/^Bearer [A-Za-z0-9_-]{43}$/.test(header)) return false
  let candidate: Buffer
  try {
    candidate = Buffer.from(header.slice(7), 'base64url')
  } catch {
    return false
  }
  return (
    candidate.length === expected.length &&
    candidate.toString('base64url') === header.slice(7) &&
    timingSafeEqual(candidate, expected)
  )
}

function sendError(response: ServerResponse, status: number, code: string): void {
  if (response.destroyed || response.headersSent) return
  const encoded = Buffer.from(JSON.stringify({ error: { code, message: code } }))
  response.writeHead(status, {
    'cache-control': 'no-store',
    connection: 'close',
    'content-type': 'application/json; charset=utf-8',
    'content-length': String(encoded.length),
    'x-content-type-options': 'nosniff',
  })
  response.end(encoded)
}

function readBody(request: IncomingMessage, maximum: number, signal: AbortSignal): Promise<Buffer> {
  const declared = request.headers['content-length']
  if (declared !== undefined && (!/^\d+$/.test(declared) || Number(declared) > maximum)) {
    request.resume()
    return Promise.reject(new BodyLimitError())
  }
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    let bytes = 0
    let done = false
    const cleanup = (): void => {
      request.off('data', onData)
      request.off('end', onEnd)
      request.off('error', onError)
      request.off('aborted', onError)
      signal.removeEventListener('abort', onError)
    }
    const finish = (action: () => void): void => {
      if (done) return
      done = true
      cleanup()
      action()
    }
    const onData = (chunk: Buffer): void => {
      bytes += chunk.length
      if (bytes > maximum) {
        finish(() => reject(new BodyLimitError()))
        request.resume()
      } else chunks.push(chunk)
    }
    const onEnd = (): void => finish(() => resolve(Buffer.concat(chunks, bytes)))
    const onError = (): void => finish(() => reject(new RequestEndedError()))
    request.on('data', onData)
    request.once('end', onEnd)
    request.once('error', onError)
    request.once('aborted', onError)
    signal.addEventListener('abort', onError, { once: true })
    if (signal.aborted) onError()
  })
}

function parseJson(body: Buffer): unknown {
  return JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(body)) as unknown
}

function exactEventStream(value: string | null): boolean {
  if (value === null) return false
  const parts = value.split(';').map((part) => part.trim().toLowerCase())
  return (
    parts.length >= 1 &&
    parts.length <= 2 &&
    parts[0] === 'text/event-stream' &&
    (parts.length === 1 || parts[1] === 'charset=utf-8')
  )
}

function cancelResponse(response: Response): void {
  try {
    void response.body?.cancel().catch(() => undefined)
  } catch {
    /* best effort */
  }
}

async function* upstreamChunks(
  body: ReadableStream<Uint8Array>,
  signal: AbortSignal,
  idleMs: number,
): AsyncGenerator<Uint8Array> {
  const reader = body.getReader()
  let bytes = 0
  let frames = 0
  let completed = false
  try {
    while (true) {
      if (signal.aborted) throw new RequestEndedError()
      let timer: ReturnType<typeof setTimeout> | undefined
      const timeout = new Promise<never>((_, reject) => {
        timer = setTimeout(() => reject(new RequestEndedError()), idleMs)
        timer.unref()
      })
      const item = await Promise.race([reader.read(), timeout]).finally(() => {
        if (timer) clearTimeout(timer)
      })
      if (item.done) {
        completed = true
        return
      }
      bytes += item.value.length
      frames += 1
      if (bytes > MAX_STREAM_BYTES || frames > MAX_STREAM_FRAMES) throw new RequestEndedError()
      yield item.value
    }
  } finally {
    if (!completed) void reader.cancel().catch(() => undefined)
    try {
      reader.releaseLock()
    } catch {
      /* hostile stream */
    }
  }
}

function awaitAbortable<T>(pending: Promise<T>, signal: AbortSignal): Promise<T> {
  return new Promise((resolve, reject) => {
    let settled = false
    const done = (action: () => void): void => {
      if (settled) return
      settled = true
      signal.removeEventListener('abort', abort)
      action()
    }
    const abort = (): void => done(() => reject(new RequestEndedError()))
    signal.addEventListener('abort', abort, { once: true })
    pending.then(
      (value) => done(() => resolve(value)),
      () => done(() => reject(new Error('upstream_error'))),
    )
    if (signal.aborted) abort()
  })
}

export async function startResponsesBridge(
  options: ResponsesBridgeOptions,
): Promise<ResponsesBridge> {
  if (
    !options ||
    typeof options.fetchWithAuth !== 'function' ||
    typeof options.prepareTurn !== 'function'
  )
    throw new TypeError('invalid_bridge_options')
  const maxBodyBytes = downward(options.maxBodyBytes, MAX_BODY, 'invalid_max_body_bytes')
  const maxActiveTurns = downward(options.maxActiveTurns, MAX_TURNS, 'invalid_max_active_turns')
  const maxTurnDurationMs = downward(
    options.maxTurnDurationMs,
    MAX_TURN_MS,
    'invalid_max_turn_duration',
  )
  const maxStreamIdleMs = downward(options.maxStreamIdleMs, MAX_IDLE_MS, 'invalid_max_stream_idle')
  const secretBytes = randomBytes(32)
  const secret = secretBytes.toString('base64url')
  const controllers = new Set<AbortController>()
  const sockets = new Set<Socket>()
  let active = 0
  let closing = false

  const server = http.createServer((request, response) => {
    const controller = new AbortController()
    controllers.add(controller)
    const abort = (): void => controller.abort()
    request.once('aborted', abort)
    response.setTimeout(maxStreamIdleMs, () => {
      abort()
      response.destroy()
    })
    response.once('close', () => {
      if (!response.writableEnded) abort()
    })
    void (async () => {
      let turnSlot = false
      let timedOut = false
      let timer: ReturnType<typeof setTimeout> | undefined
      try {
        if (closing) return sendError(response, 503, 'bridge_closed')
        // Authentication deliberately precedes routing, content type, length and JSON parsing.
        if (!authorized(request.headers.authorization, secretBytes)) {
          request.resume()
          return sendError(response, 401, 'unauthorized')
        }
        if (request.url !== PATH) {
          request.resume()
          return sendError(response, 404, 'not_found')
        }
        if (request.method !== 'POST') {
          request.resume()
          return sendError(response, 405, 'method_not_allowed')
        }
        if (request.headers['content-type']?.toLowerCase() !== 'application/json') {
          request.resume()
          return sendError(response, 415, 'unsupported_media_type')
        }
        if (active >= maxActiveTurns) {
          request.resume()
          return sendError(response, 429, 'bridge_busy')
        }
        active += 1
        turnSlot = true
        timer = setTimeout(() => {
          timedOut = true
          controller.abort()
        }, maxTurnDurationMs)
        timer.unref()
        let raw: Buffer
        try {
          raw = await readBody(request, maxBodyBytes, controller.signal)
        } catch (error) {
          if (error instanceof BodyLimitError) sendError(response, 413, 'request_too_large')
          else if (timedOut) sendError(response, 504, 'turn_timeout')
          else if (!controller.signal.aborted) sendError(response, 400, 'invalid_request')
          return
        }
        let input: unknown
        try {
          input = parseJson(raw)
        } catch {
          return sendError(response, 400, 'invalid_request')
        }
        let turn: PreparedResponsesTurn
        try {
          turn = options.prepareTurn(input)
        } catch {
          return sendError(response, 400, 'invalid_request')
        }
        if (
          !turn ||
          turn.messagesRequest?.model !== 'openai/gpt-5.6-sol' ||
          turn.messagesRequest.stream !== true ||
          typeof turn.messagesStreamToResponses !== 'function'
        ) {
          return sendError(response, 400, 'invalid_request')
        }
        let upstream: Response
        try {
          upstream = await awaitAbortable(
            options.fetchWithAuth(turn.messagesRequest, controller.signal),
            controller.signal,
          )
        } catch {
          if (timedOut) sendError(response, 504, 'turn_timeout')
          else if (!controller.signal.aborted) sendError(response, 502, 'upstream_error')
          return
        }
        if (
          !upstream.ok ||
          upstream.body === null ||
          !exactEventStream(upstream.headers.get('content-type'))
        ) {
          cancelResponse(upstream)
          return sendError(
            response,
            upstream.status === 401 ? 401 : 502,
            upstream.status === 401 ? 'auth_required' : 'upstream_error',
          )
        }
        response.writeHead(200, {
          'cache-control': 'no-cache, no-store',
          'content-type': 'text/event-stream; charset=utf-8',
          'x-content-type-options': 'nosniff',
        })
        try {
          for await (const frame of turn.messagesStreamToResponses(
            upstreamChunks(upstream.body, controller.signal, maxStreamIdleMs),
          )) {
            if (controller.signal.aborted) throw new RequestEndedError()
            if (!response.write(frame))
              await new Promise<void>((resolve, reject) => {
                response.once('drain', resolve)
                response.once('close', () => reject(new RequestEndedError()))
              })
          }
          response.end()
        } catch {
          cancelResponse(upstream)
          if (!response.destroyed) response.destroy()
        }
      } catch {
        if (!response.headersSent) sendError(response, 500, 'bridge_error')
        else if (!response.destroyed) response.destroy()
      } finally {
        if (timer) clearTimeout(timer)
        if (turnSlot) active -= 1
        controllers.delete(controller)
        request.off('aborted', abort)
      }
    })()
  })
  server.headersTimeout = 10_000
  server.requestTimeout = 30_000
  server.keepAliveTimeout = 5_000
  server.maxRequestsPerSocket = 100
  server.on('connection', (socket) => {
    sockets.add(socket)
    socket.setTimeout(MAX_IDLE_MS, () => socket.destroy())
    socket.once('close', () => sockets.delete(socket))
  })
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject)
    server.listen(0, HOST, resolve)
  })
  const address = server.address() as AddressInfo
  const baseUrl = `http://${HOST}:${address.port}`
  let closePromise: Promise<void> | undefined
  return Object.freeze({
    baseUrl,
    responsesUrl: `${baseUrl}${PATH}`,
    secret,
    close() {
      if (closePromise) return closePromise
      closing = true
      for (const controller of controllers) controller.abort()
      for (const socket of sockets) socket.destroy()
      closePromise = new Promise((resolve) => server.close(() => resolve()))
      return closePromise
    },
  })
}
