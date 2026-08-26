import http, { type IncomingMessage, type ServerResponse } from 'node:http'
import type { AddressInfo, Socket } from 'node:net'
import { TextDecoder } from 'node:util'
import { prepareResponsesTurn, ProtocolCompatibilityError } from './index.js'
import { createBridgeSecret, hasValidBearerToken } from './security.js'
import type { MessagesRequest } from './types.js'

const LOOPBACK_HOST = '127.0.0.1'
const RESPONSES_PATH = '/v1/responses'
const DEFAULT_MAX_BODY_BYTES = 8_000_000
const MAX_BODY_BYTES = 8_000_000
const HEADERS_TIMEOUT_MS = 10_000
const REQUEST_TIMEOUT_MS = 30_000
const RESPONSE_IDLE_TIMEOUT_MS = 30_000
const KEEP_ALIVE_TIMEOUT_MS = 5_000
const MAX_ACTIVE_TURNS = 8
const MAX_TURN_DURATION_MS = 10 * 60 * 1_000

export interface ResponsesBridgeOptions {
  /** Must perform the AuthClient.fetchWithAuth call and honor abort to avoid orphaned work. */
  fetchWithAuth: (request: MessagesRequest, signal: AbortSignal) => Promise<Response>
  /** Test/deployment tightening only; values above the built-in ceiling are rejected. */
  maxBodyBytes?: number
  /** Test/deployment tightening only; values above the built-in ceiling are rejected. */
  maxActiveTurns?: number
  /** Test/deployment tightening only; values above the built-in ceiling are rejected. */
  maxTurnDurationMs?: number
}

export interface ResponsesBridge {
  readonly baseUrl: string
  readonly responsesUrl: string
  readonly secret: string
  close(): Promise<void>
}

class BodyTooLargeError extends Error {}
class RequestReadError extends Error {}

function jsonError(
  response: ServerResponse,
  status: number,
  code: string,
  message: string,
  extraHeaders: Record<string, string> = {},
): void {
  if (response.headersSent || response.destroyed) return
  const body = JSON.stringify({ error: { code, message } })
  response.writeHead(status, {
    'cache-control': 'no-store',
    connection: 'close',
    'content-type': 'application/json; charset=utf-8',
    'content-length': Buffer.byteLength(body),
    ...extraHeaders,
  })
  response.end(body)
}

function hasExactContentType(value: string | null | undefined, mediaType: string): boolean {
  if (value === undefined || value === null) return false
  const parts = value.split(';').map((part) => part.trim().toLowerCase())
  return (
    parts[0] === mediaType &&
    parts.length <= 2 &&
    (parts.length === 1 || parts[1] === 'charset=utf-8')
  )
}

const isJsonContentType = (value: string | undefined): boolean =>
  hasExactContentType(value, 'application/json')

const isEventStreamContentType = (value: string | null): boolean =>
  hasExactContentType(value, 'text/event-stream')

function readRawBody(
  request: IncomingMessage,
  maxBytes: number,
  signal: AbortSignal,
): Promise<Buffer> {
  const declaredLength = request.headers['content-length']
  if (declaredLength !== undefined) {
    const length = Number(declaredLength)
    if (!Number.isSafeInteger(length) || length < 0) {
      request.resume()
      return Promise.reject(new RequestReadError())
    }
    if (length > maxBytes) {
      request.resume()
      return Promise.reject(new BodyTooLargeError())
    }
  }

  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []
    let bytes = 0
    let settled = false

    const cleanup = (): void => {
      request.off('data', onData)
      request.off('end', onEnd)
      request.off('error', onError)
      request.off('aborted', onAborted)
      signal.removeEventListener('abort', onSignalAbort)
    }
    const finish = (action: () => void): void => {
      if (settled) return
      settled = true
      cleanup()
      action()
    }
    const onData = (chunk: Buffer): void => {
      bytes += chunk.byteLength
      if (bytes > maxBytes) {
        finish(() => reject(new BodyTooLargeError()))
        request.resume()
        return
      }
      chunks.push(chunk)
    }
    const onEnd = (): void => finish(() => resolve(Buffer.concat(chunks, bytes)))
    const onError = (): void => finish(() => reject(new RequestReadError()))
    const onAborted = (): void => finish(() => reject(new RequestReadError()))
    const onSignalAbort = (): void => finish(() => reject(new RequestReadError()))

    request.on('data', onData)
    request.once('end', onEnd)
    request.once('error', onError)
    request.once('aborted', onAborted)
    signal.addEventListener('abort', onSignalAbort, { once: true })
    if (signal.aborted) onSignalAbort()
  })
}

function parseJsonBody(body: Buffer): unknown {
  const source = new TextDecoder('utf-8', { fatal: true }).decode(body)
  return JSON.parse(source) as unknown
}

function isAuthRequired(error: unknown): boolean {
  return (
    typeof error === 'object' &&
    error !== null &&
    'code' in error &&
    (error as { code?: unknown }).code === 'auth_required'
  )
}

function cancelBody(response: Response): void {
  try {
    void response.body?.cancel().catch(() => undefined)
  } catch {
    // Cancellation is best-effort after a redacted status response.
  }
}

function awaitAbortableUpstream(
  pending: Promise<Response>,
  signal: AbortSignal,
): Promise<Response> {
  return new Promise((resolve, reject) => {
    let waiting = true
    const cleanup = (): void => signal.removeEventListener('abort', onAbort)
    const onAbort = (): void => {
      if (!waiting) return
      waiting = false
      cleanup()
      reject(new RequestReadError())
    }
    signal.addEventListener('abort', onAbort, { once: true })
    pending.then(
      (response) => {
        if (!waiting) {
          cancelBody(response)
          return
        }
        waiting = false
        cleanup()
        resolve(response)
      },
      (error: unknown) => {
        if (!waiting) return
        waiting = false
        cleanup()
        reject(error)
      },
    )
    if (signal.aborted) onAbort()
  })
}

async function* responseChunks(
  body: ReadableStream<Uint8Array>,
  signal: AbortSignal,
): AsyncGenerator<Uint8Array> {
  const reader = body.getReader()
  let ended = false
  const cancel = (): void => {
    void reader.cancel().catch(() => undefined)
  }
  signal.addEventListener('abort', cancel, { once: true })
  try {
    while (true) {
      if (signal.aborted) throw new RequestReadError()
      const chunk = await reader.read()
      if (signal.aborted) throw new RequestReadError()
      if (chunk.done) {
        ended = true
        return
      }
      yield chunk.value
    }
  } finally {
    signal.removeEventListener('abort', cancel)
    if (!ended) void reader.cancel().catch(() => undefined)
    try {
      reader.releaseLock()
    } catch {
      // A hostile stream must not block bridge cancellation or shutdown.
    }
  }
}

function waitForDrain(response: ServerResponse, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    const cleanup = (): void => {
      response.off('drain', onDrain)
      response.off('close', onClose)
      response.off('error', onError)
      signal.removeEventListener('abort', onAbort)
    }
    const onDrain = (): void => {
      cleanup()
      resolve()
    }
    const onClose = (): void => {
      cleanup()
      reject(new RequestReadError())
    }
    const onError = (): void => {
      cleanup()
      reject(new RequestReadError())
    }
    const onAbort = (): void => {
      cleanup()
      reject(new RequestReadError())
    }
    response.once('drain', onDrain)
    response.once('close', onClose)
    response.once('error', onError)
    signal.addEventListener('abort', onAbort, { once: true })
    if (signal.aborted) onAbort()
  })
}

function terminalFailureFrame(responseId: string, code: string, message: string): string {
  return `event: response.failed\ndata: ${JSON.stringify({
    type: 'response.failed',
    response: {
      id: responseId,
      object: 'response',
      model: 'gpt-5.6-sol',
      status: 'failed',
      output: [],
      error: { code, message },
    },
  })}\n\n`
}

function responseIdFromCreatedFrame(frame: string): string {
  try {
    const data = /^data: (.+)$/m.exec(frame)?.[1]
    const parsed = data === undefined ? undefined : (JSON.parse(data) as unknown)
    if (
      typeof parsed === 'object' &&
      parsed !== null &&
      'response' in parsed &&
      typeof parsed.response === 'object' &&
      parsed.response !== null &&
      'id' in parsed.response &&
      typeof parsed.response.id === 'string'
    ) {
      return parsed.response.id
    }
  } catch {
    // The frame is generated by the bound converter; keep the fallback redacted.
  }
  return 'response_failed'
}

async function streamResponse(
  response: ServerResponse,
  upstream: Response,
  iterator: AsyncIterator<string>,
  firstFrame: string,
  signal: AbortSignal,
  timedOut: () => boolean,
): Promise<void> {
  response.writeHead(200, {
    'cache-control': 'no-cache, no-store',
    connection: 'keep-alive',
    'content-type': 'text/event-stream; charset=utf-8',
    'x-content-type-options': 'nosniff',
  })
  const responseId = responseIdFromCreatedFrame(firstFrame)
  try {
    let next: IteratorResult<string> = { done: false, value: firstFrame }
    while (!next.done) {
      if (signal.aborted) throw new RequestReadError()
      if (!response.write(next.value)) await waitForDrain(response, signal)
      next = await iterator.next()
    }
    response.end()
  } catch {
    cancelBody(upstream)
    void iterator.return?.().catch(() => undefined)
    if (response.destroyed || response.writableEnded) return
    const timeout = timedOut()
    try {
      response.write(
        terminalFailureFrame(
          responseId,
          timeout ? 'turn_timeout' : 'upstream_error',
          timeout ? 'Turn timed out' : 'Upstream stream failed',
        ),
      )
      response.end('data: [DONE]\n\n')
    } catch {
      response.destroy()
    }
  }
}

function validateMaxBodyBytes(value: number | undefined): number {
  const resolved = value ?? DEFAULT_MAX_BODY_BYTES
  if (!Number.isSafeInteger(resolved) || resolved <= 0 || resolved > MAX_BODY_BYTES) {
    throw new TypeError('invalid_max_body_bytes')
  }
  return resolved
}

function validateDownwardLimit(value: number | undefined, maximum: number, code: string): number {
  const resolved = value ?? maximum
  if (!Number.isSafeInteger(resolved) || resolved <= 0 || resolved > maximum) {
    throw new TypeError(code)
  }
  return resolved
}

export async function startResponsesBridge(
  options: ResponsesBridgeOptions,
): Promise<ResponsesBridge> {
  if (!options || typeof options.fetchWithAuth !== 'function') {
    throw new TypeError('invalid_bridge_options')
  }
  const maxBodyBytes = validateMaxBodyBytes(options.maxBodyBytes)
  const maxActiveTurns = validateDownwardLimit(
    options.maxActiveTurns,
    MAX_ACTIVE_TURNS,
    'invalid_max_active_turns',
  )
  const maxTurnDurationMs = validateDownwardLimit(
    options.maxTurnDurationMs,
    MAX_TURN_DURATION_MS,
    'invalid_max_turn_duration',
  )
  const secret = createBridgeSecret()
  const controllers = new Set<AbortController>()
  const sockets = new Set<Socket>()
  let closing = false
  let activeTurns = 0

  const server = http.createServer((request, response) => {
    const controller = new AbortController()
    controllers.add(controller)
    let ownsTurnSlot = false
    let timedOut = false
    let turnTimer: ReturnType<typeof setTimeout> | undefined
    const abort = (): void => controller.abort()
    const onResponseClose = (): void => {
      if (!response.writableEnded) abort()
    }
    request.once('aborted', abort)
    response.once('close', onResponseClose)
    response.setTimeout(RESPONSE_IDLE_TIMEOUT_MS, () => {
      abort()
      response.destroy()
    })

    void (async () => {
      try {
        if (closing) {
          jsonError(response, 503, 'bridge_closed', 'Bridge unavailable')
          return
        }
        if (!hasValidBearerToken(request.headers.authorization, secret)) {
          request.resume()
          jsonError(response, 401, 'unauthorized', 'Unauthorized', {
            'www-authenticate': 'Bearer',
          })
          return
        }
        if (request.url !== RESPONSES_PATH) {
          request.resume()
          jsonError(response, 404, 'not_found', 'Not found')
          return
        }
        if (request.method !== 'POST') {
          request.resume()
          jsonError(response, 405, 'method_not_allowed', 'Method not allowed', { allow: 'POST' })
          return
        }
        if (!isJsonContentType(request.headers['content-type'])) {
          request.resume()
          jsonError(response, 415, 'unsupported_media_type', 'Unsupported media type')
          return
        }
        if (activeTurns >= maxActiveTurns) {
          request.resume()
          jsonError(response, 429, 'bridge_busy', 'Bridge is busy')
          return
        }
        activeTurns += 1
        ownsTurnSlot = true
        turnTimer = setTimeout(() => {
          timedOut = true
          controller.abort()
        }, maxTurnDurationMs)
        turnTimer.unref()

        let body: Buffer
        try {
          body = await readRawBody(request, maxBodyBytes, controller.signal)
        } catch (error) {
          if (error instanceof BodyTooLargeError) {
            jsonError(response, 413, 'request_too_large', 'Request too large')
          } else if (timedOut) {
            jsonError(response, 504, 'turn_timeout', 'Turn timed out')
          } else if (!controller.signal.aborted) {
            jsonError(response, 400, 'invalid_request', 'Invalid request')
          }
          return
        }

        let input: unknown
        try {
          input = parseJsonBody(body)
        } catch {
          jsonError(response, 400, 'invalid_request', 'Invalid request')
          return
        }

        let turn
        try {
          turn = prepareResponsesTurn(input)
        } catch (error) {
          const code = error instanceof ProtocolCompatibilityError ? error.code : 'invalid_request'
          jsonError(response, 400, code, 'Invalid request')
          return
        }
        if (controller.signal.aborted) return

        let upstream: Response
        try {
          const pendingUpstream = options.fetchWithAuth(turn.messagesRequest, controller.signal)
          upstream = await awaitAbortableUpstream(pendingUpstream, controller.signal)
        } catch (error) {
          if (timedOut) {
            jsonError(response, 504, 'turn_timeout', 'Turn timed out')
            return
          }
          if (controller.signal.aborted) return
          if (isAuthRequired(error)) {
            jsonError(response, 401, 'auth_required', 'Authentication required')
          } else {
            jsonError(response, 502, 'upstream_error', 'Upstream request failed')
          }
          return
        }
        if (controller.signal.aborted) {
          cancelBody(upstream)
          if (timedOut) jsonError(response, 504, 'turn_timeout', 'Turn timed out')
          return
        }
        if (upstream.status === 401) {
          cancelBody(upstream)
          jsonError(response, 401, 'auth_required', 'Authentication required')
          return
        }
        if (!upstream.ok || upstream.body === null) {
          cancelBody(upstream)
          jsonError(response, 502, 'upstream_error', 'Upstream request failed')
          return
        }
        if (!isEventStreamContentType(upstream.headers.get('content-type'))) {
          cancelBody(upstream)
          jsonError(response, 502, 'upstream_error', 'Upstream request failed')
          return
        }
        const convertedFrames = turn.messagesStreamToResponses(
          responseChunks(upstream.body, controller.signal),
        )
        const iterator = convertedFrames[Symbol.asyncIterator]()
        let firstFrame: IteratorResult<string>
        try {
          firstFrame = await iterator.next()
        } catch {
          cancelBody(upstream)
          jsonError(
            response,
            timedOut ? 504 : 502,
            timedOut ? 'turn_timeout' : 'upstream_error',
            timedOut ? 'Turn timed out' : 'Upstream request failed',
          )
          return
        }
        if (firstFrame.done) {
          cancelBody(upstream)
          jsonError(response, 502, 'upstream_error', 'Upstream request failed')
          return
        }
        await streamResponse(
          response,
          upstream,
          iterator,
          firstFrame.value,
          controller.signal,
          () => timedOut,
        )
      } catch {
        controller.abort()
        if (response.headersSent) {
          if (!response.destroyed) response.destroy()
        } else {
          jsonError(response, 500, 'bridge_error', 'Bridge request failed')
        }
      } finally {
        if (turnTimer !== undefined) clearTimeout(turnTimer)
        if (ownsTurnSlot) activeTurns -= 1
        request.off('aborted', abort)
        response.off('close', onResponseClose)
        controllers.delete(controller)
      }
    })()
  })

  server.headersTimeout = HEADERS_TIMEOUT_MS
  server.requestTimeout = REQUEST_TIMEOUT_MS
  server.keepAliveTimeout = KEEP_ALIVE_TIMEOUT_MS
  server.maxRequestsPerSocket = 100
  server.on('connection', (socket) => {
    sockets.add(socket)
    socket.once('close', () => sockets.delete(socket))
  })
  server.on('error', () => {
    closing = true
    for (const controller of controllers) controller.abort()
    for (const socket of sockets) socket.destroy()
    if (server.listening) server.close()
  })

  await new Promise<void>((resolve, reject) => {
    const onError = (error: Error): void => {
      server.off('listening', onListening)
      reject(error)
    }
    const onListening = (): void => {
      server.off('error', onError)
      resolve()
    }
    server.once('error', onError)
    server.once('listening', onListening)
    server.listen(0, LOOPBACK_HOST)
  })

  const address = server.address() as AddressInfo
  const baseUrl = `http://${LOOPBACK_HOST}:${address.port}`
  let closePromise: Promise<void> | undefined
  const close = (): Promise<void> => {
    if (closePromise) return closePromise
    closing = true
    for (const controller of controllers) controller.abort()
    closePromise = new Promise((resolve) => {
      server.close(() => resolve())
      server.closeIdleConnections()
      for (const socket of sockets) socket.destroy()
    })
    return closePromise
  }

  return Object.freeze({
    baseUrl,
    responsesUrl: `${baseUrl}${RESPONSES_PATH}`,
    secret,
    close,
  })
}
