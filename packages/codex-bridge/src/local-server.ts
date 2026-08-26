import http, { type IncomingMessage, type ServerResponse } from 'node:http'
import type { AddressInfo, Socket } from 'node:net'
import { TextDecoder } from 'node:util'
import { prepareResponsesTurn, ProtocolCompatibilityError } from './index.js'
import { createBridgeSecret, hasValidBearerToken } from './security.js'
import type { MessagesRequest, ProtocolLimits } from './types.js'

const LOOPBACK_HOST = '127.0.0.1'
const RESPONSES_PATH = '/v1/responses'
const DEFAULT_MAX_BODY_BYTES = 8_000_000
const MAX_BODY_BYTES = 8_000_000
const HEADERS_TIMEOUT_MS = 10_000
const REQUEST_TIMEOUT_MS = 30_000
const RESPONSE_IDLE_TIMEOUT_MS = 30_000
const KEEP_ALIVE_TIMEOUT_MS = 5_000

export interface ResponsesBridgeOptions {
  fetchWithAuth: (request: MessagesRequest, signal: AbortSignal) => Promise<Response>
  maxBodyBytes?: number
  protocolLimits?: Partial<ProtocolLimits>
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

function isJsonContentType(value: string | undefined): boolean {
  if (value === undefined) return false
  const parts = value
    .split(';')
    .map((part) => part.trim().toLowerCase())
    .filter(Boolean)
  return (
    parts[0] === 'application/json' &&
    parts.length <= 2 &&
    (parts.length === 1 || parts[1] === 'charset=utf-8')
  )
}

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

async function cancelBody(response: Response): Promise<void> {
  try {
    await response.body?.cancel()
  } catch {
    // Cancellation is best-effort after a redacted status response.
  }
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
    while (!signal.aborted) {
      const chunk = await reader.read()
      if (chunk.done) {
        ended = true
        return
      }
      yield chunk.value
    }
  } finally {
    signal.removeEventListener('abort', cancel)
    if (!ended) await reader.cancel().catch(() => undefined)
    reader.releaseLock()
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

async function streamResponse(
  response: ServerResponse,
  upstream: Response,
  frames: AsyncIterable<string>,
  signal: AbortSignal,
): Promise<void> {
  response.writeHead(200, {
    'cache-control': 'no-cache, no-store',
    connection: 'keep-alive',
    'content-type': 'text/event-stream; charset=utf-8',
    'x-content-type-options': 'nosniff',
  })
  try {
    for await (const frame of frames) {
      if (signal.aborted) throw new RequestReadError()
      if (!response.write(frame)) await waitForDrain(response, signal)
    }
    response.end()
  } catch {
    await cancelBody(upstream)
    if (!response.destroyed) response.destroy()
  }
}

function validateMaxBodyBytes(value: number | undefined): number {
  const resolved = value ?? DEFAULT_MAX_BODY_BYTES
  if (!Number.isSafeInteger(resolved) || resolved <= 0 || resolved > MAX_BODY_BYTES) {
    throw new TypeError('invalid_max_body_bytes')
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
  const secret = createBridgeSecret()
  const controllers = new Set<AbortController>()
  const sockets = new Set<Socket>()
  let closing = false

  const server = http.createServer((request, response) => {
    const controller = new AbortController()
    controllers.add(controller)
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

        let body: Buffer
        try {
          body = await readRawBody(request, maxBodyBytes, controller.signal)
        } catch (error) {
          if (error instanceof BodyTooLargeError) {
            jsonError(response, 413, 'request_too_large', 'Request too large')
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
          turn = prepareResponsesTurn(input, options.protocolLimits)
        } catch (error) {
          const code = error instanceof ProtocolCompatibilityError ? error.code : 'invalid_request'
          jsonError(response, 400, code, 'Invalid request')
          return
        }
        if (controller.signal.aborted) return

        let upstream: Response
        try {
          upstream = await options.fetchWithAuth(turn.messagesRequest, controller.signal)
        } catch (error) {
          if (controller.signal.aborted) return
          if (isAuthRequired(error)) {
            jsonError(response, 401, 'auth_required', 'Authentication required')
          } else {
            jsonError(response, 502, 'upstream_error', 'Upstream request failed')
          }
          return
        }
        if (controller.signal.aborted) {
          await cancelBody(upstream)
          return
        }
        if (upstream.status === 401) {
          await cancelBody(upstream)
          jsonError(response, 401, 'auth_required', 'Authentication required')
          return
        }
        if (!upstream.ok || upstream.body === null) {
          await cancelBody(upstream)
          jsonError(response, 502, 'upstream_error', 'Upstream request failed')
          return
        }
        await streamResponse(
          response,
          upstream,
          turn.messagesStreamToResponses(responseChunks(upstream.body, controller.signal)),
          controller.signal,
        )
      } catch {
        controller.abort()
        if (response.headersSent) {
          if (!response.destroyed) response.destroy()
        } else {
          jsonError(response, 500, 'bridge_error', 'Bridge request failed')
        }
      } finally {
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
