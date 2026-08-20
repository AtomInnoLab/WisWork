import { createHash, randomBytes, timingSafeEqual } from 'node:crypto'

export interface MessagesProxyRequest {
  body: unknown
  signal: AbortSignal
}

export interface MessagesProxyResponse {
  status: number
  contentType?: string
  body: Uint8Array | AsyncIterable<Uint8Array>
}

export type MessagesProxy = (request: MessagesProxyRequest) => Promise<MessagesProxyResponse>

export interface OfficeBridgeOptions {
  allowedOrigin: string
  proxy: MessagesProxy
  now?: () => number
  pairingTtlMs?: number
  capabilityTtlMs?: number
  maxPairings?: number
  maxCapabilities?: number
  maxPairingCreatesPerMinute?: number
  maxConcurrentPairingCreates?: number
  pairingRequestTimeoutMs?: number
  maxBodyBytes?: number
  maxResponseBytes?: number
  maxConcurrentMessages?: number
  messageTimeoutMs?: number
}

interface Pairing {
  id: string
  pollingSecretHash: Buffer
  hostLabel: string
  origin: string
  expiresAt: number
  status: 'pending' | 'approved' | 'rejected'
}

interface Capability {
  hash: Buffer
  expiresAt: number
}

export interface PendingPairing {
  pairingId: string
  hostLabel: string
  origin: string
}

export interface OfficeBridge {
  handle(request: Request): Promise<Response>
  listPending(): PendingPairing[]
  approve(pairingId: string, hasValidSession: boolean): boolean
  reject(pairingId: string): boolean
  revokeAll(): void
  shutdown(): void
}

export function assertLoopbackHost(host: string): '127.0.0.1' {
  if (host !== '127.0.0.1') throw new Error('loopback_host_required')
  return host
}

function opaqueValue(): string {
  return randomBytes(32).toString('base64url')
}

function digest(value: string): Buffer {
  return createHash('sha256').update(value, 'utf8').digest()
}

function constantTimeMatches(stored: Buffer, candidate: string): boolean {
  return timingSafeEqual(stored, digest(candidate))
}

function jsonResponse(status: number, body: unknown, origin?: string): Response {
  const headers = new Headers({ 'content-type': 'application/json; charset=utf-8' })
  if (origin) addCors(headers, origin)
  return new Response(JSON.stringify(body), { status, headers })
}

function addCors(headers: Headers, origin: string): void {
  headers.set('access-control-allow-origin', origin)
  headers.set('vary', 'Origin')
}

function authorization(request: Request, scheme: string): string | null {
  const value = request.headers.get('authorization')
  const prefix = `${scheme} `
  return value?.startsWith(prefix) ? value.slice(prefix.length) : null
}

async function readBoundedJson(
  request: Request,
  maximum: number,
  signal: AbortSignal = request.signal,
): Promise<unknown> {
  const declared = Number(request.headers.get('content-length'))
  if (Number.isFinite(declared) && declared > maximum) throw new Error('body_too_large')
  if (!request.body) throw new Error('invalid_json')
  const reader = request.body.getReader()
  const chunks: Uint8Array[] = []
  let size = 0
  let rejectAbort!: (error: Error) => void
  const aborted = new Promise<never>((_, reject) => (rejectAbort = reject))
  const onAbort = () => {
    void reader.cancel('request_aborted').catch(() => undefined)
    rejectAbort(new Error('request_aborted'))
  }
  signal.addEventListener('abort', onAbort, { once: true })
  if (signal.aborted) onAbort()
  try {
    while (true) {
      const result = await Promise.race([reader.read(), aborted])
      if (result.done) break
      size += result.value.byteLength
      if (size > maximum) {
        void reader.cancel('body_too_large').catch(() => undefined)
        throw new Error('body_too_large')
      }
      chunks.push(result.value)
    }
    const bytes = new Uint8Array(size)
    let offset = 0
    for (const chunk of chunks) {
      bytes.set(chunk, offset)
      offset += chunk.byteLength
    }
    return JSON.parse(new TextDecoder().decode(bytes)) as unknown
  } catch (error) {
    if (error instanceof Error && error.message === 'body_too_large') throw error
    void reader.cancel('invalid_request').catch(() => undefined)
    throw new Error('invalid_json', { cause: error })
  } finally {
    signal.removeEventListener('abort', onAbort)
  }
}

function isAsyncIterable(
  value: Uint8Array | AsyncIterable<Uint8Array>,
): value is AsyncIterable<Uint8Array> {
  return Symbol.asyncIterator in value
}

function bodyIterator(body: Uint8Array | AsyncIterable<Uint8Array>): AsyncIterator<Uint8Array> {
  if (isAsyncIterable(body)) return body[Symbol.asyncIterator]()
  let emitted = false
  return {
    async next() {
      if (emitted) return { done: true, value: undefined }
      emitted = true
      return { done: false, value: body }
    },
    async return() {
      return { done: true, value: undefined }
    },
  }
}

interface ActiveOperation {
  abort(): void
}

export function createOfficeBridge(options: OfficeBridgeOptions): OfficeBridge {
  let allowedOrigin: URL
  try {
    allowedOrigin = new URL(options.allowedOrigin)
  } catch {
    throw new Error('invalid_allowed_origin')
  }
  if (allowedOrigin.protocol !== 'https:' || allowedOrigin.origin !== options.allowedOrigin) {
    throw new Error('invalid_allowed_origin')
  }
  const configuredLimits = [
    options.pairingTtlMs,
    options.capabilityTtlMs,
    options.maxPairings,
    options.maxCapabilities,
    options.maxPairingCreatesPerMinute,
    options.maxConcurrentPairingCreates,
    options.pairingRequestTimeoutMs,
    options.maxBodyBytes,
    options.maxResponseBytes,
    options.maxConcurrentMessages,
    options.messageTimeoutMs,
  ].filter((value): value is number => value !== undefined)
  if (configuredLimits.some((value) => !Number.isSafeInteger(value) || value <= 0)) {
    throw new Error('invalid_bridge_limits')
  }
  const now = options.now ?? Date.now
  const pairingTtlMs = options.pairingTtlMs ?? 120_000
  const capabilityTtlMs = options.capabilityTtlMs ?? 15 * 60_000
  const maxPairings = options.maxPairings ?? 20
  const maxCapabilities = options.maxCapabilities ?? 50
  const maxCreatesPerMinute = options.maxPairingCreatesPerMinute ?? 30
  const maxConcurrentPairingCreates = options.maxConcurrentPairingCreates ?? 4
  const pairingRequestTimeoutMs = options.pairingRequestTimeoutMs ?? 5_000
  const maxBodyBytes = options.maxBodyBytes ?? 256 * 1024
  const maxResponseBytes = options.maxResponseBytes ?? 4 * 1024 * 1024
  const maxConcurrentMessages = options.maxConcurrentMessages ?? 2
  const messageTimeoutMs = options.messageTimeoutMs ?? 120_000
  const pairings = new Map<string, Pairing>()
  let capabilities: Capability[] = []
  let createTimes: number[] = []
  let activePairingCreates = 0
  const activePairingControllers = new Set<AbortController>()
  let activeMessages = 0
  const activeOperations = new Set<ActiveOperation>()
  let stopped = false

  const findCapability = (candidate: string): Capability | undefined => {
    const timestamp = now()
    capabilities = capabilities.filter((entry) => entry.expiresAt > timestamp)
    return capabilities.find(
      (entry) => entry.expiresAt > timestamp && constantTimeMatches(entry.hash, candidate),
    )
  }

  const handlePreflight = (request: Request): Response => {
    const path = new URL(request.url).pathname
    const pairingPoll = /^\/v1\/office\/pairings\/[A-Za-z0-9_-]+$/.test(path)
    const expectedMethod = pairingPoll
      ? 'GET'
      : path === '/v1/office/pairings' || path === '/v1/office/messages'
        ? 'POST'
        : null
    const requestedMethod = request.headers.get('access-control-request-method')?.toUpperCase()
    const requestedHeaders = (request.headers.get('access-control-request-headers') ?? '')
      .split(',')
      .map((header) => header.trim().toLowerCase())
      .filter(Boolean)
    const allowedHeaders = new Set(['authorization', 'content-type'])
    if (
      !expectedMethod ||
      requestedMethod !== expectedMethod ||
      requestedHeaders.some((header) => !allowedHeaders.has(header))
    ) {
      return jsonResponse(403, { error: 'preflight_denied' })
    }
    const headers = new Headers()
    addCors(headers, options.allowedOrigin)
    headers.set('access-control-allow-methods', expectedMethod)
    headers.set('access-control-allow-headers', 'Authorization, Content-Type')
    headers.append('vary', 'Access-Control-Request-Headers')
    if (request.headers.get('access-control-request-private-network') === 'true') {
      headers.set('access-control-allow-private-network', 'true')
      headers.append('vary', 'Access-Control-Request-Private-Network')
    }
    return new Response(null, { status: 204, headers })
  }

  const createPairing = async (request: Request): Promise<Response> => {
    if (request.headers.get('content-type')?.split(';', 1)[0] !== 'application/json') {
      return jsonResponse(415, { error: 'unsupported_media_type' }, options.allowedOrigin)
    }
    const timestamp = now()
    for (const [id, pairing] of pairings) {
      if (pairing.expiresAt <= timestamp) pairings.delete(id)
    }
    createTimes = createTimes.filter((entry) => timestamp - entry < 60_000)
    if (
      createTimes.length >= maxCreatesPerMinute ||
      pairings.size >= maxPairings ||
      activePairingCreates >= maxConcurrentPairingCreates
    ) {
      return jsonResponse(429, { error: 'pairing_capacity' }, options.allowedOrigin)
    }
    activePairingCreates += 1
    const bodyController = new AbortController()
    activePairingControllers.add(bodyController)
    const abortBody = () => bodyController.abort()
    request.signal.addEventListener('abort', abortBody, { once: true })
    const bodyTimer = setTimeout(() => bodyController.abort(), pairingRequestTimeoutMs)
    if (request.signal.aborted) bodyController.abort()
    let parsed: unknown
    try {
      parsed = await readBoundedJson(request, 4096, bodyController.signal)
    } catch (error) {
      const tooLarge = error instanceof Error && error.message === 'body_too_large'
      const interrupted = bodyController.signal.aborted
      return jsonResponse(
        tooLarge ? 413 : interrupted ? 408 : 400,
        {
          error: tooLarge ? 'body_too_large' : interrupted ? 'request_timeout' : 'invalid_request',
        },
        options.allowedOrigin,
      )
    } finally {
      clearTimeout(bodyTimer)
      request.signal.removeEventListener('abort', abortBody)
      activePairingControllers.delete(bodyController)
      activePairingCreates -= 1
    }
    const requestedHostLabel =
      typeof parsed === 'object' &&
      parsed !== null &&
      typeof (parsed as Record<string, unknown>).host_label === 'string'
        ? (parsed as Record<string, string>).host_label
        : ''
    const allowedHostLabels = new Set(['Word', 'Excel', 'PowerPoint'])
    if (!allowedHostLabels.has(requestedHostLabel)) {
      return jsonResponse(400, { error: 'invalid_request' }, options.allowedOrigin)
    }
    const id = opaqueValue()
    const pollingSecret = opaqueValue()
    pairings.set(id, {
      id,
      pollingSecretHash: digest(pollingSecret),
      hostLabel: requestedHostLabel,
      origin: options.allowedOrigin,
      expiresAt: timestamp + pairingTtlMs,
      status: 'pending',
    })
    createTimes.push(timestamp)
    return jsonResponse(
      202,
      { pairing_id: id, polling_secret: pollingSecret, expires_in: Math.ceil(pairingTtlMs / 1000) },
      options.allowedOrigin,
    )
  }

  const pollPairing = (request: Request, id: string): Response => {
    const entry = pairings.get(id)
    const secret = authorization(request, 'Pairing')
    if (!entry || !secret || !constantTimeMatches(entry.pollingSecretHash, secret)) {
      return jsonResponse(401, { error: 'invalid_pairing' }, options.allowedOrigin)
    }
    if (entry.expiresAt <= now()) {
      pairings.delete(id)
      return jsonResponse(200, { status: 'expired' }, options.allowedOrigin)
    }
    if (entry.status === 'pending') {
      return jsonResponse(200, { status: 'pending' }, options.allowedOrigin)
    }
    if (entry.status === 'rejected') {
      pairings.delete(id)
      return jsonResponse(200, { status: 'rejected' }, options.allowedOrigin)
    }
    capabilities = capabilities.filter((capability) => capability.expiresAt > now())
    if (capabilities.length >= maxCapabilities) {
      pairings.delete(id)
      return jsonResponse(429, { error: 'capability_capacity' }, options.allowedOrigin)
    }
    const capability = opaqueValue()
    capabilities.push({ hash: digest(capability), expiresAt: now() + capabilityTtlMs })
    pairings.delete(id)
    return jsonResponse(
      200,
      { status: 'approved', capability, expires_in: Math.ceil(capabilityTtlMs / 1000) },
      options.allowedOrigin,
    )
  }

  const proxyMessages = async (request: Request): Promise<Response> => {
    const secret = authorization(request, 'Bridge')
    if (!secret || !findCapability(secret)) {
      return jsonResponse(401, { error: 'invalid_capability' }, options.allowedOrigin)
    }
    if (request.headers.get('content-type')?.split(';', 1)[0] !== 'application/json') {
      return jsonResponse(415, { error: 'unsupported_media_type' }, options.allowedOrigin)
    }
    if (activeMessages >= maxConcurrentMessages) {
      return jsonResponse(429, { error: 'message_capacity' }, options.allowedOrigin)
    }
    activeMessages += 1
    const controller = new AbortController()
    let iterator: AsyncIterator<Uint8Array> | undefined
    let streamController: ReadableStreamDefaultController<Uint8Array> | undefined
    let finished = false
    let rejectInterrupted!: (error: Error) => void
    const interrupted = new Promise<never>((_, reject) => (rejectInterrupted = reject))
    void interrupted.catch(() => undefined)
    const timer = setTimeout(() => operation.abort(), messageTimeoutMs)
    const finish = (abort: boolean, errorStream = false): void => {
      if (finished) return
      finished = true
      clearTimeout(timer)
      request.signal.removeEventListener('abort', abortFromClient)
      activeOperations.delete(operation)
      activeMessages -= 1
      if (abort) controller.abort()
      rejectInterrupted(new Error('operation_interrupted'))
      if (iterator?.return) void iterator.return().catch(() => undefined)
      if (errorStream && streamController) {
        try {
          streamController.error(new Error('stream_interrupted'))
        } catch {
          // The consumer may already have closed or cancelled the stream.
        }
      }
    }
    const operation: ActiveOperation = { abort: () => finish(true, true) }
    const abortFromClient = () => operation.abort()
    activeOperations.add(operation)
    request.signal.addEventListener('abort', abortFromClient, { once: true })
    if (request.signal.aborted) operation.abort()
    let body: unknown
    try {
      body = await readBoundedJson(request, maxBodyBytes, controller.signal)
    } catch (error) {
      const tooLarge = error instanceof Error && error.message === 'body_too_large'
      const interruptedBody = controller.signal.aborted
      finish(true)
      return jsonResponse(
        tooLarge ? 413 : interruptedBody ? 408 : 400,
        {
          error: tooLarge
            ? 'body_too_large'
            : interruptedBody
              ? 'request_timeout'
              : 'invalid_request',
        },
        options.allowedOrigin,
      )
    }
    try {
      const proxyPromise = options.proxy({ body, signal: controller.signal })
      void proxyPromise.then(
        (lateResponse) => {
          if (finished && isAsyncIterable(lateResponse.body)) {
            void lateResponse.body[Symbol.asyncIterator]()
              .return?.()
              .catch(() => undefined)
          }
        },
        () => undefined,
      )
      const upstream = await Promise.race([proxyPromise, interrupted])
      if (upstream.status < 200 || upstream.status >= 300) throw new Error('upstream_status')
      iterator = bodyIterator(upstream.body)
      let streamedBytes = 0
      const stream = new ReadableStream<Uint8Array>({
        start(readableController) {
          streamController = readableController
        },
        async pull(readableController) {
          try {
            const result = await Promise.race([iterator!.next(), interrupted])
            if (result.done) {
              finish(false)
              readableController.close()
              return
            }
            if (!(result.value instanceof Uint8Array)) throw new Error('invalid_stream_chunk')
            streamedBytes += result.value.byteLength
            if (streamedBytes > maxResponseBytes) throw new Error('response_too_large')
            readableController.enqueue(result.value)
          } catch {
            finish(true)
            try {
              readableController.error(new Error('stream_failed'))
            } catch {
              // The stream was already closed by a concurrent abort.
            }
          }
        },
        cancel() {
          finish(true)
        },
      })
      const headers = new Headers({
        'content-type': upstream.contentType ?? 'application/octet-stream',
      })
      addCors(headers, options.allowedOrigin)
      return new Response(stream, { status: 200, headers })
    } catch {
      finish(true)
      return jsonResponse(502, { error: 'upstream_failed' }, options.allowedOrigin)
    }
  }

  return {
    async handle(request) {
      if (stopped) return jsonResponse(503, { error: 'bridge_unavailable' })
      if (request.headers.get('origin') !== options.allowedOrigin) {
        return jsonResponse(403, { error: 'origin_denied' })
      }
      if (request.method === 'OPTIONS') return handlePreflight(request)
      const path = new URL(request.url).pathname
      if (request.method === 'POST' && path === '/v1/office/pairings') {
        return createPairing(request)
      }
      const pairingMatch = /^\/v1\/office\/pairings\/([A-Za-z0-9_-]+)$/.exec(path)
      if (request.method === 'GET' && pairingMatch) return pollPairing(request, pairingMatch[1])
      if (request.method === 'POST' && path === '/v1/office/messages') {
        return proxyMessages(request)
      }
      if (path === '/v1/office/pairings' || path === '/v1/office/messages' || pairingMatch) {
        return jsonResponse(405, { error: 'method_not_allowed' }, options.allowedOrigin)
      }
      return jsonResponse(404, { error: 'not_found' }, options.allowedOrigin)
    },
    listPending() {
      const timestamp = now()
      return [...pairings.values()]
        .filter((entry) => entry.status === 'pending' && entry.expiresAt > timestamp)
        .map((entry) => ({
          pairingId: entry.id,
          hostLabel: entry.hostLabel,
          origin: entry.origin,
        }))
    },
    approve(pairingId, hasValidSession) {
      const entry = pairings.get(pairingId)
      if (!hasValidSession || !entry || entry.status !== 'pending' || entry.expiresAt <= now()) {
        return false
      }
      entry.status = 'approved'
      return true
    },
    reject(pairingId) {
      const entry = pairings.get(pairingId)
      if (!entry || entry.status !== 'pending' || entry.expiresAt <= now()) return false
      entry.status = 'rejected'
      return true
    },
    revokeAll() {
      pairings.clear()
      capabilities = []
      for (const controller of [...activePairingControllers]) controller.abort()
      for (const operation of [...activeOperations]) operation.abort()
    },
    shutdown() {
      stopped = true
      pairings.clear()
      capabilities = []
      for (const controller of [...activePairingControllers]) controller.abort()
      for (const operation of [...activeOperations]) operation.abort()
    },
  }
}
