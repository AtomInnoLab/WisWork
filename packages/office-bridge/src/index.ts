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

async function readBoundedJson(request: Request, maximum: number): Promise<unknown> {
  const declared = Number(request.headers.get('content-length'))
  if (Number.isFinite(declared) && declared > maximum) throw new Error('body_too_large')
  const bytes = new Uint8Array(await request.arrayBuffer())
  if (bytes.byteLength > maximum) throw new Error('body_too_large')
  try {
    return JSON.parse(new TextDecoder().decode(bytes)) as unknown
  } catch {
    throw new Error('invalid_json')
  }
}

function isAsyncIterable(
  value: Uint8Array | AsyncIterable<Uint8Array>,
): value is AsyncIterable<Uint8Array> {
  return Symbol.asyncIterator in value
}

async function collectBounded(
  body: Uint8Array | AsyncIterable<Uint8Array>,
  maximum: number,
): Promise<Uint8Array> {
  if (!isAsyncIterable(body)) {
    if (body.byteLength > maximum) throw new Error('response_too_large')
    return body
  }
  const chunks: Uint8Array[] = []
  let size = 0
  for await (const chunk of body) {
    size += chunk.byteLength
    if (size > maximum) throw new Error('response_too_large')
    chunks.push(chunk)
  }
  const result = new Uint8Array(size)
  let offset = 0
  for (const chunk of chunks) {
    result.set(chunk, offset)
    offset += chunk.byteLength
  }
  return result
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
  const now = options.now ?? Date.now
  const pairingTtlMs = options.pairingTtlMs ?? 120_000
  const capabilityTtlMs = options.capabilityTtlMs ?? 15 * 60_000
  const maxPairings = options.maxPairings ?? 20
  const maxCapabilities = options.maxCapabilities ?? 50
  const maxCreatesPerMinute = options.maxPairingCreatesPerMinute ?? 30
  const maxBodyBytes = options.maxBodyBytes ?? 256 * 1024
  const maxResponseBytes = options.maxResponseBytes ?? 4 * 1024 * 1024
  const maxConcurrentMessages = options.maxConcurrentMessages ?? 2
  const messageTimeoutMs = options.messageTimeoutMs ?? 120_000
  const pairings = new Map<string, Pairing>()
  let capabilities: Capability[] = []
  let createTimes: number[] = []
  let activeMessages = 0
  let stopped = false

  const findCapability = (candidate: string): Capability | undefined => {
    const timestamp = now()
    capabilities = capabilities.filter((entry) => entry.expiresAt > timestamp)
    return capabilities.find(
      (entry) => entry.expiresAt > timestamp && constantTimeMatches(entry.hash, candidate),
    )
  }

  const handlePreflight = (request: Request): Response => {
    const headers = new Headers()
    addCors(headers, options.allowedOrigin)
    headers.set('access-control-allow-methods', 'GET, POST, OPTIONS')
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
    if (createTimes.length >= maxCreatesPerMinute || pairings.size >= maxPairings) {
      return jsonResponse(429, { error: 'pairing_capacity' }, options.allowedOrigin)
    }
    let parsed: unknown
    try {
      parsed = await readBoundedJson(request, 4096)
    } catch (error) {
      const tooLarge = error instanceof Error && error.message === 'body_too_large'
      return jsonResponse(
        tooLarge ? 413 : 400,
        {
          error: tooLarge ? 'body_too_large' : 'invalid_request',
        },
        options.allowedOrigin,
      )
    }
    const hostLabel =
      typeof parsed === 'object' &&
      parsed !== null &&
      typeof (parsed as Record<string, unknown>).host_label === 'string'
        ? (parsed as Record<string, string>).host_label.trim().slice(0, 80)
        : ''
    if (!hostLabel) return jsonResponse(400, { error: 'invalid_request' }, options.allowedOrigin)
    const id = opaqueValue()
    const pollingSecret = opaqueValue()
    pairings.set(id, {
      id,
      pollingSecretHash: digest(pollingSecret),
      hostLabel,
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
    let body: unknown
    try {
      body = await readBoundedJson(request, maxBodyBytes)
    } catch (error) {
      const tooLarge = error instanceof Error && error.message === 'body_too_large'
      return jsonResponse(
        tooLarge ? 413 : 400,
        {
          error: tooLarge ? 'body_too_large' : 'invalid_request',
        },
        options.allowedOrigin,
      )
    }
    activeMessages += 1
    const controller = new AbortController()
    const abort = () => controller.abort()
    request.signal.addEventListener('abort', abort, { once: true })
    let timer: ReturnType<typeof setTimeout> | undefined
    try {
      const timeout = new Promise<never>((_, reject) => {
        timer = setTimeout(() => {
          controller.abort()
          reject(new Error('timeout'))
        }, messageTimeoutMs)
      })
      const upstream = await Promise.race([
        options.proxy({ body, signal: controller.signal }),
        timeout,
      ])
      if (upstream.status < 200 || upstream.status >= 300) throw new Error('upstream_status')
      const bytes = await Promise.race([collectBounded(upstream.body, maxResponseBytes), timeout])
      const headers = new Headers({
        'content-type': upstream.contentType ?? 'application/octet-stream',
      })
      addCors(headers, options.allowedOrigin)
      return new Response(Uint8Array.from(bytes).buffer, { status: 200, headers })
    } catch {
      return jsonResponse(502, { error: 'upstream_failed' }, options.allowedOrigin)
    } finally {
      if (timer) clearTimeout(timer)
      request.signal.removeEventListener('abort', abort)
      activeMessages -= 1
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
    },
    shutdown() {
      stopped = true
      pairings.clear()
      capabilities = []
    },
  }
}
