import type { MessagesProxy } from '@wiswork/office-bridge'
import WebSocket from 'ws'
import type { OfficePairingRequest, OfficeRelayStatus } from '../shared/home-api'
import type { OfficeRetrievalProxy } from './office-retrieval-proxy'
export type { OfficeRelayStatus } from '../shared/home-api'

const MAX_CONTROL_BYTES = 16 * 1024
const MAX_REQUEST_BYTES = 256 * 1024
const MAX_CHUNK_BYTES = 64 * 1024
const MAX_RESPONSE_BYTES = 16 * 1024 * 1024
// Relay owns the 120s deadline; this only handles a lost relay.cancel.
const REQUEST_TIMEOUT_MS = 125_000
const CONNECT_TIMEOUT_MS = 10_000
// Relay owns the renewable idle TTL. PC keeps only a bounded absolute-lifetime watchdog.
const SESSION_ABSOLUTE_MAX_MS = 8 * 60 * 60 * 1_000
const IDENTIFIER = /^[A-Za-z0-9_-]{8,128}$/
const HOSTS = new Set(['Word', 'Excel', 'PowerPoint'])
const MAX_REQUEST_IDS = 2_048
const RELAY_ERROR_CODES = new Set([
  'already_claimed',
  'binary_not_supported',
  'chunk_too_large',
  'capability_not_negotiated',
  'claim_limit',
  'claim_rate_limited',
  'create_rate_limited',
  'duplicate_request',
  'frame_too_large',
  'invalid_capability',
  'invalid_code',
  'invalid_content_type',
  'invalid_frame',
  'invalid_pairing',
  'invalid_request',
  'invalid_sequence',
  'pairing_limit',
  'peer_unavailable',
  'relay_busy',
  'request_active',
  'request_limit',
  'request_timeout',
  'request_too_large',
  'response_too_large',
  'role_not_allowed',
  'session_expired',
  'session_revoked',
  'unknown_type',
  'unsupported_host',
])
const TERMINAL_REQUEST_CACHE_SIZE = 64
const PRODUCTION_RELAY_ENDPOINT = 'wss://office.8-216-134-194.sslip.io/office-relay'
const V2_CAPABILITIES = ['agent.v1', 'web-search.v1', 'web-fetch.v1', 'image-search.v1'] as const

export interface RelaySocket {
  readyState: number
  addEventListener(name: string, listener: (event: any) => void): void
  send(data: string): void
  close(code?: number, reason?: string): void
}

export interface OfficeRelayClient {
  claim(code: string): Promise<void>
  approve(pairingId: string): Promise<boolean>
  reject(pairingId: string): boolean
  listPending(): OfficePairingRequest[]
  status(): OfficeRelayStatus
  revoke(reason?: string): void
}

export function connectAuthenticatedRelaySocket(url: string, accessToken: string): RelaySocket {
  return new WebSocket(url, { headers: { Authorization: `Bearer ${accessToken}` } }) as RelaySocket
}

function exact(value: unknown, keys: readonly string[]): value is Record<string, unknown> {
  return Boolean(
    value &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    Object.keys(value).length === keys.length &&
    keys.every((key) => Object.prototype.hasOwnProperty.call(value, key)),
  )
}

function validId(value: unknown): value is string {
  return typeof value === 'string' && IDENTIFIER.test(value)
}

function jsonObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value))
}

export function officeRelayEndpointFromEnv(env: Record<string, string | undefined>): string {
  const value = env.WISWORK_OFFICE_RELAY_URL ?? PRODUCTION_RELAY_ENDPOINT
  const url = new URL(value)
  if (url.href !== PRODUCTION_RELAY_ENDPOINT) throw new Error('invalid_office_relay_url')
  return url.href
}

export function createOfficeRelayClient(options: {
  endpoint: string
  connect?: (url: string, accessToken: string) => RelaySocket
  getValidAccountStatus(): Promise<{ loggedIn: boolean }>
  getAccessToken(): Promise<string | null>
  proxy: MessagesProxy
  retrievalProxy?: OfficeRetrievalProxy
  negotiateCapabilities?: boolean
  onPending(pairing: OfficePairingRequest): void
  onPendingExpired?: (pairingId: string) => void
  onStatus?: (status: OfficeRelayStatus) => void
  maxRequestIds?: number
}): OfficeRelayClient {
  const connect = options.connect ?? connectAuthenticatedRelaySocket
  let socket: RelaySocket | null = null
  let diagnostic: OfficeRelayStatus = 'disconnected'
  let protocolVersion: 1 | 2 = 1
  const negotiateCapabilities =
    options.negotiateCapabilities === true || Boolean(options.retrievalProxy)
  const offeredCapabilities = options.retrievalProxy ? [...V2_CAPABILITIES] : ['agent.v1']
  let pending: (OfficePairingRequest & { capabilities?: string[] }) | null = null
  let session: { sessionId: string; capability: string; capabilities: string[] } | null = null
  let active: { requestId: string; controller: AbortController; remoteCancelled: boolean } | null =
    null
  let claimedCode: string | null = null
  let negotiationPending = false
  const requestIds = new Set<string>()
  const terminalRequestIds = new Set<string>()
  const terminalRequestOrder: string[] = []
  let generation = 0
  let approvalSentFor: string | null = null
  let pairingTimer: ReturnType<typeof setTimeout> | null = null
  let sessionTimer: ReturnType<typeof setTimeout> | null = null

  const setStatus = (value: OfficeRelayStatus) => {
    diagnostic = value
    options.onStatus?.(value)
  }
  const send = (frame: Record<string, unknown>) => {
    if (!socket || socket.readyState !== 1) throw new Error('relay_disconnected')
    const raw = JSON.stringify(frame)
    if (Buffer.byteLength(raw) > MAX_CONTROL_BYTES && frame.type !== 'pc.chunk')
      throw new Error('control_frame_too_large')
    socket.send(raw)
  }
  const clearTimers = () => {
    if (pairingTimer) clearTimeout(pairingTimer)
    if (sessionTimer) clearTimeout(sessionTimer)
    pairingTimer = null
    sessionTimer = null
  }
  const rememberTerminalRequest = (requestId: string) => {
    if (terminalRequestIds.has(requestId)) return
    terminalRequestIds.add(requestId)
    terminalRequestOrder.push(requestId)
    while (terminalRequestOrder.length > TERMINAL_REQUEST_CACHE_SIZE)
      terminalRequestIds.delete(terminalRequestOrder.shift()!)
  }
  const cancelActive = (remoteCancelled = false) => {
    if (!active) return
    active.remoteCancelled ||= remoteCancelled
    active.controller.abort()
  }
  const clear = (reason: string, close: boolean) => {
    generation += 1
    cancelActive()
    active = null
    const expiredPendingId = pending?.pairingId
    pending = null
    approvalSentFor = null
    claimedCode = null
    negotiationPending = false
    session = null
    requestIds.clear()
    terminalRequestIds.clear()
    terminalRequestOrder.length = 0
    clearTimers()
    const current = socket
    socket = null
    if (close && current && current.readyState < 2) current.close(1000, 'session_revoked')
    if (expiredPendingId) options.onPendingExpired?.(expiredPendingId)
    setStatus(`disconnected:${reason}` as OfficeRelayStatus)
  }

  const runRequest = async (frame: Record<string, unknown>, owner: number) => {
    if (
      !session ||
      !validId(frame.session_id) ||
      frame.session_id !== session.sessionId ||
      !validId(frame.request_id)
    )
      return clear('protocol_violation', true)
    if (
      active ||
      requestIds.has(frame.request_id) ||
      requestIds.size >= (options.maxRequestIds ?? MAX_REQUEST_IDS)
    )
      return clear('protocol_violation', true)
    if (!jsonObject(frame.body)) return clear('protocol_violation', true)
    requestIds.add(frame.request_id)
    const bodyBytes = Buffer.byteLength(JSON.stringify(frame.body))
    if (bodyBytes > MAX_REQUEST_BYTES) return clear('request_too_large', true)
    const controller = new AbortController()
    active = { requestId: frame.request_id, controller, remoteCancelled: false }
    const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS)
    try {
      const capabilityName = protocolVersion === 2 ? frame.capability_name : 'agent.v1'
      if (
        typeof capabilityName !== 'string' ||
        !session.capabilities.includes(capabilityName) ||
        (capabilityName !== 'agent.v1' && !options.retrievalProxy)
      )
        return clear('protocol_violation', true)
      const response =
        capabilityName === 'agent.v1'
          ? await options.proxy({ body: frame.body, signal: controller.signal })
          : {
              status: 200,
              contentType: 'application/json',
              body: await options.retrievalProxy!(capabilityName, frame.body, controller.signal),
            }
      if (
        !Number.isSafeInteger(response.status) ||
        response.status < 200 ||
        response.status > 599 ||
        response.status === 204 ||
        response.status === 205 ||
        response.status === 304
      )
        throw new Error('unsupported_stream_status')
      const contentType = response.contentType ?? 'application/octet-stream'
      if (!/^[\x20-\x7e]{1,128}$/.test(contentType)) throw new Error('invalid_content_type')
      send({
        version: protocolVersion,
        type: 'pc.start',
        session_id: session.sessionId,
        capability: session.capability,
        request_id: frame.request_id,
        status: response.status,
        content_type: contentType,
      })
      let sequence = 0
      let total = 0
      const responseBody: AsyncIterable<Uint8Array> =
        response.body instanceof Uint8Array
          ? (async function* () {
              yield response.body as Uint8Array
            })()
          : response.body
      for await (const source of responseBody) {
        for (let offset = 0; offset < source.byteLength; offset += MAX_CHUNK_BYTES) {
          if (owner !== generation || controller.signal.aborted || !session) return
          const chunk = source.subarray(offset, offset + MAX_CHUNK_BYTES)
          total += chunk.byteLength
          if (total > MAX_RESPONSE_BYTES) throw new Error('response_too_large')
          send({
            version: protocolVersion,
            type: 'pc.chunk',
            session_id: session.sessionId,
            capability: session.capability,
            request_id: frame.request_id,
            sequence,
            data: Buffer.from(chunk).toString('base64'),
          })
          sequence += 1
        }
      }
      if (owner !== generation || controller.signal.aborted || !session) return
      send({
        version: protocolVersion,
        type: 'pc.done',
        session_id: session.sessionId,
        capability: session.capability,
        request_id: frame.request_id,
      })
    } catch (error) {
      if (owner !== generation || !session) return
      if (active?.requestId === frame.request_id && active.remoteCancelled) return
      if (controller.signal.aborted) return
      const code =
        error instanceof Error && error.message === 'auth_required'
          ? 'auth_required'
          : error instanceof Error && error.message === 'unsupported_stream_status'
            ? 'request_failed'
            : 'upstream_error'
      if (code === 'auth_required') return clear(code, true)
      send({
        version: protocolVersion,
        type: 'pc.error',
        session_id: session.sessionId,
        capability: session.capability,
        request_id: frame.request_id,
        code,
      })
    } finally {
      clearTimeout(timeout)
      rememberTerminalRequest(frame.request_id as string)
      if (active?.requestId === frame.request_id) active = null
    }
  }

  const receive = (event: { data?: unknown }, owner: number) => {
    if (owner !== generation || typeof event.data !== 'string')
      return clear('protocol_violation', true)
    const frameBytes = Buffer.byteLength(event.data)
    if (frameBytes > MAX_REQUEST_BYTES + MAX_CONTROL_BYTES) return clear('protocol_violation', true)
    let frame: unknown
    try {
      frame = JSON.parse(event.data)
    } catch {
      return clear('protocol_violation', true)
    }
    const candidate = frame as Record<string, unknown>
    if (candidate.version === 2 && candidate.type === 'pc.negotiated') {
      if (
        !negotiateCapabilities ||
        !negotiationPending ||
        diagnostic !== 'claiming' ||
        !claimedCode ||
        !exact(candidate, ['version', 'type', 'pairing_version', 'capabilities']) ||
        (candidate.pairing_version !== 1 && candidate.pairing_version !== 2) ||
        !Array.isArray(candidate.capabilities) ||
        candidate.capabilities.length < 1 ||
        candidate.capabilities.length > offeredCapabilities.length ||
        candidate.capabilities.some(
          (value, index, values) =>
            typeof value !== 'string' ||
            !offeredCapabilities.includes(value) ||
            values.indexOf(value) !== index,
        ) ||
        (candidate.pairing_version === 1 &&
          JSON.stringify(candidate.capabilities) !== JSON.stringify(['agent.v1']))
      )
        return clear('protocol_violation', true)
      protocolVersion = candidate.pairing_version
      negotiationPending = false
      send({
        version: protocolVersion,
        type: 'pc.claim',
        verification_code: claimedCode,
        ...(protocolVersion === 2 ? { capabilities: offeredCapabilities } : {}),
      })
      return
    }
    if (
      negotiateCapabilities &&
      negotiationPending &&
      diagnostic === 'claiming' &&
      candidate.version === 2 &&
      candidate.type === 'relay.error' &&
      exact(candidate, ['version', 'type', 'code']) &&
      typeof candidate.code === 'string' &&
      RELAY_ERROR_CODES.has(candidate.code)
    )
      return clear('relay_error', true)
    if (negotiationPending) return clear('protocol_violation', true)
    if (
      !frame ||
      typeof frame !== 'object' ||
      (frame as any).version !== protocolVersion ||
      typeof (frame as any).type !== 'string'
    )
      return clear('protocol_violation', true)
    const typed = frame as Record<string, unknown>
    if (frameBytes > MAX_CONTROL_BYTES && typed.type !== 'relay.request')
      return clear('protocol_violation', true)
    if (typed.type === 'pc.claimed') {
      const keys = [
        'version',
        'type',
        'pairing_id',
        'host',
        'origin',
        'verification_code',
        'expires_in',
      ]
      if (protocolVersion === 2) keys.push('capabilities')
      const negotiated =
        protocolVersion === 2 &&
        Array.isArray(typed.capabilities) &&
        typed.capabilities.length > 0 &&
        typed.capabilities.every(
          (value, index, values) =>
            typeof value === 'string' &&
            V2_CAPABILITIES.includes(value as (typeof V2_CAPABILITIES)[number]) &&
            values.indexOf(value) === index,
        )
          ? (typed.capabilities as string[])
          : protocolVersion === 1
            ? ['agent.v1']
            : null
      if (
        !exact(frame, keys) ||
        !validId(typed.pairing_id) ||
        !HOSTS.has(String(typed.host)) ||
        typeof typed.origin !== 'string' ||
        typed.verification_code !== claimedCode ||
        !Number.isSafeInteger(typed.expires_in) ||
        Number(typed.expires_in) < 1 ||
        Number(typed.expires_in) > 120 ||
        diagnostic !== 'claiming' ||
        pending !== null ||
        !negotiated
      )
        return clear('protocol_violation', true)
      try {
        const origin = new URL(typed.origin)
        if (origin.protocol !== 'https:' || origin.origin !== typed.origin)
          return clear('protocol_violation', true)
      } catch {
        return clear('protocol_violation', true)
      }
      pending = {
        pairingId: typed.pairing_id,
        hostLabel: typed.host as OfficePairingRequest['hostLabel'],
        origin: typed.origin,
        verificationCode: typed.verification_code as string,
        ...(protocolVersion === 2 ? { capabilities: negotiated } : {}),
      }
      pairingTimer = setTimeout(
        () => {
          clear('pairing_expired', true)
        },
        Number(typed.expires_in) * 1_000,
      )
      setStatus('awaiting_approval')
      options.onPending({
        pairingId: pending.pairingId,
        hostLabel: pending.hostLabel,
        origin: pending.origin,
        verificationCode: pending.verificationCode,
      })
      return
    }
    if (typed.type === 'pc.approved') {
      const approvedKeys = ['version', 'type', 'session_id', 'capability', 'expires_in']
      if (protocolVersion === 2) approvedKeys.push('capabilities')
      if (
        !exact(frame, approvedKeys) ||
        !validId(typed.session_id) ||
        !validId(typed.capability) ||
        !Number.isSafeInteger(typed.expires_in) ||
        Number(typed.expires_in) < 1 ||
        Number(typed.expires_in) > 1_800 ||
        !pending ||
        approvalSentFor !== pending.pairingId ||
        diagnostic !== 'awaiting_approval' ||
        (protocolVersion === 2 &&
          JSON.stringify(typed.capabilities) !== JSON.stringify(pending.capabilities))
      )
        return clear('protocol_violation', true)
      session = {
        sessionId: typed.session_id,
        capability: typed.capability,
        capabilities: pending.capabilities ?? ['agent.v1'],
      }
      if (pairingTimer) clearTimeout(pairingTimer)
      pairingTimer = null
      pending = null
      approvalSentFor = null
      sessionTimer = setTimeout(() => clear('session_expired', true), SESSION_ABSOLUTE_MAX_MS)
      setStatus('paired')
      return
    }
    if (typed.type === 'relay.request') {
      const requestKeys = ['version', 'type', 'session_id', 'request_id', 'body']
      if (protocolVersion === 2) requestKeys.push('capability_name')
      if (!exact(frame, requestKeys) || !jsonObject(typed.body))
        return clear('protocol_violation', true)
      void runRequest(typed, owner)
      return
    }
    if (typed.type === 'relay.cancel') {
      if (
        !exact(frame, ['version', 'type', 'session_id', 'request_id']) ||
        typed.session_id !== session?.sessionId ||
        !validId(typed.request_id)
      )
        return clear('protocol_violation', true)
      if (terminalRequestIds.has(typed.request_id)) return
      if (typed.request_id !== active?.requestId) return clear('protocol_violation', true)
      cancelActive(true)
      return
    }
    if (
      typed.type === 'relay.error' &&
      exact(frame, ['version', 'type', 'code']) &&
      typeof typed.code === 'string' &&
      RELAY_ERROR_CODES.has(typed.code)
    )
      return clear('relay_error', true)
    clear('protocol_violation', true)
  }

  return {
    async claim(code) {
      if (!/^\d{6}$/.test(code)) throw new Error('invalid_verification_code')
      const account = await options.getValidAccountStatus()
      if (!account.loggedIn) {
        clear('auth_required', true)
        throw new Error('auth_required')
      }
      const accessToken = await options.getAccessToken().catch((error) => {
        clear('auth_required', true)
        throw error
      })
      if (!accessToken || !/^[\x21-\x7e]+$/.test(accessToken)) {
        clear('auth_required', true)
        throw new Error('auth_required')
      }
      clear('new_claim', true)
      generation += 1
      const owner = generation
      setStatus('connecting')
      let next: RelaySocket
      try {
        next = connect(options.endpoint, accessToken)
      } catch {
        clear('network_error', true)
        throw new Error('relay_connection_failed')
      }
      socket = next
      next.addEventListener('message', (event) => receive(event, owner))
      next.addEventListener('close', () => {
        if (owner === generation) clear('relay_closed', false)
      })
      next.addEventListener('error', () => {
        if (owner === generation) clear('network_error', true)
      })
      await new Promise<void>((resolve, reject) => {
        const timeout = setTimeout(
          () => reject(new Error('relay_connection_timeout')),
          CONNECT_TIMEOUT_MS,
        )
        next.addEventListener('open', () => {
          clearTimeout(timeout)
          resolve()
        })
        next.addEventListener('error', () => {
          clearTimeout(timeout)
          reject(new Error('relay_connection_failed'))
        })
        next.addEventListener('close', () => {
          clearTimeout(timeout)
          reject(new Error('relay_connection_failed'))
        })
      }).catch((error) => {
        if (owner === generation) clear('network_error', true)
        throw error
      })
      if (owner !== generation) throw new Error('relay_connection_failed')
      claimedCode = code
      setStatus('claiming')
      negotiationPending = negotiateCapabilities
      send({
        version: negotiateCapabilities ? 2 : protocolVersion,
        type: negotiateCapabilities ? 'pc.negotiate' : 'pc.claim',
        verification_code: code,
        ...(negotiateCapabilities ? { capabilities: offeredCapabilities } : {}),
      })
    },
    async approve(pairingId) {
      const account = await options.getValidAccountStatus().catch((error) => {
        clear('auth_required', true)
        throw error
      })
      if (!account.loggedIn) {
        clear('auth_required', true)
        return false
      }
      if (!pending || pending.pairingId !== pairingId) return false
      if (approvalSentFor) return false
      approvalSentFor = pairingId
      send({
        version: protocolVersion,
        type: 'pc.approve',
        pairing_id: pairingId,
        ...(protocolVersion === 2 ? { capabilities: pending.capabilities } : {}),
      })
      return true
    },
    reject(pairingId) {
      if (!pending || pending.pairingId !== pairingId) return false
      send({ version: protocolVersion, type: 'pc.reject', pairing_id: pairingId })
      clear('rejected', true)
      return true
    },
    listPending: () => (pending ? [pending] : []),
    status: () => diagnostic,
    revoke: (reason = 'revoked') => clear(reason, true),
  }
}
