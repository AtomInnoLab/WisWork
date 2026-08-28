import type { MessagesProxy } from '@wiswork/office-bridge'
import WebSocket from 'ws'
import type { OfficePairingRequest, OfficeRelayStatus } from '../shared/home-api'
import type { OfficeRelayBinding } from './office-relay-binding-store'
import type { OfficeRetrievalProxy } from './office-retrieval-proxy'
export type { OfficeRelayStatus } from '../shared/home-api'

const MAX_CONTROL_BYTES = 16 * 1024
const MAX_REQUEST_BYTES = 256 * 1024
const MAX_CHUNK_BYTES = 64 * 1024
const MAX_RESPONSE_BYTES = 16 * 1024 * 1024
// Relay owns the 300s deadline; this only handles a lost relay.cancel.
const REQUEST_TIMEOUT_MS = 305_000
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
  'auth_required',
  'binding_unavailable',
  'resume_limit',
  'resume_rate_limited',
])
const TERMINAL_REQUEST_CACHE_SIZE = 64
const PRODUCTION_RELAY_ENDPOINT = 'wss://office.8-216-134-194.sslip.io/office-relay'
const V2_CAPABILITIES = ['agent.v1', 'web-search.v1', 'web-fetch.v1', 'image-search.v1'] as const
const PAIRING_RESUME_FEATURE = 'pairing-resume.v1'

export interface RelaySocket {
  readyState: number
  addEventListener(name: string, listener: (event: any) => void): void
  send(data: string): void
  close(code?: number, reason?: string): void
}

export interface OfficeRelayClient {
  claim(code: string): Promise<void>
  resume(binding: OfficeRelayBinding): Promise<void>
  revokeBinding(bindingId: string): Promise<void>
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
  getValidAccountStatus(): Promise<{ loggedIn: boolean; userId?: string }>
  getAccessToken(): Promise<string | null>
  proxy: MessagesProxy
  retrievalProxy?: OfficeRetrievalProxy
  negotiateCapabilities?: boolean
  persistentPairing?: boolean
  onBinding?: (binding: OfficeRelayBinding) => void | Promise<void>
  onBindingInvalidated?: (bindingId: string) => void | Promise<void>
  onPending(pairing: OfficePairingRequest): void
  onPendingExpired?: (pairingId: string) => void
  onStatus?: (status: OfficeRelayStatus) => void
  maxRequestIds?: number
  now?: () => number
}): OfficeRelayClient {
  const connect = options.connect ?? connectAuthenticatedRelaySocket
  let socket: RelaySocket | null = null
  let diagnostic: OfficeRelayStatus = 'disconnected'
  let protocolVersion: 1 | 2 = 1
  const negotiateCapabilities =
    options.negotiateCapabilities === true ||
    Boolean(options.retrievalProxy) ||
    options.persistentPairing === true
  const persistentPairing = options.persistentPairing === true
  const offeredCapabilities = options.retrievalProxy ? [...V2_CAPABILITIES] : ['agent.v1']
  let pending: (OfficePairingRequest & { capabilities?: string[]; features?: string[] }) | null =
    null
  let session: { sessionId: string; capability: string; capabilities: string[] } | null = null
  let active: { requestId: string; controller: AbortController; remoteCancelled: boolean } | null =
    null
  let claimedCode: string | null = null
  let claimedAccountId: string | null = null
  let negotiationPending = false
  let enhancedNegotiation = false
  let negotiatedFeatures: string[] | null = null
  let legacyFallbackAttempted = false
  let action: 'idle' | 'claim' | 'resume' | 'revoke' = 'idle'
  let resumeBinding: OfficeRelayBinding | null = null
  let revocation: {
    bindingId: string
    resolve(): void
    reject(error: Error): void
    timer: ReturnType<typeof setTimeout>
    completionReason: 'binding_revoked' | 'binding_not_remembered'
  } | null = null
  const requestIds = new Set<string>()
  const terminalRequestIds = new Set<string>()
  const terminalRequestOrder: string[] = []
  let generation = 0
  let approvalSentFor: string | null = null
  let pairingTimer: ReturnType<typeof setTimeout> | null = null
  let sessionTimer: ReturnType<typeof setTimeout> | null = null
  let acceptedApprovalSignature: string | null = null

  const frameSignature = (frame: Record<string, unknown>): string =>
    JSON.stringify(Object.entries(frame).sort(([left], [right]) => left.localeCompare(right)))

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
    claimedAccountId = null
    negotiationPending = false
    enhancedNegotiation = false
    negotiatedFeatures = null
    action = 'idle'
    resumeBinding = null
    if (revocation) {
      clearTimeout(revocation.timer)
      revocation.reject(new Error('relay_connection_failed'))
      revocation = null
    }
    session = null
    acceptedApprovalSignature = null
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

  function fallbackToLegacyClaim(): boolean {
    if (
      action !== 'claim' ||
      !negotiationPending ||
      !enhancedNegotiation ||
      legacyFallbackAttempted ||
      !claimedCode
    )
      return false
    const code = claimedCode
    const current = socket
    socket = null
    legacyFallbackAttempted = true
    negotiationPending = false
    enhancedNegotiation = false
    void beginClaim(code, false, false).catch(() => undefined)
    if (current && current.readyState < 2) current.close(1000, 'legacy_fallback')
    return true
  }

  const receive = (event: { data?: unknown }, owner: number): void | Promise<void> => {
    if (owner !== generation) return
    if (typeof event.data !== 'string') return clear('protocol_violation', true)
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
      const negotiatedKeys = ['version', 'type', 'pairing_version', 'capabilities']
      if (enhancedNegotiation) negotiatedKeys.push('features')
      if (
        !negotiateCapabilities ||
        !negotiationPending ||
        diagnostic !== 'claiming' ||
        !claimedCode ||
        !exact(candidate, negotiatedKeys) ||
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
          JSON.stringify(candidate.capabilities) !== JSON.stringify(['agent.v1'])) ||
        (enhancedNegotiation &&
          (!Array.isArray(candidate.features) ||
            candidate.features.some(
              (value, index, values) =>
                value !== PAIRING_RESUME_FEATURE || values.indexOf(value) !== index,
            ) ||
            candidate.features.length > 1))
      )
        return clear('protocol_violation', true)
      protocolVersion = candidate.pairing_version
      negotiationPending = false
      negotiatedFeatures = enhancedNegotiation ? [...(candidate.features as string[])] : null
      enhancedNegotiation = enhancedNegotiation && protocolVersion === 2
      send({
        version: protocolVersion,
        type: 'pc.claim',
        verification_code: claimedCode,
        ...(protocolVersion === 2 ? { capabilities: offeredCapabilities } : {}),
        ...(protocolVersion === 2 && negotiatedFeatures ? { features: negotiatedFeatures } : {}),
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
    ) {
      if (
        (candidate.code === 'invalid_frame' || candidate.code === 'unknown_type') &&
        fallbackToLegacyClaim()
      )
        return
      return clear('relay_error', true)
    }
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
    if (typed.type === 'pc.waiting_for_office') {
      if (
        action !== 'resume' ||
        !resumeBinding ||
        diagnostic !== 'connecting' ||
        !exact(frame, ['version', 'type'])
      )
        return clear('protocol_violation', true)
      setStatus('waiting_for_office')
      return
    }
    if (typed.type === 'pc.binding_revoked') {
      if (
        action !== 'revoke' ||
        !revocation ||
        !exact(frame, ['version', 'type', 'binding_id']) ||
        typed.binding_id !== revocation.bindingId
      )
        return clear('protocol_violation', true)
      const completed = revocation
      revocation = null
      clearTimeout(completed.timer)
      completed.resolve()
      clear(completed.completionReason, true)
      return
    }
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
      if (enhancedNegotiation) keys.push('features')
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
        !negotiated ||
        (enhancedNegotiation &&
          (!Array.isArray(typed.features) ||
            typed.features.length > 1 ||
            typed.features.some((value) => value !== PAIRING_RESUME_FEATURE) ||
            JSON.stringify(typed.features) !== JSON.stringify(negotiatedFeatures)))
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
        ...(enhancedNegotiation ? { features: typed.features as string[] } : {}),
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
      if (
        acceptedApprovalSignature !== null &&
        frameSignature(typed) === acceptedApprovalSignature
      )
        return
      const approvedKeys = ['version', 'type', 'session_id', 'capability', 'expires_in']
      if (protocolVersion === 2) approvedKeys.push('capabilities')
      const resumed = action === 'resume'
      const remembersBinding =
        !resumed && pending?.features?.includes(PAIRING_RESUME_FEATURE) === true
      if (remembersBinding) approvedKeys.push('features', 'binding_id')
      if (
        !exact(frame, approvedKeys) ||
        !validId(typed.session_id) ||
        !validId(typed.capability) ||
        !Number.isSafeInteger(typed.expires_in) ||
        Number(typed.expires_in) < 1 ||
        Number(typed.expires_in) > 1_800 ||
        (!resumed && (!pending || approvalSentFor !== pending.pairingId)) ||
        (resumed && !resumeBinding) ||
        (!resumed && diagnostic !== 'awaiting_approval') ||
        (resumed && diagnostic !== 'connecting' && diagnostic !== 'waiting_for_office') ||
        (protocolVersion === 2 &&
          JSON.stringify(typed.capabilities) !==
            JSON.stringify(resumed ? resumeBinding?.capabilities : pending?.capabilities)) ||
        (remembersBinding &&
          (!validId(typed.binding_id) ||
            JSON.stringify(typed.features) !== JSON.stringify([PAIRING_RESUME_FEATURE])))
      )
        return clear('protocol_violation', true)
      const approvedPending = pending
      const approvedAccountId = claimedAccountId
      const finalizeApproval = () => {
        session = {
          sessionId: typed.session_id as string,
          capability: typed.capability as string,
          capabilities: resumed
            ? [...resumeBinding!.capabilities]
            : (approvedPending?.capabilities ?? ['agent.v1']),
        }
        acceptedApprovalSignature = frameSignature(typed)
        pending = null
        approvalSentFor = null
        action = 'idle'
        resumeBinding = null
        sessionTimer = setTimeout(() => clear('session_expired', true), SESSION_ABSOLUTE_MAX_MS)
        setStatus('paired')
      }
      if (pairingTimer) clearTimeout(pairingTimer)
      pairingTimer = null
      if (!remembersBinding || !approvedPending || !approvedAccountId) {
        finalizeApproval()
        return
      }
      return (async () => {
        try {
          await options.onBinding?.({
            bindingId: typed.binding_id as string,
            accountId: approvedAccountId,
            host: approvedPending.hostLabel,
            origin: approvedPending.origin,
            capabilities: [...(approvedPending.capabilities ?? ['agent.v1'])],
            createdAt: (options.now ?? Date.now)(),
          })
        } catch {
          await revokeBindingRemote(
            typed.binding_id as string,
            'binding_not_remembered',
          ).catch(() => {
            clear('binding_not_remembered', true)
          })
          return
        }
        if (owner !== generation) return
        finalizeApproval()
      })()
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
    ) {
      if (
        action === 'resume' &&
        resumeBinding &&
        (typed.code === 'binding_unavailable' || typed.code === 'capability_not_negotiated')
      )
        void Promise.resolve(options.onBindingInvalidated?.(resumeBinding.bindingId)).catch(
          () => undefined,
        )
      const reason =
        typed.code === 'session_expired' ||
        typed.code === 'auth_required' ||
        typed.code === 'binding_unavailable' ||
        typed.code === 'capability_not_negotiated' ||
        typed.code === 'resume_rate_limited' ||
        typed.code === 'resume_limit' ||
        typed.code === 'peer_unavailable'
          ? typed.code
          : 'relay_error'
      return clear(reason, true)
    }
    clear('protocol_violation', true)
  }

  const openAuthenticated = async (
    owner: number,
    expectedAccountId?: string,
  ): Promise<{ accountId: string | null }> => {
    const ensureCurrent = () => {
      if (owner !== generation) throw new Error('relay_connection_failed')
    }
    const account = await options.getValidAccountStatus().catch((error) => {
      ensureCurrent()
      clear('auth_required', true)
      throw error
    })
    ensureCurrent()
    if (
      !account.loggedIn ||
      (expectedAccountId !== undefined && account.userId !== expectedAccountId)
    ) {
      clear('auth_required', true)
      throw new Error('auth_required')
    }
    const accessToken = await options.getAccessToken().catch((error) => {
      ensureCurrent()
      clear('auth_required', true)
      throw error
    })
    ensureCurrent()
    if (!accessToken || !/^[\x21-\x7e]+$/.test(accessToken)) {
      clear('auth_required', true)
      throw new Error('auth_required')
    }
    setStatus('connecting')
    let next: RelaySocket
    try {
      next = connect(options.endpoint, accessToken)
    } catch {
      clear('network_error', true)
      throw new Error('relay_connection_failed')
    }
    socket = next
    let receivePending = false
    const receiveQueue: Array<{ data?: unknown }> = []
    const dispatch = (event: { data?: unknown }) => {
      let nextEvent: { data?: unknown } | undefined = event
      while (nextEvent) {
        const result = receive(nextEvent, owner)
        if (result) {
          receivePending = true
          void result
            .catch(() => {
              if (owner === generation) clear('protocol_violation', true)
            })
            .finally(() => {
              receivePending = false
              const queued = receiveQueue.shift()
              if (queued) dispatch(queued)
            })
          return
        }
        nextEvent = receiveQueue.shift()
      }
    }
    next.addEventListener('message', (event) => {
      if (receivePending) receiveQueue.push(event)
      else dispatch(event)
    })
    next.addEventListener('close', () => {
      if (owner !== generation) return
      clear('relay_closed', false)
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
    ensureCurrent()
    return { accountId: account.userId ?? null }
  }

  async function beginClaim(code: string, enhanced: boolean, reset: boolean): Promise<void> {
    if (reset) {
      clear('new_claim', true)
      legacyFallbackAttempted = false
    }
    generation += 1
    const owner = generation
    action = 'claim'
    claimedCode = code
    enhancedNegotiation = enhanced
    const account = await openAuthenticated(owner)
    if (enhanced && !account.accountId) {
      clear('auth_required', true)
      throw new Error('auth_required')
    }
    claimedAccountId = account.accountId
    setStatus('claiming')
    negotiationPending = negotiateCapabilities
    send({
      version: negotiateCapabilities ? 2 : protocolVersion,
      type: negotiateCapabilities ? 'pc.negotiate' : 'pc.claim',
      verification_code: code,
      ...(negotiateCapabilities ? { capabilities: offeredCapabilities } : {}),
      ...(enhanced ? { features: [PAIRING_RESUME_FEATURE] } : {}),
    })
  }

  async function revokeBindingRemote(
    bindingId: string,
    completionReason: 'binding_revoked' | 'binding_not_remembered' = 'binding_revoked',
  ): Promise<void> {
    if (!validId(bindingId)) throw new Error('invalid_office_binding')
    clear('new_revocation', true)
    generation += 1
    const owner = generation
    action = 'revoke'
    protocolVersion = 2
    await openAuthenticated(owner)
    return new Promise<void>((resolve, reject) => {
      const timer = setTimeout(() => {
        if (revocation?.bindingId !== bindingId) return
        revocation = null
        reject(new Error('relay_connection_timeout'))
        clear('network_error', true)
      }, CONNECT_TIMEOUT_MS)
      revocation = { bindingId, resolve, reject, timer, completionReason }
      send({ version: 2, type: 'pc.revoke_binding', binding_id: bindingId })
    })
  }

  return {
    async claim(code) {
      if (!/^\d{6}$/.test(code)) throw new Error('invalid_verification_code')
      await beginClaim(code, persistentPairing, true)
    },
    async resume(binding) {
      if (
        !validId(binding.bindingId) ||
        typeof binding.accountId !== 'string' ||
        !HOSTS.has(binding.host) ||
        !Array.isArray(binding.capabilities) ||
        binding.capabilities.length < 1 ||
        binding.capabilities.some(
          (capability, index, values) =>
            !V2_CAPABILITIES.includes(capability as (typeof V2_CAPABILITIES)[number]) ||
            values.indexOf(capability) !== index,
        )
      )
        throw new Error('invalid_office_binding')
      clear('new_resume', true)
      generation += 1
      const owner = generation
      action = 'resume'
      resumeBinding = { ...binding, capabilities: [...binding.capabilities] }
      protocolVersion = 2
      await openAuthenticated(owner, binding.accountId)
      send({
        version: 2,
        type: 'pc.resume',
        binding_id: binding.bindingId,
        capabilities: binding.capabilities,
      })
    },
    async revokeBinding(bindingId) {
      return revokeBindingRemote(bindingId)
    },
    async approve(pairingId) {
      const account = await options.getValidAccountStatus().catch((error) => {
        clear('auth_required', true)
        throw error
      })
      if (!account.loggedIn || (claimedAccountId && account.userId !== claimedAccountId)) {
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
        ...(pending.features ? { features: pending.features } : {}),
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
