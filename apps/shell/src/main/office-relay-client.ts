import type {
  MessagesProxy,
  MessagesProxyResponse,
  OfficeEnhancedSessionStatement,
} from '@wiswork/office-bridge'
import WebSocket from 'ws'
import type { OfficePairingRequest, OfficeRelayStatus } from '../shared/home-api'
import type { OfficeRelayBinding } from './office-relay-binding-store'
import type { OfficeRetrievalProxy, OfficeWebCapability } from './office-retrieval-proxy'
import type { OfficeDesignDocumentHandler } from './office-design-document'
export type { OfficeRelayStatus } from '../shared/home-api'

const MAX_CONTROL_BYTES = 16 * 1024
const MAX_REQUEST_BYTES = 256 * 1024
const MAX_CHUNK_BYTES = 64 * 1024
const MAX_RESPONSE_BYTES = 16 * 1024 * 1024
// Relay owns the 300s deadline; this only handles a lost relay.cancel.
const REQUEST_TIMEOUT_MS = 305_000
// Match Relay's extended agent.v1 envelope, with five seconds for relay.cancel.
const AGENT_REQUEST_TIMEOUT_MS = 30 * 60_000 + 25_000
const CONNECT_TIMEOUT_MS = 10_000
const ENHANCED_LEASE_RENEW_BEFORE_MS = 5 * 60_000
const ENHANCED_LEASE_MS = 15 * 60_000
const ENHANCED_LEASE_CAPABILITY = 'enhanced-lease.v1'
// Relay owns the renewable idle TTL. PC keeps only a bounded absolute-lifetime watchdog.
const SESSION_ABSOLUTE_MAX_MS = 8 * 60 * 60 * 1_000
const IDENTIFIER = /^[A-Za-z0-9_-]{8,128}$/
const HOSTS = new Set(['Word', 'Excel', 'PowerPoint'])
const MAX_REQUEST_IDS = 2_048
// A visible DESIGN reader can poll every 5s for the full 8h session. Reserve
// those IDs plus the existing interactive budget; retain every ID for replay
// detection rather than evicting old requests to make room for polling.
const MAX_DESIGN_REQUEST_IDS = Math.ceil(SESSION_ABSOLUTE_MAX_MS / 5_000) + MAX_REQUEST_IDS
const MAX_PENDING_TOOLS = 8
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
const V2_CAPABILITIES = [
  'agent.v1',
  'web-search.v1',
  'web-fetch.v1',
  'image-search.v1',
  'image-fetch.v1',
  'design-document.v1',
  ENHANCED_LEASE_CAPABILITY,
] as const
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
  revokeBinding(bindingId: string, expectedAccountId: string): Promise<void>
  approve(pairingId: string): Promise<boolean>
  reject(pairingId: string): boolean
  listPending(): OfficePairingRequest[]
  status(): OfficeRelayStatus
  revoke(reason?: string): void
}

export interface OfficeRelayToolCall {
  readonly turnId: string
  readonly callId: string
  readonly generation: number
  readonly toolName: string
  readonly input: Record<string, unknown>
}
export interface OfficeRelayToolResult {
  readonly output: string
  readonly isError: boolean
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
  enhancedProxy?: (request: {
    body: unknown
    signal: AbortSignal
    host: OfficePairingRequest['hostLabel']
    sessionId: string
    requestId: string
    statement: Readonly<OfficeEnhancedSessionStatement>
    executeTool(call: OfficeRelayToolCall, signal?: AbortSignal): Promise<OfficeRelayToolResult>
  }) => Promise<MessagesProxyResponse>
  enhancedStatement?: (
    host: OfficePairingRequest['hostLabel'],
  ) => Readonly<OfficeEnhancedSessionStatement> | undefined
  renewEnhancedStatement?: (
    previous: Readonly<OfficeEnhancedSessionStatement>,
  ) => Promise<Readonly<OfficeEnhancedSessionStatement> | undefined>
  isEnhancedStatementCurrent?: (statement: Readonly<OfficeEnhancedSessionStatement>) => boolean
  retrievalProxy?: OfficeRetrievalProxy
  designDocument?: OfficeDesignDocumentHandler
  retrievalCapabilities?: readonly OfficeWebCapability[]
  negotiateCapabilities?: boolean
  persistentPairing?: boolean | (() => boolean)
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
  const enhancedLeaseAvailable = Boolean(
    options.renewEnhancedStatement && options.isEnhancedStatementCurrent,
  )
  const negotiateCapabilities =
    options.negotiateCapabilities === true ||
    enhancedLeaseAvailable ||
    Boolean(options.retrievalProxy) ||
    options.persistentPairing === true ||
    typeof options.persistentPairing === 'function'
  const persistentPairingEnabled = () =>
    typeof options.persistentPairing === 'function'
      ? options.persistentPairing() === true
      : options.persistentPairing === true
  const offeredCapabilities = [
    'agent.v1',
    ...(options.retrievalProxy
      ? (options.retrievalCapabilities ??
        V2_CAPABILITIES.filter(
          (name) =>
            name !== 'agent.v1' &&
            name !== 'design-document.v1' &&
            name !== ENHANCED_LEASE_CAPABILITY,
        ))
      : []),
    ...(options.designDocument ? ['design-document.v1'] : []),
    ...(enhancedLeaseAvailable ? [ENHANCED_LEASE_CAPABILITY] : []),
  ]
  let pending: (OfficePairingRequest & { capabilities?: string[]; features?: string[] }) | null =
    null
  let session: {
    sessionId: string
    capability: string
    capabilities: string[]
    host: OfficePairingRequest['hostLabel']
    accountId: string | null
    enhanced?: Readonly<OfficeEnhancedSessionStatement>
  } | null = null
  let active: { requestId: string; controller: AbortController; remoteCancelled: boolean } | null =
    null
  let pendingTool: {
    requestId: string
    call: OfficeRelayToolCall
    resolve(value: OfficeRelayToolResult): void
    reject(error: Error): void
  } | null = null
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
  // Only dispatched, cancelled calls may produce a late result on this socket.
  const cancelledToolResults = new Set<string>()
  const toolResultKey = (requestId: string, call: OfficeRelayToolCall) =>
    JSON.stringify([requestId, call.turnId, call.callId, call.generation])
  let generation = 0
  let approvalSentFor: string | null = null
  let pairingTimer: ReturnType<typeof setTimeout> | null = null
  let sessionTimer: ReturnType<typeof setTimeout> | null = null
  let leaseTimer: ReturnType<typeof setTimeout> | null = null
  let renewalTimer: ReturnType<typeof setTimeout> | null = null
  let renewalAttempt: { timer: ReturnType<typeof setTimeout> } | null = null
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
    const maximum =
      frame.type === 'pc.tool_call' ? MAX_REQUEST_BYTES + MAX_CONTROL_BYTES : MAX_CONTROL_BYTES
    if (Buffer.byteLength(raw) > maximum && frame.type !== 'pc.chunk')
      throw new Error('control_frame_too_large')
    socket.send(raw)
  }
  const clearTimers = () => {
    if (pairingTimer) clearTimeout(pairingTimer)
    if (sessionTimer) clearTimeout(sessionTimer)
    if (leaseTimer) clearTimeout(leaseTimer)
    if (renewalTimer) clearTimeout(renewalTimer)
    if (renewalAttempt) clearTimeout(renewalAttempt.timer)
    pairingTimer = null
    sessionTimer = null
    leaseTimer = null
    renewalTimer = null
    renewalAttempt = null
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
    pendingTool?.reject(new Error('tool_cancelled'))
    pendingTool = null
  }
  const clear = (reason: string, close: boolean) => {
    generation += 1
    cancelActive()
    try {
      options.retrievalProxy?.clear?.()
    } catch {
      // A cleanup failure must not prevent revoking the Relay session.
    }
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
    cancelledToolResults.clear()
    clearTimers()
    const current = socket
    socket = null
    if (close && current && current.readyState < 2) current.close(1000, 'session_revoked')
    if (expiredPendingId) options.onPendingExpired?.(expiredPendingId)
    setStatus(`disconnected:${reason}` as OfficeRelayStatus)
  }

  const armEnhancedLease = () => {
    if (leaseTimer) clearTimeout(leaseTimer)
    if (renewalTimer) clearTimeout(renewalTimer)
    leaseTimer = null
    renewalTimer = null
    const current = session
    const previous = current?.enhanced
    if (!current || !previous) return
    const now = options.now ?? Date.now
    const remaining = previous.expires_at - now()
    if (remaining <= 0) return clear('session_expired', true)
    leaseTimer = setTimeout(() => clear('session_expired', true), remaining)
    const renew = options.renewEnhancedStatement
    if (
      !enhancedLeaseAvailable ||
      !renew ||
      !current.accountId ||
      !current.capabilities.includes(ENHANCED_LEASE_CAPABILITY)
    )
      return
    const owner = generation
    renewalTimer = setTimeout(
      () => {
        renewalTimer = null
        if (renewalAttempt) return
        const attempt = {
          timer: setTimeout(() => {
            if (renewalAttempt === attempt) renewalAttempt = null
          }, CONNECT_TIMEOUT_MS),
        }
        renewalAttempt = attempt
        const isCurrent = () =>
          renewalAttempt === attempt &&
          owner === generation &&
          session === current &&
          current.enhanced === previous &&
          socket?.readyState === 1 &&
          previous.expires_at > now()
        const checkAccount = async () => {
          const account = await options.getValidAccountStatus()
          if (!isCurrent()) return false
          if (!account.loggedIn || account.userId !== current.accountId) {
            clear('auth_required', true)
            return false
          }
          return true
        }
        void (async () => {
          if (!isCurrent()) return
          if (!(await checkAccount()) || !isCurrent()) return
          const renewed = await renew(previous)
          if (!isCurrent()) return
          if (!(await checkAccount()) || !isCurrent()) return
          if (!renewed) return clear('enhanced_authority_unavailable', true)
          if (
            !exact(renewed, Object.keys(previous)) ||
            Object.keys(previous).some(
              (key) =>
                key !== 'expires_at' &&
                renewed[key as keyof OfficeEnhancedSessionStatement] !==
                  previous[key as keyof OfficeEnhancedSessionStatement],
            ) ||
            !Number.isSafeInteger(renewed.expires_at) ||
            renewed.expires_at <= previous.expires_at ||
            renewed.expires_at <= now() ||
            renewed.expires_at > now() + ENHANCED_LEASE_MS
          )
            return clear('protocol_violation', true)
          // The last account await may outlive a runtime crash or policy revocation.
          if (options.isEnhancedStatementCurrent?.(renewed) !== true)
            return clear('enhanced_authority_unavailable', true)
          current.enhanced = renewed
          send({
            version: 2,
            type: 'pc.session_state',
            session_id: current.sessionId,
            capability: current.capability,
            generation: renewed.session_generation,
            enhanced: renewed,
          })
          armEnhancedLease()
        })()
          .catch(() => {
            // A failed or stalled renewal never moves the existing expiry watchdog.
          })
          .finally(() => {
            clearTimeout(attempt.timer)
            if (renewalAttempt === attempt) renewalAttempt = null
          })
      },
      // A buggy short extension must not cause a tight renewal loop.
      Math.max(60_000, remaining - ENHANCED_LEASE_RENEW_BEFORE_MS),
    )
  }

  const promoteEnhancedSession = (): void => {
    if (!session || session.enhanced || protocolVersion !== 2) return
    const enhanced = options.enhancedStatement?.(session.host)
    if (!enhanced) return
    session = { ...session, enhanced }
    send({
      version: 2,
      type: 'pc.session_state',
      session_id: session.sessionId,
      capability: session.capability,
      generation: enhanced.session_generation,
      enhanced,
    })
    armEnhancedLease()
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
      requestIds.size >=
        (options.maxRequestIds ??
          (session.capabilities.includes('design-document.v1')
            ? MAX_DESIGN_REQUEST_IDS
            : MAX_REQUEST_IDS))
    )
      return clear('protocol_violation', true)
    if (!jsonObject(frame.body)) return clear('protocol_violation', true)
    promoteEnhancedSession()
    requestIds.add(frame.request_id)
    const bodyBytes = Buffer.byteLength(JSON.stringify(frame.body))
    if (bodyBytes > MAX_REQUEST_BYTES) return clear('request_too_large', true)
    const controller = new AbortController()
    active = { requestId: frame.request_id, controller, remoteCancelled: false }
    const queue: NonNullable<typeof pendingTool>[] = []
    const disposeTools = () => {
      const dispatched = pendingTool?.requestId === frame.request_id ? pendingTool : null
      if (dispatched) {
        pendingTool = null
        cancelledToolResults.add(toolResultKey(dispatched.requestId, dispatched.call))
        while (cancelledToolResults.size > TERMINAL_REQUEST_CACHE_SIZE)
          cancelledToolResults.delete(cancelledToolResults.values().next().value!)
        dispatched.reject(new Error('tool_cancelled'))
      }
      for (const tool of queue.splice(0)) tool.reject(new Error('tool_cancelled'))
    }
    const abortRequest = () => {
      if (active?.controller !== controller) return
      const remoteCancelled = active.remoteCancelled
      active = null
      rememberTerminalRequest(frame.request_id as string)
      disposeTools()
      if (!remoteCancelled && session && owner === generation) {
        try {
          // The wire is single-flight and has no per-tool cancel. End this
          // request before releasing its slot; never replay an unresolved write.
          send({
            version: protocolVersion,
            type: 'pc.error',
            session_id: session.sessionId,
            capability: session.capability,
            request_id: frame.request_id,
            code: 'cancelled',
          })
        } catch {
          // Local cancellation must settle even if the socket has already gone.
        }
      }
    }
    controller.signal.addEventListener('abort', abortRequest, { once: true })
    const timeout = setTimeout(
      () => controller.abort(),
      protocolVersion === 2 && frame.capability_name === 'agent.v1'
        ? AGENT_REQUEST_TIMEOUT_MS
        : REQUEST_TIMEOUT_MS,
    )
    try {
      const capabilityName = protocolVersion === 2 ? frame.capability_name : 'agent.v1'
      if (
        typeof capabilityName !== 'string' ||
        capabilityName === ENHANCED_LEASE_CAPABILITY ||
        !session.capabilities.includes(capabilityName) ||
        (capabilityName === 'design-document.v1'
          ? !options.designDocument
          : capabilityName !== 'agent.v1' && !options.retrievalProxy)
      )
        return clear('protocol_violation', true)
      const dispatchNext = () => {
        if (
          pendingTool ||
          controller.signal.aborted ||
          active?.controller !== controller ||
          !session
        )
          return
        const tool = queue.shift()
        if (!tool) return
        pendingTool = tool
        try {
          send({
            version: 2,
            type: 'pc.tool_call',
            session_id: session.sessionId,
            capability: session.capability,
            request_id: frame.request_id,
            turn_id: tool.call.turnId,
            call_id: tool.call.callId,
            generation: tool.call.generation,
            tool_name: tool.call.toolName,
            input: tool.call.input,
          })
        } catch {
          controller.abort()
        }
      }
      const executeTool = (
        call: OfficeRelayToolCall,
        signal?: AbortSignal,
      ): Promise<OfficeRelayToolResult> => {
        if (signal?.aborted || controller.signal.aborted)
          return Promise.reject(new Error('tool_cancelled'))
        if (!session || active?.requestId !== frame.request_id)
          return Promise.reject(new Error('tool_unavailable'))
        if (
          !validId(call.turnId) ||
          !validId(call.callId) ||
          !Number.isSafeInteger(call.generation) ||
          call.generation !== session.enhanced?.session_generation ||
          !/^[A-Za-z0-9_-]{1,128}$/.test(call.toolName) ||
          !jsonObject(call.input) ||
          Buffer.byteLength(JSON.stringify(call.input)) > MAX_REQUEST_BYTES
        )
          return Promise.reject(new Error('invalid_tool_call'))
        if (queue.length + (pendingTool ? 1 : 0) >= MAX_PENDING_TOOLS)
          return Promise.reject(new Error('tool_queue_full'))
        return new Promise((resolve, reject) => {
          const finish = () => signal?.removeEventListener('abort', abort)
          const tool: NonNullable<typeof pendingTool> = {
            requestId: frame.request_id as string,
            call: structuredClone(call),
            resolve(result) {
              finish()
              resolve(result)
              dispatchNext()
            },
            reject(error) {
              finish()
              reject(error)
            },
          }
          const abort = () => {
            if (pendingTool === tool) controller.abort()
            else {
              const index = queue.indexOf(tool)
              if (index >= 0) queue.splice(index, 1)
              tool.reject(new Error('tool_cancelled'))
            }
          }
          signal?.addEventListener('abort', abort, { once: true })
          queue.push(tool)
          dispatchNext()
        })
      }
      const response =
        capabilityName === 'agent.v1'
          ? session.enhanced
            ? options.enhancedProxy
              ? await options.enhancedProxy({
                  body: frame.body,
                  signal: controller.signal,
                  host: session.host,
                  sessionId: session.sessionId,
                  requestId: frame.request_id as string,
                  statement: session.enhanced,
                  executeTool,
                  ...(options.retrievalProxy ? { executeRetrieval: options.retrievalProxy } : {}),
                })
              : (() => {
                  throw new Error('enhanced_proxy_unavailable')
                })()
            : await options.proxy({ body: frame.body, signal: controller.signal })
          : capabilityName === 'design-document.v1'
            ? {
                status: 200,
                contentType: 'application/json',
                body: new TextEncoder().encode(
                  JSON.stringify(
                    await options.designDocument!({
                      sessionId: session.sessionId,
                      host: session.host,
                      body: frame.body,
                      signal: controller.signal,
                    }),
                  ),
                ),
              }
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
      if (owner !== generation || controller.signal.aborted || !session) return
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
      if (pendingTool?.requestId === frame.request_id || queue.length)
        throw new Error('unresolved_tool')
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
      controller.signal.removeEventListener('abort', abortRequest)
      disposeTools()
      rememberTerminalRequest(frame.request_id as string)
      if (active?.controller === controller) active = null
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
    // Tool results are data frames too; the outer 272 KiB ceiling still applies.
    if (
      frameBytes > MAX_CONTROL_BYTES &&
      typed.type !== 'relay.request' &&
      typed.type !== 'relay.tool_result'
    )
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
            offeredCapabilities.includes(value) &&
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
      if (acceptedApprovalSignature !== null && frameSignature(typed) === acceptedApprovalSignature)
        return
      const approvedKeys = ['version', 'type', 'session_id', 'capability', 'expires_in']
      if (protocolVersion === 2) approvedKeys.push('capabilities')
      const resumed = action === 'resume'
      const featureResultExpected =
        !resumed && pending !== null && Object.hasOwn(pending, 'features')
      const remembersBinding =
        featureResultExpected && pending?.features?.includes(PAIRING_RESUME_FEATURE) === true
      const durableBindingApproval =
        remembersBinding && exact(frame, [...approvedKeys, 'features', 'binding_id'])
      const shortSessionFallback =
        featureResultExpected &&
        exact(frame, [...approvedKeys, 'features']) &&
        Array.isArray(typed.features) &&
        typed.features.length === 0
      if (
        (featureResultExpected
          ? !durableBindingApproval && !shortSessionFallback
          : !exact(frame, approvedKeys)) ||
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
        (durableBindingApproval &&
          (!validId(typed.binding_id) ||
            JSON.stringify(typed.features) !== JSON.stringify([PAIRING_RESUME_FEATURE])))
      )
        return clear('protocol_violation', true)
      const approvedPending = pending
      const approvedAccountId = claimedAccountId
      const finalizeApproval = () => {
        const host = resumed ? resumeBinding!.host : approvedPending!.hostLabel
        const enhanced = protocolVersion === 2 ? options.enhancedStatement?.(host) : undefined
        session = {
          sessionId: typed.session_id as string,
          capability: typed.capability as string,
          capabilities: resumed
            ? [...resumeBinding!.capabilities]
            : (approvedPending?.capabilities ?? ['agent.v1']),
          host,
          accountId: resumed ? resumeBinding!.accountId : approvedAccountId,
          ...(enhanced ? { enhanced } : {}),
        }
        if (protocolVersion === 2) {
          send({
            version: 2,
            type: 'pc.session_state',
            session_id: session.sessionId,
            capability: session.capability,
            generation: enhanced?.session_generation ?? 0,
            enhanced: enhanced ?? null,
          })
        }
        acceptedApprovalSignature = frameSignature(typed)
        pending = null
        approvalSentFor = null
        action = 'idle'
        resumeBinding = null
        sessionTimer = setTimeout(() => clear('session_expired', true), SESSION_ABSOLUTE_MAX_MS)
        setStatus('paired')
        armEnhancedLease()
      }
      if (pairingTimer) clearTimeout(pairingTimer)
      pairingTimer = null
      if (!durableBindingApproval || !approvedPending || !approvedAccountId) {
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
        } catch (error) {
          if (error instanceof Error && error.message === 'stale_binding_account') {
            clear('binding_not_remembered', true)
            return
          }
          await revokeBindingRemote(
            typed.binding_id as string,
            approvedAccountId,
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
    if (typed.type === 'relay.tool_result') {
      const tool = pendingTool
      if (
        !session ||
        !exact(frame, [
          'version',
          'type',
          'session_id',
          'request_id',
          'turn_id',
          'call_id',
          'generation',
          'output',
          'is_error',
        ]) ||
        typed.session_id !== session.sessionId ||
        !validId(typed.request_id) ||
        !validId(typed.turn_id) ||
        !validId(typed.call_id) ||
        !Number.isSafeInteger(typed.generation) ||
        typeof typed.output !== 'string' ||
        Buffer.byteLength(typed.output) > MAX_RESPONSE_BYTES ||
        typeof typed.is_error !== 'boolean'
      )
        return clear('protocol_violation', true)
      const cancelledKey = JSON.stringify([
        typed.request_id,
        typed.turn_id,
        typed.call_id,
        typed.generation,
      ])
      if (cancelledToolResults.has(cancelledKey)) return
      if (
        !active ||
        !tool ||
        typed.request_id !== active.requestId ||
        typed.request_id !== tool.requestId ||
        typed.turn_id !== tool.call.turnId ||
        typed.call_id !== tool.call.callId ||
        typed.generation !== tool.call.generation
      )
        return clear('protocol_violation', true)
      pendingTool = null
      tool.resolve({ output: typed.output, isError: typed.is_error })
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
      if (action === 'revoke' && revocation && typed.code === 'binding_unavailable') {
        const completed = revocation
        revocation = null
        clearTimeout(completed.timer)
        completed.resolve()
        clear(completed.completionReason, true)
        return
      }
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
    if (expectedAccountId !== undefined) {
      const confirmed = await options.getValidAccountStatus().catch((error) => {
        ensureCurrent()
        clear('auth_required', true)
        throw error
      })
      ensureCurrent()
      if (!confirmed.loggedIn || confirmed.userId !== expectedAccountId) {
        clear('auth_required', true)
        throw new Error('auth_required')
      }
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
    expectedAccountId: string,
    completionReason: 'binding_revoked' | 'binding_not_remembered' = 'binding_revoked',
  ): Promise<void> {
    if (!validId(bindingId) || typeof expectedAccountId !== 'string' || !expectedAccountId)
      throw new Error('invalid_office_binding')
    clear('new_revocation', true)
    generation += 1
    const owner = generation
    action = 'revoke'
    protocolVersion = 2
    await openAuthenticated(owner, expectedAccountId)
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
      await beginClaim(code, persistentPairingEnabled(), true)
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
    async revokeBinding(bindingId, expectedAccountId) {
      return revokeBindingRemote(bindingId, expectedAccountId)
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
