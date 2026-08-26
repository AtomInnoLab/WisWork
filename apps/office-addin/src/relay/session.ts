import type { OfficeHost } from '../office-document.js'
import type { OfficeDiagnosticEvent } from '../diagnostics/office-diagnostics.js'
import { officeTransportMode, type OfficeTransportMode } from '../../build-config.js'

export const OFFICE_RELAY_URL = 'wss://office.8-216-134-194.sslip.io/office-relay'
const MESSAGES_PATH = '/v1/office/messages'
const MAX_CONTROL_FRAME_BYTES = 16 * 1024
const MAX_REQUEST_BYTES = 256 * 1024
const MAX_CHUNK_BYTES = 64 * 1024
const MAX_RELAY_FRAME_BYTES = Math.ceil((MAX_CHUNK_BYTES * 4) / 3) + 4096
const MAX_RESPONSE_BYTES = 16 * 1024 * 1024
// Finish before Relay's 300s deadline so the client owns cancellation and preserves pairing.
const REQUEST_TIMEOUT_MS = 290_000
const MAX_OPAQUE_LENGTH = 512
const MAX_DIAGNOSTIC_EVENT_BYTES = 4 * 1024
const MAX_PENDING_DIAGNOSTICS = 16
const TERMINAL_REQUEST_CACHE_SIZE = 64
const DIAGNOSTIC_ERROR_CODES = new Set([
  'diagnostic_limit',
  'diagnostic_rate_limited',
  'diagnostic_host_mismatch',
  'diagnostic_too_large',
  'invalid_capability',
  'invalid_frame',
  'invalid_session',
])

export { officeTransportMode, type OfficeTransportMode }

export interface RelayWebSocket {
  readonly readyState: number
  onopen: (() => void) | null
  onmessage: ((event: { data: unknown }) => void) | null
  onclose: (() => void) | null
  onerror: (() => void) | null
  send(value: string): void
  close(): void
}

export type OfficeRelayStatus =
  'offline' | 'connecting' | 'pending' | 'waiting_for_pc' | 'rejected' | 'expired' | 'connected'

export interface OfficeRelaySnapshot {
  status: OfficeRelayStatus
  verificationCode?: string
  capabilities?: readonly string[]
}

export interface OfficeRelaySession {
  snapshot(): OfficeRelaySnapshot
  subscribe(listener: () => void): () => void
  connect(host: OfficeHost): Promise<void>
  disconnect(): void
  authenticatedFetch(path: typeof MESSAGES_PATH, init: RequestInit): Promise<Response>
  capabilityFetch(
    capability: OfficeRelayCapability,
    body: unknown,
    signal?: AbortSignal,
  ): Promise<Response>
  sendDiagnostic(event: OfficeDiagnosticEvent): Promise<void>
}

export type OfficeRelayCapability =
  'agent.v1' | 'web-search.v1' | 'web-fetch.v1' | 'image-search.v1'

interface Dependencies {
  createSocket?: (url: string) => RelayWebSocket
  randomUUID?: () => string
  capabilities?: readonly OfficeRelayCapability[]
}

interface ActiveRequest {
  id: string
  sequence: number
  bytes: number
  controller?: ReadableStreamDefaultController<Uint8Array>
  responseResolved: boolean
  resolve: (response: Response) => void
  reject: (error: Error) => void
  timer: ReturnType<typeof setTimeout>
  abort?: () => void
  signal?: AbortSignal
}

const hostLabels: Record<Exclude<OfficeHost, 'unknown'>, string> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
}
const encoder = new TextEncoder()
const exactKeys = (value: Record<string, unknown>, keys: readonly string[]) =>
  Object.keys(value).length === keys.length && keys.every((key) => key in value)
const opaque = (value: unknown): value is string =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= MAX_OPAQUE_LENGTH &&
  /^[A-Za-z0-9_-]+$/.test(value)
const expiry = (value: unknown, maximum: number): value is number =>
  typeof value === 'number' && Number.isSafeInteger(value) && value > 0 && value <= maximum

function browserSocket(url: string): RelayWebSocket {
  return new WebSocket(url) as unknown as RelayWebSocket
}

export function createOfficeRelaySession(dependencies: Dependencies = {}): OfficeRelaySession {
  const createSocket = dependencies.createSocket ?? browserSocket
  const randomUUID = dependencies.randomUUID ?? (() => crypto.randomUUID())
  const requestedCapabilities = [...(dependencies.capabilities ?? [])]
  const protocolVersion = requestedCapabilities.length > 0 ? 2 : 1
  const listeners = new Set<() => void>()
  let state: OfficeRelaySnapshot = { status: 'offline' }
  let socket: RelayWebSocket | undefined
  let pairingId: string | undefined
  let sessionId: string | undefined
  let capability: string | undefined
  let negotiatedCapabilities: OfficeRelayCapability[] = []
  let request: ActiveRequest | undefined
  let generation = 0
  let settleConnect: (() => void) | undefined
  let pairingTimer: ReturnType<typeof setTimeout> | undefined
  const pendingDiagnostics = new Map<string, { resolve(): void; reject(error: Error): void }>()
  const terminalRequestIds = new Set<string>()
  const terminalRequestOrder: string[] = []

  const rememberTerminalRequest = (requestId: string) => {
    if (terminalRequestIds.has(requestId)) return
    terminalRequestIds.add(requestId)
    terminalRequestOrder.push(requestId)
    while (terminalRequestOrder.length > TERMINAL_REQUEST_CACHE_SIZE)
      terminalRequestIds.delete(terminalRequestOrder.shift()!)
  }

  const publish = (next: OfficeRelaySnapshot) => {
    state = next
    listeners.forEach((listener) => listener())
  }
  const finishRequest = (error?: string) => {
    const active = request
    if (!active) return
    rememberTerminalRequest(active.id)
    request = undefined
    clearTimeout(active.timer)
    if (active.abort && active.signal) active.signal.removeEventListener('abort', active.abort)
    if (active.responseResolved && active.controller) {
      try {
        if (error) active.controller.error(new Error(error))
        else active.controller.close()
      } catch {
        /* cancelled */
      }
    } else active.reject(new Error(error ?? 'relay_disconnected'))
  }
  const revoke = (status: OfficeRelayStatus = 'offline', close = true) => {
    generation += 1
    finishRequest('relay_disconnected')
    pairingId = undefined
    sessionId = undefined
    capability = undefined
    negotiatedCapabilities = []
    for (const pending of pendingDiagnostics.values())
      pending.reject(new Error('diagnostic_unavailable'))
    pendingDiagnostics.clear()
    terminalRequestIds.clear()
    terminalRequestOrder.length = 0
    if (pairingTimer !== undefined) clearTimeout(pairingTimer)
    pairingTimer = undefined
    const activeSocket = socket
    socket = undefined
    if (close && activeSocket && activeSocket.readyState <= 1) activeSocket.close()
    settleConnect?.()
    settleConnect = undefined
    publish({ status })
  }
  const protocolFailure = () => revoke('offline')
  const send = (value: Record<string, unknown>) => {
    if (!socket || socket.readyState !== 1) throw new Error('relay_disconnected')
    const encoded = JSON.stringify(value)
    if (encoder.encode(encoded).byteLength > MAX_REQUEST_BYTES + 4096)
      throw new Error('relay_request_too_large')
    socket.send(encoded)
  }

  const handleFrame = (data: unknown, epoch: number) => {
    if (epoch !== generation || typeof data !== 'string') return protocolFailure()
    const frameBytes = encoder.encode(data).byteLength
    if (frameBytes > MAX_RELAY_FRAME_BYTES) return protocolFailure()
    let frame: Record<string, unknown>
    try {
      const parsed: unknown = JSON.parse(data)
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) throw new Error()
      frame = parsed as Record<string, unknown>
    } catch {
      return protocolFailure()
    }
    if (frame.version !== protocolVersion || typeof frame.type !== 'string')
      return protocolFailure()

    if (frame.type === 'office.diagnostic.accepted') {
      if (
        protocolVersion !== 2 ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        !exactKeys(frame, ['version', 'type', 'event_id']) ||
        typeof frame.event_id !== 'string' ||
        !pendingDiagnostics.has(frame.event_id)
      )
        return protocolFailure()
      const pending = pendingDiagnostics.get(frame.event_id)!
      pendingDiagnostics.delete(frame.event_id)
      pending.resolve()
      return
    }
    if (
      frame.type === 'relay.error' &&
      exactKeys(frame, ['version', 'type', 'code']) &&
      typeof frame.code === 'string' &&
      DIAGNOSTIC_ERROR_CODES.has(frame.code) &&
      pendingDiagnostics.size > 0
    ) {
      const oldest = pendingDiagnostics.entries().next().value as
        [string, { reject(error: Error): void }] | undefined
      if (oldest) {
        pendingDiagnostics.delete(oldest[0])
        oldest[1].reject(new Error(frame.code))
      }
      return
    }
    if (
      frame.type === 'relay.error' &&
      exactKeys(frame, ['version', 'type', 'code']) &&
      frame.code === 'session_expired'
    ) {
      revoke('expired')
      return
    }

    if (frame.type === 'office.created') {
      if (
        state.status !== 'connecting' ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        !exactKeys(frame, ['version', 'type', 'pairing_id', 'verification_code', 'expires_in']) ||
        !opaque(frame.pairing_id) ||
        typeof frame.verification_code !== 'string' ||
        !/^\d{6}$/.test(frame.verification_code) ||
        !expiry(frame.expires_in, 120)
      )
        return protocolFailure()
      pairingId = frame.pairing_id
      if (pairingTimer !== undefined) clearTimeout(pairingTimer)
      pairingTimer = setTimeout(() => revoke('expired'), frame.expires_in * 1000)
      publish({ status: 'pending', verificationCode: frame.verification_code })
      return
    }
    if (frame.type === 'office.approved') {
      const approvedKeys = ['version', 'type', 'session_id', 'capability', 'expires_in']
      if (protocolVersion === 2) approvedKeys.push('capabilities')
      const approvedCapabilities =
        protocolVersion === 2 &&
        Array.isArray(frame.capabilities) &&
        frame.capabilities.length > 0 &&
        frame.capabilities.every(
          (value, index, values) =>
            typeof value === 'string' &&
            requestedCapabilities.includes(value as OfficeRelayCapability) &&
            values.indexOf(value) === index,
        )
          ? (frame.capabilities as OfficeRelayCapability[])
          : protocolVersion === 1
            ? (['agent.v1'] as OfficeRelayCapability[])
            : undefined
      if (
        (state.status !== 'pending' && state.status !== 'waiting_for_pc') ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        !pairingId ||
        !exactKeys(frame, approvedKeys) ||
        !opaque(frame.session_id) ||
        !opaque(frame.capability) ||
        !expiry(frame.expires_in, 1800) ||
        !approvedCapabilities
      )
        return protocolFailure()
      sessionId = frame.session_id
      capability = frame.capability
      negotiatedCapabilities = approvedCapabilities
      pairingId = undefined
      if (pairingTimer !== undefined) clearTimeout(pairingTimer)
      pairingTimer = undefined
      settleConnect?.()
      settleConnect = undefined
      publish({
        status: 'connected',
        ...(protocolVersion === 2
          ? { capabilities: Object.freeze([...negotiatedCapabilities]) }
          : {}),
      })
      return
    }
    if (
      frame.type === 'office.rejected' ||
      frame.type === 'office.expired' ||
      frame.type === 'office.pc_offline'
    ) {
      if (frameBytes > MAX_CONTROL_FRAME_BYTES || !exactKeys(frame, ['version', 'type']))
        return protocolFailure()
      const status =
        frame.type === 'office.rejected'
          ? 'rejected'
          : frame.type === 'office.expired'
            ? 'expired'
            : 'waiting_for_pc'
      if (status === 'waiting_for_pc') publish({ status, verificationCode: state.verificationCode })
      else revoke(status)
      return
    }
    if (
      typeof frame.request_id === 'string' &&
      frame.session_id === sessionId &&
      terminalRequestIds.has(frame.request_id) &&
      ['relay.start', 'relay.chunk', 'relay.done', 'relay.error'].includes(String(frame.type))
    )
      return
    const active = request
    if (!active || frame.request_id !== active.id || frame.session_id !== sessionId)
      return protocolFailure()
    if (frame.type === 'relay.start') {
      if (
        active.controller ||
        !exactKeys(frame, [
          'version',
          'type',
          'session_id',
          'request_id',
          'status',
          'content_type',
        ]) ||
        !Number.isSafeInteger(frame.status) ||
        (frame.status as number) < 200 ||
        (frame.status as number) > 599 ||
        frame.status === 204 ||
        frame.status === 205 ||
        frame.status === 304 ||
        typeof frame.content_type !== 'string' ||
        frame.content_type.length < 1 ||
        frame.content_type.length > 128
      )
        return protocolFailure()
      const body = new ReadableStream<Uint8Array>({
        start(controller) {
          active.controller = controller
        },
        cancel() {
          if (request?.id === active.id) {
            try {
              send({
                version: protocolVersion,
                type: 'office.cancel',
                session_id: sessionId,
                capability,
                request_id: active.id,
              })
            } catch {
              /* revoked */
            }
            finishRequest('relay_cancelled')
          }
        },
      })
      try {
        const response = new Response(body, {
          status: frame.status as number,
          headers: { 'content-type': frame.content_type },
        })
        active.responseResolved = true
        active.resolve(response)
      } catch {
        protocolFailure()
      }
      return
    }
    if (frame.type === 'relay.chunk') {
      if (
        !active.responseResolved ||
        !active.controller ||
        !exactKeys(frame, ['version', 'type', 'session_id', 'request_id', 'sequence', 'data']) ||
        frame.sequence !== active.sequence ||
        typeof frame.data !== 'string'
      )
        return protocolFailure()
      if (!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(frame.data))
        return protocolFailure()
      let chunk: Uint8Array
      try {
        const decoded = atob(frame.data)
        if (btoa(decoded) !== frame.data) return protocolFailure()
        chunk = Uint8Array.from(decoded, (character) => character.charCodeAt(0))
      } catch {
        return protocolFailure()
      }
      if (
        chunk.byteLength > MAX_CHUNK_BYTES ||
        active.bytes + chunk.byteLength > MAX_RESPONSE_BYTES
      )
        return protocolFailure()
      active.sequence += 1
      active.bytes += chunk.byteLength
      active.controller.enqueue(chunk)
      return
    }
    if (frame.type === 'relay.done') {
      if (
        !active.responseResolved ||
        !active.controller ||
        !exactKeys(frame, ['version', 'type', 'session_id', 'request_id'])
      )
        return protocolFailure()
      finishRequest()
      return
    }
    if (frame.type === 'relay.error') {
      if (
        !exactKeys(frame, ['version', 'type', 'session_id', 'request_id', 'code']) ||
        !opaque(frame.code)
      )
        return protocolFailure()
      finishRequest('relay_error')
      return
    }
    protocolFailure()
  }

  const api: OfficeRelaySession = {
    snapshot: () => state,
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    connect(host) {
      revoke('offline')
      if (host === 'unknown') return Promise.resolve()
      const epoch = generation
      publish({ status: 'connecting' })
      return new Promise<void>((resolve) => {
        settleConnect = resolve
        let opened: RelayWebSocket
        try {
          opened = createSocket(OFFICE_RELAY_URL)
        } catch {
          revoke('offline')
          return
        }
        socket = opened
        opened.onopen = () => {
          if (epoch !== generation) return
          try {
            send({
              version: protocolVersion,
              type: 'office.create',
              host: hostLabels[host],
              ...(protocolVersion === 2 ? { capabilities: requestedCapabilities } : {}),
            })
          } catch {
            protocolFailure()
          }
        }
        opened.onmessage = (event) => handleFrame(event.data, epoch)
        opened.onerror = () => {
          if (epoch === generation) revoke('offline')
        }
        opened.onclose = () => {
          if (epoch === generation) revoke('offline', false)
        }
      })
    },
    disconnect() {
      revoke('offline')
    },
    async authenticatedFetch(path, init) {
      if (
        path !== MESSAGES_PATH ||
        !socket ||
        !sessionId ||
        !capability ||
        state.status !== 'connected'
      )
        throw new Error('relay_disconnected')
      if (request) throw new Error('relay_busy')
      if (init.method !== 'POST' || typeof init.body !== 'string')
        throw new Error('relay_invalid_request')
      if (encoder.encode(init.body).byteLength > MAX_REQUEST_BYTES)
        throw new Error('relay_request_too_large')
      let parsedBody: unknown
      try {
        parsedBody = JSON.parse(init.body)
      } catch {
        throw new Error('relay_invalid_request')
      }
      if (!parsedBody || typeof parsedBody !== 'object' || Array.isArray(parsedBody))
        throw new Error('relay_invalid_request')
      return api.capabilityFetch('agent.v1', parsedBody, init.signal ?? undefined)
    },
    async capabilityFetch(capabilityName, parsedBody, signal) {
      if (
        !socket ||
        !sessionId ||
        !capability ||
        state.status !== 'connected' ||
        !negotiatedCapabilities.includes(capabilityName)
      )
        throw new Error('relay_capability_unavailable')
      if (request) throw new Error('relay_busy')
      if (!parsedBody || typeof parsedBody !== 'object' || Array.isArray(parsedBody))
        throw new Error('relay_invalid_request')
      const bodyBytes = encoder.encode(JSON.stringify(parsedBody)).byteLength
      if (bodyBytes > MAX_REQUEST_BYTES) throw new Error('relay_request_too_large')
      const id = randomUUID()
      return new Promise<Response>((resolve, reject) => {
        const timer = setTimeout(() => {
          if (request?.id !== id) return
          try {
            send({
              version: protocolVersion,
              type: 'office.cancel',
              session_id: sessionId,
              capability,
              request_id: id,
            })
          } catch {
            /* revoked */
          }
          finishRequest('relay_timeout')
        }, REQUEST_TIMEOUT_MS)
        request = { id, sequence: 0, bytes: 0, responseResolved: false, resolve, reject, timer }
        if (signal) {
          const abort = () => {
            if (request?.id !== id) return
            try {
              send({
                version: protocolVersion,
                type: 'office.cancel',
                session_id: sessionId,
                capability,
                request_id: id,
              })
            } catch {
              /* revoked */
            }
            finishRequest('relay_cancelled')
          }
          request.abort = abort
          request.signal = signal
          if (signal.aborted) abort()
          else signal.addEventListener('abort', abort, { once: true })
        }
        if (request?.id === id) {
          try {
            send({
              version: protocolVersion,
              type: 'office.request',
              session_id: sessionId,
              capability,
              request_id: id,
              ...(protocolVersion === 2 ? { capability_name: capabilityName } : {}),
              body: parsedBody,
            })
          } catch {
            protocolFailure()
          }
        }
      })
    },
    sendDiagnostic(event) {
      if (
        protocolVersion !== 2 ||
        !socket ||
        !sessionId ||
        !capability ||
        state.status !== 'connected' ||
        pendingDiagnostics.size >= MAX_PENDING_DIAGNOSTICS ||
        pendingDiagnostics.has(event.event_id)
      )
        return Promise.reject(new Error('diagnostic_unavailable'))
      const safeEvent: OfficeDiagnosticEvent = {
        event_id: event.event_id,
        trace_id: event.trace_id,
        timestamp_ms: event.timestamp_ms,
        host: event.host,
        platform: event.platform,
        build: event.build,
        tool: event.tool,
        phase: event.phase,
        outcome: event.outcome,
        error_code:
          event.error_code === 'invalid_tool_input'
            ? 'agent_run_failed'
            : event.error_code === 'office_concurrent_change' ||
                event.error_code === 'office_state_uncertain'
              ? // Relay v2 does not yet advertise the local transaction-detail vocabulary.
                'office_verify_failed'
              : event.error_code.startsWith('office_recovery_failed:word_')
                ? 'office_recovery_failed'
                : event.error_code,
        ...(event.office_error_code ? { office_error_code: event.office_error_code } : {}),
        ...(event.office_error_name ? { office_error_name: event.office_error_name } : {}),
        ...(event.office_error_location
          ? { office_error_location: event.office_error_location }
          : {}),
        duration_ms: event.duration_ms,
        requirement_sets: { ...event.requirement_sets },
      }
      const wireFrame = {
        version: 2,
        type: 'office.diagnostic',
        session_id: sessionId,
        capability,
        ...safeEvent,
      }
      let serialized: string
      try {
        serialized = JSON.stringify(wireFrame)
      } catch {
        return Promise.reject(new Error('diagnostic_invalid'))
      }
      if (encoder.encode(serialized).byteLength > MAX_DIAGNOSTIC_EVENT_BYTES)
        return Promise.reject(new Error('diagnostic_too_large'))
      let resolve!: () => void
      let reject!: (error: Error) => void
      const result = new Promise<void>((next, fail) => {
        resolve = next
        reject = fail
      })
      pendingDiagnostics.set(event.event_id, { resolve, reject })
      try {
        send(wireFrame)
      } catch {
        pendingDiagnostics.delete(event.event_id)
        reject(new Error('diagnostic_unavailable'))
      }
      return result
    },
  }
  return api
}
