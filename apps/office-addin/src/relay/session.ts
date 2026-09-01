import type { OfficeHost } from '../office-document.js'
import type { OfficeDiagnosticEvent } from '../diagnostics/office-diagnostics.js'
import { officeTransportMode, type OfficeTransportMode } from '../../build-config.js'
import {
  OFFICE_RELAY_ORIGIN,
  PAIRING_RESUME_FEATURE,
  createBrowserOfficeBindingStore,
  type OfficeBindingEnrollment,
  type OfficeBindingStore,
  type OfficeStoredBinding,
} from './binding-store.js'
import {
  parseOfficeEnhancedStatement,
  type OfficeEnhancedStatement,
} from '../agent/enhanced-session.js'

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
  | 'offline'
  | 'connecting'
  | 'reconnecting'
  | 'pending'
  | 'waiting_for_pc'
  | 'rejected'
  | 'expired'
  | 'connected'

export interface OfficeRelaySnapshot {
  status: OfficeRelayStatus
  verificationCode?: string
  capabilities?: readonly string[]
  remembered?: false
  enhanced?: OfficeEnhancedStatement
}

export interface OfficeRelaySession {
  snapshot(): OfficeRelaySnapshot
  subscribe(listener: () => void): () => void
  connect(host: OfficeHost): Promise<void>
  disconnect(): void
  forget(): Promise<void>
  authenticatedFetch(path: typeof MESSAGES_PATH, init: RequestInit): Promise<Response>
  capabilityFetch(
    capability: OfficeRelayCapability,
    body: unknown,
    signal?: AbortSignal,
  ): Promise<Response>
  sendDiagnostic(event: OfficeDiagnosticEvent): Promise<void>
  setToolHandler?(handler: OfficeRelayToolHandler | undefined): void
}

export interface OfficeRelayToolCall {
  readonly turnId: string
  readonly callId: string
  readonly generation: number
  readonly toolName: string
  readonly input: Record<string, unknown>
  readonly signal: AbortSignal
}
export type OfficeRelayToolHandler = (
  call: OfficeRelayToolCall,
) => Promise<{ output: string; isError?: boolean }>

export type OfficeRelayCapability =
  'agent.v1' | 'web-search.v1' | 'web-fetch.v1' | 'image-search.v1'

export interface OfficeBindingInvalidation {
  readonly origin: typeof OFFICE_RELAY_ORIGIN
  readonly host: Exclude<OfficeHost, 'unknown'>
  readonly bindingId: string
}

export interface OfficeBindingInvalidationChannel {
  subscribe(listener: (message: OfficeBindingInvalidation) => void): () => void
  broadcast(message: OfficeBindingInvalidation): void
}

export interface OfficeRelaySessionDependencies {
  createSocket?: (url: string) => RelayWebSocket
  randomUUID?: () => string
  capabilities?: readonly OfficeRelayCapability[]
  persistentPairing?: boolean
  bindingStore?: OfficeBindingStore
  schedule?: (callback: () => void, delay: number) => unknown
  cancelSchedule?: (handle: unknown) => void
  random?: () => number
  bindingInvalidationChannel?: OfficeBindingInvalidationChannel
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
const matchesHost = (value: unknown): value is Exclude<OfficeHost, 'unknown'> =>
  value === 'word' || value === 'excel' || value === 'powerpoint'
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
const resumeFeature = (value: unknown): value is [typeof PAIRING_RESUME_FEATURE] =>
  Array.isArray(value) && value.length === 1 && value[0] === PAIRING_RESUME_FEATURE

function browserSocket(url: string): RelayWebSocket {
  return new WebSocket(url) as unknown as RelayWebSocket
}

interface BindingBroadcastChannel {
  onmessage: ((event: MessageEvent<unknown>) => void) | null
  postMessage(message: unknown): void
}

export function createBrowserOfficeBindingInvalidationChannel(
  createChannel: (name: string) => BindingBroadcastChannel = (name) => new BroadcastChannel(name),
): OfficeBindingInvalidationChannel | undefined {
  let channel: BindingBroadcastChannel
  try {
    channel = createChannel('wiswork-office-pairing-v1')
  } catch {
    return undefined
  }
  const listeners = new Set<(message: OfficeBindingInvalidation) => void>()
  channel.onmessage = (event) => {
    const value = event.data
    if (!value || typeof value !== 'object' || Array.isArray(value)) return
    const message = value as Record<string, unknown>
    if (
      !exactKeys(message, ['version', 'type', 'origin', 'host', 'binding_id']) ||
      message.version !== 1 ||
      message.type !== 'binding.forgotten' ||
      message.origin !== OFFICE_RELAY_ORIGIN ||
      !matchesHost(message.host) ||
      !opaque(message.binding_id)
    )
      return
    const invalidation = Object.freeze({
      origin: OFFICE_RELAY_ORIGIN,
      host: message.host,
      bindingId: message.binding_id,
    })
    listeners.forEach((listener) => listener(invalidation))
  }
  return Object.freeze({
    subscribe(listener: (message: OfficeBindingInvalidation) => void) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    broadcast(message: OfficeBindingInvalidation) {
      if (
        message.origin !== OFFICE_RELAY_ORIGIN ||
        !matchesHost(message.host) ||
        !opaque(message.bindingId)
      )
        return
      channel.postMessage({
        version: 1,
        type: 'binding.forgotten',
        origin: OFFICE_RELAY_ORIGIN,
        host: message.host,
        binding_id: message.bindingId,
      })
    },
  })
}

export function createOfficeRelaySession(
  dependencies: OfficeRelaySessionDependencies = {},
): OfficeRelaySession {
  const createSocket = dependencies.createSocket ?? browserSocket
  const randomUUID = dependencies.randomUUID ?? (() => crypto.randomUUID())
  const requestedCapabilities = [...(dependencies.capabilities ?? [])]
  const persistentPairing = dependencies.persistentPairing !== false
  const configuredBindingStore = persistentPairing ? dependencies.bindingStore : undefined
  const browserBindingStore =
    persistentPairing && !configuredBindingStore ? createBrowserOfficeBindingStore() : undefined
  const browserInvalidationChannel =
    persistentPairing && browserBindingStore
      ? createBrowserOfficeBindingInvalidationChannel()
      : undefined
  const bindingInvalidationChannel =
    (persistentPairing ? dependencies.bindingInvalidationChannel : undefined) ??
    (configuredBindingStore ? undefined : browserInvalidationChannel)
  const bindingStore =
    persistentPairing && requestedCapabilities.length > 0
      ? (configuredBindingStore ?? (bindingInvalidationChannel ? browserBindingStore : undefined))
      : undefined
  const schedule = dependencies.schedule ?? ((callback, delay) => setTimeout(callback, delay))
  const cancelSchedule =
    dependencies.cancelSchedule ??
    ((handle) => clearTimeout(handle as ReturnType<typeof setTimeout>))
  const random = dependencies.random ?? Math.random
  const protocolVersion = requestedCapabilities.length > 0 ? 2 : 1
  const listeners = new Set<() => void>()
  let state: OfficeRelaySnapshot = { status: 'offline' }
  let socket: RelayWebSocket | undefined
  let host: Exclude<OfficeHost, 'unknown'> | undefined
  let connectionMode: 'legacy' | 'enroll' | 'resume' = 'legacy'
  let storedBinding: OfficeStoredBinding | undefined
  let enrollment: OfficeBindingEnrollment | undefined
  let bindingOffer:
    | {
        bindingId: string
        capabilities: OfficeRelayCapability[]
        status: 'staging' | 'staged' | 'committed' | 'aborting'
      }
    | undefined
  let enhancedFallbackUsed = false
  let resumeFallbackUsed = false
  let resumeRecognized = false
  let explicitlyDisconnected = false
  let retryAttempt = 0
  let retryTimer: unknown
  let challengeInFlight = false
  let remembered: false | undefined
  let abortedEnrollmentApproval = false
  let pairingId: string | undefined
  let sessionId: string | undefined
  let capability: string | undefined
  let negotiatedCapabilities: OfficeRelayCapability[] = []
  let request: ActiveRequest | undefined
  let toolHandler: OfficeRelayToolHandler | undefined
  let activeTool: { requestId: string; callId: string; controller: AbortController } | undefined
  let enhancedStatement: OfficeEnhancedStatement | undefined
  let runtimeGeneration = -1
  let generation = 0
  let settleConnect: (() => void) | undefined
  let pairingTimer: ReturnType<typeof setTimeout> | undefined
  let unsubscribeInvalidation: (() => void) | undefined
  let bindingDeletionBarrier: Promise<void> | undefined
  let pendingForgetInvalidation: OfficeBindingInvalidation | undefined
  let pendingBindingCleanup: OfficeBindingInvalidation | undefined
  let bindingInvalidationEpoch = 0
  const invalidatedBindings = new Map<string, number>()
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
    if (activeTool?.requestId === active.id) {
      activeTool.controller.abort()
      activeTool = undefined
    }
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
  const revoke = (status: OfficeRelayStatus = 'offline', close = true, settle = true) => {
    generation += 1
    finishRequest('relay_disconnected')
    activeTool?.controller.abort()
    activeTool = undefined
    enhancedStatement = undefined
    runtimeGeneration = -1
    pairingId = undefined
    sessionId = undefined
    capability = undefined
    negotiatedCapabilities = []
    challengeInFlight = false
    remembered = undefined
    abortedEnrollmentApproval = false
    for (const pending of pendingDiagnostics.values())
      pending.reject(new Error('diagnostic_unavailable'))
    pendingDiagnostics.clear()
    terminalRequestIds.clear()
    terminalRequestOrder.length = 0
    if (pairingTimer !== undefined) clearTimeout(pairingTimer)
    pairingTimer = undefined
    const activeSocket = socket
    socket = undefined
    if (activeSocket) {
      activeSocket.onopen = null
      activeSocket.onmessage = null
      activeSocket.onerror = null
      activeSocket.onclose = null
    }
    if (close && activeSocket && activeSocket.readyState <= 1) activeSocket.close()
    if (settle) {
      settleConnect?.()
      settleConnect = undefined
    }
    publish({ status })
  }
  const cancelRetry = () => {
    if (retryTimer === undefined) return
    cancelSchedule(retryTimer)
    retryTimer = undefined
  }
  const markBindingInvalid = (bindingId: string) => {
    bindingInvalidationEpoch += 1
    invalidatedBindings.delete(bindingId)
    invalidatedBindings.set(bindingId, bindingInvalidationEpoch)
    while (invalidatedBindings.size > 32)
      invalidatedBindings.delete(invalidatedBindings.keys().next().value!)
  }
  const beginBindingDeletion = (
    targetHost: Exclude<OfficeHost, 'unknown'> | undefined,
    bindingId: string | undefined,
    afterCommit?: () => void,
  ) => {
    const previous = bindingDeletionBarrier
    const deletion = (async () => {
      if (previous) await previous.catch(() => undefined)
      if (targetHost) await bindingStore?.forget(targetHost, bindingId)
      afterCommit?.()
    })()
    bindingDeletionBarrier = deletion
    void deletion
      .finally(() => {
        if (bindingDeletionBarrier === deletion) bindingDeletionBarrier = undefined
      })
      .catch(() => undefined)
    return deletion
  }
  const deleteActivatedBinding = async (binding: OfficeStoredBinding, epoch: number) => {
    const cleanup = {
      origin: OFFICE_RELAY_ORIGIN,
      host: binding.host,
      bindingId: binding.bindingId,
    } as const
    pendingBindingCleanup = cleanup
    markBindingInvalid(binding.bindingId)
    storedBinding = undefined
    try {
      await beginBindingDeletion(binding.host, binding.bindingId, () => {
        if (pendingBindingCleanup?.bindingId === binding.bindingId) {
          bindingInvalidationChannel?.broadcast(cleanup)
          pendingBindingCleanup = undefined
        }
      })
      return epoch === generation
    } catch {
      if (epoch === generation) {
        explicitlyDisconnected = true
        revoke('offline')
      }
      return false
    }
  }
  const send = (value: Record<string, unknown>) => {
    if (!socket || socket.readyState !== 1) throw new Error('relay_disconnected')
    const encoded = JSON.stringify(value)
    if (encoder.encode(encoded).byteLength > MAX_REQUEST_BYTES + 4096)
      throw new Error('relay_request_too_large')
    socket.send(encoded)
  }

  const scheduleResume = () => {
    if (explicitlyDisconnected || !storedBinding || !host) return revoke('offline')
    revoke('reconnecting', true, false)
    cancelRetry()
    const base = Math.min(30_000, 500 * 2 ** Math.min(retryAttempt, 6))
    const delay = Math.min(30_000, Math.round(base * (0.75 + 0.5 * random())))
    retryAttempt += 1
    retryTimer = schedule(() => {
      retryTimer = undefined
      if (explicitlyDisconnected || !storedBinding || !host) return
      startAttempt('resume')
    }, delay)
  }

  const abandonEnrollment = () => {
    const abandoned = enrollment
    const abandonedBindingId = bindingOffer?.bindingId
    enrollment = undefined
    bindingOffer = undefined
    if (abandoned) void bindingStore?.abort(abandoned, abandonedBindingId).catch(() => undefined)
  }

  const fallbackToLegacy = () => {
    if (enhancedFallbackUsed || explicitlyDisconnected || !host) return revoke('offline')
    enhancedFallbackUsed = true
    abandonEnrollment()
    startAttempt('legacy')
  }

  const fallbackStoredBindingToLegacy = () => {
    if (resumeFallbackUsed || explicitlyDisconnected || !host) return revoke('offline')
    resumeFallbackUsed = true
    startAttempt('legacy')
  }

  const handleDisconnect = (epoch: number) => {
    if (epoch !== generation || explicitlyDisconnected) return
    if (connectionMode === 'enroll') return fallbackToLegacy()
    if (connectionMode === 'resume' && storedBinding) return scheduleResume()
    revoke('offline', false)
  }

  const startAttempt = (mode: 'legacy' | 'enroll' | 'resume') => {
    if (!host || explicitlyDisconnected) return
    revoke(state.status === 'reconnecting' ? 'reconnecting' : 'connecting', true, false)
    connectionMode = mode
    if (mode === 'resume') resumeRecognized = false
    const epoch = generation
    let opened: RelayWebSocket
    try {
      opened = createSocket(OFFICE_RELAY_URL)
    } catch {
      if (mode === 'enroll') fallbackToLegacy()
      else if (mode === 'resume') scheduleResume()
      else revoke('offline')
      return
    }
    socket = opened
    opened.onopen = () => {
      if (epoch !== generation || !host) return
      try {
        if (mode === 'resume' && storedBinding) {
          send({
            version: 2,
            type: 'office.resume',
            binding_id: storedBinding.bindingId,
            host: hostLabels[host],
            capabilities: requestedCapabilities,
          })
        } else if (mode === 'enroll' && enrollment) {
          send({
            version: 2,
            type: 'office.create',
            host: hostLabels[host],
            capabilities: requestedCapabilities,
            features: [PAIRING_RESUME_FEATURE],
            binding_public_key: enrollment.publicKey,
          })
        } else {
          send({
            version: protocolVersion,
            type: 'office.create',
            host: hostLabels[host],
            ...(protocolVersion === 2 ? { capabilities: requestedCapabilities } : {}),
          })
        }
      } catch {
        handleDisconnect(epoch)
      }
    }
    let frameQueue = Promise.resolve()
    opened.onmessage = (event) => {
      frameQueue = frameQueue
        .then(() => handleFrame(event.data, epoch))
        .catch(() => {
          if (epoch === generation) protocolFailure()
        })
    }
    opened.onerror = () => handleDisconnect(epoch)
    opened.onclose = () => handleDisconnect(epoch)
  }

  const prepareEnrollment = async (intent: number) => {
    if (!bindingStore || !host) return startAttempt('legacy')
    try {
      const prepared = await bindingStore.createEnrollment(host, requestedCapabilities)
      if (intent !== generation || explicitlyDisconnected) {
        await bindingStore.abort(prepared).catch(() => undefined)
        return
      }
      enrollment = prepared
      bindingOffer = undefined
      startAttempt('enroll')
    } catch {
      if (intent === generation && !explicitlyDisconnected) startAttempt('legacy')
    }
  }

  const beginConnection = async (intent: number) => {
    if (!bindingStore || !host || protocolVersion !== 2) return startAttempt('legacy')
    const connectingHost = host
    try {
      if (pendingBindingCleanup?.host === connectingHost) {
        const cleanup = pendingBindingCleanup
        await bindingStore.forget(connectingHost, cleanup.bindingId)
        if (intent !== generation || explicitlyDisconnected) return
        bindingInvalidationChannel?.broadcast(cleanup)
        pendingBindingCleanup = undefined
      } else {
        // A previous taskpane may have durably blocked a binding before its delete failed.
        await bindingStore.forget(connectingHost)
        if (intent !== generation || explicitlyDisconnected) return
      }
    } catch {
      if (intent === generation) {
        explicitlyDisconnected = true
        revoke('offline')
      }
      return
    }
    try {
      const loaded = await bindingStore.load(connectingHost, requestedCapabilities)
      if (intent !== generation || explicitlyDisconnected) return
      if (loaded) {
        if (invalidatedBindings.has(loaded.bindingId)) await prepareEnrollment(intent)
        else {
          storedBinding = loaded
          startAttempt('resume')
        }
      } else {
        await prepareEnrollment(intent)
      }
    } catch {
      if (intent === generation && !explicitlyDisconnected) startAttempt('legacy')
    }
  }

  const invalidateBindingAndEnroll = () => {
    const invalidated = storedBinding
    if (invalidated) markBindingInvalid(invalidated.bindingId)
    storedBinding = undefined
    enrollment = undefined
    bindingOffer = undefined
    retryAttempt = 0
    cancelRetry()
    revoke('connecting', true, false)
    const intent = generation
    void beginBindingDeletion(host, invalidated?.bindingId)
      .then(async () => {
        if (intent === generation && !explicitlyDisconnected) await prepareEnrollment(intent)
      })
      .catch(() => {
        if (intent === generation) revoke('offline')
      })
  }

  const protocolFailure = () => {
    if (connectionMode === 'enroll') fallbackToLegacy()
    else if (connectionMode === 'resume' && storedBinding) scheduleResume()
    else revoke('offline')
  }

  const subscribeInvalidation = () => {
    if (!bindingInvalidationChannel || unsubscribeInvalidation) return
    unsubscribeInvalidation = bindingInvalidationChannel.subscribe((message) => {
      if (message.origin !== OFFICE_RELAY_ORIGIN || !host || message.host !== host) return
      markBindingInvalid(message.bindingId)
      if (message.bindingId !== storedBinding?.bindingId) return
      explicitlyDisconnected = true
      cancelRetry()
      storedBinding = undefined
      enrollment = undefined
      bindingOffer = undefined
      unsubscribeInvalidation?.()
      unsubscribeInvalidation = undefined
      revoke('offline')
    })
  }

  const handleFrame = async (data: unknown, epoch: number) => {
    if (epoch !== generation) return
    if (typeof data !== 'string') return protocolFailure()
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

    if (
      frame.type === 'relay.error' &&
      exactKeys(frame, ['version', 'type', 'code']) &&
      typeof frame.code === 'string'
    ) {
      if (connectionMode === 'enroll' && frame.code === 'invalid_frame') {
        fallbackToLegacy()
        return
      }
      if (connectionMode === 'resume') {
        if (
          [
            'binding_unavailable',
            'binding_revoked',
            'invalid_proof',
            'capability_not_negotiated',
          ].includes(frame.code)
        ) {
          invalidateBindingAndEnroll()
          return
        }
        if (!resumeRecognized && ['invalid_frame', 'unknown_type'].includes(frame.code)) {
          fallbackStoredBindingToLegacy()
          return
        }
        if (
          [
            'session_revoked',
            'challenge_expired',
            'resume_rate_limited',
            'resume_limit',
            'peer_unavailable',
          ].includes(frame.code)
        ) {
          scheduleResume()
          return
        }
      }
    }

    if (frame.type === 'office.challenge') {
      if (
        connectionMode !== 'resume' ||
        (state.status !== 'connecting' && state.status !== 'reconnecting') ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        challengeInFlight ||
        !storedBinding ||
        !exactKeys(frame, ['version', 'type', 'binding_id', 'challenge', 'expires_in']) ||
        frame.binding_id !== storedBinding.bindingId ||
        !opaque(frame.challenge) ||
        !expiry(frame.expires_in, 30)
      )
        return protocolFailure()
      resumeRecognized = true
      challengeInFlight = true
      try {
        const signature = await bindingStore?.sign(storedBinding, frame.challenge)
        if (epoch !== generation || !signature) return
        send({
          version: 2,
          type: 'office.prove',
          binding_id: storedBinding.bindingId,
          challenge: frame.challenge,
          signature,
        })
      } catch {
        if (epoch === generation) invalidateBindingAndEnroll()
      }
      return
    }

    if (frame.type === 'office.waiting_for_pc') {
      if (
        connectionMode !== 'resume' ||
        (state.status !== 'connecting' && state.status !== 'reconnecting') ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        !exactKeys(frame, ['version', 'type'])
      )
        return protocolFailure()
      publish({ status: 'waiting_for_pc' })
      return
    }

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
      if (connectionMode === 'resume' && storedBinding) scheduleResume()
      else revoke('expired')
      return
    }

    if (frame.type === 'office.created') {
      const expectedKeys = ['version', 'type', 'pairing_id', 'verification_code', 'expires_in']
      if (connectionMode === 'enroll') expectedKeys.push('features')
      if (
        (connectionMode !== 'legacy' && connectionMode !== 'enroll') ||
        (state.status !== 'connecting' && state.status !== 'reconnecting') ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        !exactKeys(frame, expectedKeys) ||
        !opaque(frame.pairing_id) ||
        typeof frame.verification_code !== 'string' ||
        !/^\d{6}$/.test(frame.verification_code) ||
        !expiry(frame.expires_in, 120) ||
        (connectionMode === 'enroll' && !resumeFeature(frame.features))
      )
        return protocolFailure()
      pairingId = frame.pairing_id
      if (pairingTimer !== undefined) clearTimeout(pairingTimer)
      pairingTimer = setTimeout(() => {
        abandonEnrollment()
        revoke('expired')
      }, frame.expires_in * 1000)
      publish({ status: 'pending', verificationCode: frame.verification_code })
      return
    }
    if (frame.type === 'office.binding_offer') {
      const offeredCapabilities =
        Array.isArray(frame.capabilities) &&
        frame.capabilities.length > 0 &&
        frame.capabilities.every(
          (value, index, values) =>
            typeof value === 'string' &&
            requestedCapabilities.includes(value as OfficeRelayCapability) &&
            values.indexOf(value) === index,
        )
          ? (frame.capabilities as OfficeRelayCapability[])
          : undefined
      if (
        connectionMode !== 'enroll' ||
        state.status !== 'pending' ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        !exactKeys(frame, [
          'version',
          'type',
          'pairing_id',
          'binding_id',
          'capabilities',
          'features',
        ]) ||
        frame.pairing_id !== pairingId ||
        !opaque(frame.binding_id) ||
        !offeredCapabilities ||
        !resumeFeature(frame.features) ||
        !bindingStore ||
        !enrollment ||
        bindingOffer
      )
        return protocolFailure()
      const offeredEnrollment = enrollment
      const offeredBindingId = frame.binding_id
      bindingOffer = {
        bindingId: offeredBindingId,
        capabilities: [...offeredCapabilities],
        status: 'staging',
      }
      try {
        await bindingStore.stage(offeredEnrollment, offeredBindingId, offeredCapabilities)
        if (epoch !== generation) {
          await bindingStore.abort(offeredEnrollment, offeredBindingId).catch(() => undefined)
          return
        }
        bindingOffer = {
          bindingId: offeredBindingId,
          capabilities: [...offeredCapabilities],
          status: 'staged',
        }
        send({
          version: 2,
          type: 'office.binding_ready',
          pairing_id: pairingId,
          binding_id: offeredBindingId,
        })
      } catch {
        await bindingStore.abort(offeredEnrollment, offeredBindingId).catch(() => undefined)
        if (epoch !== generation) return
        enrollment = undefined
        bindingOffer = {
          bindingId: offeredBindingId,
          capabilities: [...offeredCapabilities],
          status: 'aborting',
        }
        try {
          send({
            version: 2,
            type: 'office.binding_abort',
            pairing_id: pairingId,
            binding_id: offeredBindingId,
          })
        } catch {
          handleDisconnect(epoch)
        }
      }
      return
    }
    if (frame.type === 'office.binding_commit') {
      if (
        connectionMode !== 'enroll' ||
        state.status !== 'pending' ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        !exactKeys(frame, ['version', 'type', 'pairing_id', 'binding_id']) ||
        frame.pairing_id !== pairingId ||
        frame.binding_id !== bindingOffer?.bindingId ||
        bindingOffer?.status !== 'staged' ||
        !bindingStore ||
        !enrollment
      )
        return protocolFailure()
      const committingEnrollment = enrollment
      const committingOffer = bindingOffer
      let activated: OfficeStoredBinding
      try {
        activated = await bindingStore.activate(
          committingEnrollment,
          committingOffer.bindingId,
          committingOffer.capabilities,
        )
      } catch {
        await bindingStore
          .abort(committingEnrollment, committingOffer.bindingId)
          .catch(() => undefined)
        if (epoch !== generation) return
        enrollment = undefined
        storedBinding = undefined
        bindingOffer = { ...committingOffer, status: 'aborting' }
        try {
          send({
            version: 2,
            type: 'office.binding_abort',
            pairing_id: pairingId,
            binding_id: committingOffer.bindingId,
          })
        } catch {
          handleDisconnect(epoch)
        }
        return
      }
      if (epoch !== generation) return
      storedBinding = activated
      enrollment = undefined
      bindingOffer = { ...committingOffer, status: 'committed' }
      try {
        send({
          version: 2,
          type: 'office.binding_committed',
          pairing_id: pairingId,
          binding_id: committingOffer.bindingId,
        })
      } catch {
        await deleteActivatedBinding(activated, epoch)
        if (epoch !== generation) return
        enrollment = undefined
        storedBinding = undefined
        bindingOffer = { ...committingOffer, status: 'aborting' }
        explicitlyDisconnected = true
        revoke('offline')
      }
      return
    }
    if (frame.type === 'office.binding_aborted') {
      if (
        connectionMode !== 'enroll' ||
        state.status !== 'pending' ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        !exactKeys(frame, ['version', 'type', 'pairing_id', 'binding_id']) ||
        frame.pairing_id !== pairingId ||
        frame.binding_id !== bindingOffer?.bindingId ||
        !bindingOffer
      )
        return protocolFailure()
      if (bindingOffer.status === 'staged' && enrollment && bindingStore) {
        await bindingStore.abort(enrollment, bindingOffer.bindingId).catch(() => undefined)
        if (epoch !== generation) return
      }
      if (bindingOffer.status === 'committed' && storedBinding && bindingStore) {
        const deleted = await deleteActivatedBinding(storedBinding, epoch)
        if (!deleted || epoch !== generation) return
      }
      enrollment = undefined
      bindingOffer = undefined
      connectionMode = 'legacy'
      remembered = false
      abortedEnrollmentApproval = true
      return
    }
    if (frame.type === 'office.approved') {
      const approvedKeys = ['version', 'type', 'session_id', 'capability', 'expires_in']
      if (protocolVersion === 2) approvedKeys.push('capabilities')
      const committedOffer = bindingOffer?.status === 'committed' ? bindingOffer : undefined
      const enhancedApproval =
        connectionMode === 'enroll' &&
        Boolean(committedOffer) &&
        resumeFeature(frame.features) &&
        frame.binding_id === committedOffer?.bindingId &&
        exactKeys(frame, [...approvedKeys, 'features', 'binding_id'])
      const abortedShortApproval =
        connectionMode === 'legacy' &&
        remembered === false &&
        abortedEnrollmentApproval &&
        Array.isArray(frame.features) &&
        frame.features.length === 0 &&
        exactKeys(frame, [...approvedKeys, 'features'])
      const standardApproval =
        connectionMode !== 'enroll' && !abortedEnrollmentApproval && exactKeys(frame, approvedKeys)
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
        (connectionMode === 'resume'
          ? state.status !== 'connecting' &&
            state.status !== 'reconnecting' &&
            state.status !== 'waiting_for_pc'
          : state.status !== 'pending' && state.status !== 'waiting_for_pc') ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        (connectionMode !== 'resume' && !pairingId) ||
        (!standardApproval && !enhancedApproval && !abortedShortApproval) ||
        !opaque(frame.session_id) ||
        !opaque(frame.capability) ||
        !expiry(frame.expires_in, 1800) ||
        !approvedCapabilities ||
        (enhancedApproval && !storedBinding) ||
        (committedOffer &&
          enhancedApproval &&
          (approvedCapabilities.length !== committedOffer.capabilities.length ||
            approvedCapabilities.some(
              (value, index) => value !== committedOffer.capabilities[index],
            )))
      )
        return protocolFailure()
      if (enhancedApproval && storedBinding) {
        bindingOffer = undefined
        connectionMode = 'resume'
      }
      sessionId = frame.session_id
      capability = frame.capability
      negotiatedCapabilities = approvedCapabilities
      pairingId = undefined
      if (pairingTimer !== undefined) clearTimeout(pairingTimer)
      pairingTimer = undefined
      settleConnect?.()
      settleConnect = undefined
      retryAttempt = 0
      abortedEnrollmentApproval = false
      publish({
        status: 'connected',
        ...(protocolVersion === 2
          ? { capabilities: Object.freeze([...negotiatedCapabilities]) }
          : {}),
        ...(remembered === false ? { remembered: false as const } : {}),
      })
      return
    }
    if (frame.type === 'relay.session_state') {
      if (
        state.status !== 'connected' ||
        !sessionId ||
        frameBytes > MAX_CONTROL_FRAME_BYTES ||
        !exactKeys(frame, ['version', 'type', 'session_id', 'generation', 'enhanced']) ||
        frame.session_id !== sessionId ||
        !Number.isSafeInteger(frame.generation) ||
        Number(frame.generation) <= runtimeGeneration
      )
        return protocolFailure()
      let parsed: OfficeEnhancedStatement | undefined
      if (frame.enhanced !== null) {
        try {
          parsed = parseOfficeEnhancedStatement(frame.enhanced)
        } catch {
          return protocolFailure()
        }
        if (
          !host ||
          parsed.host !== `office-${host}` ||
          parsed.expires_at <= Date.now() ||
          parsed.session_generation !== frame.generation
        )
          return protocolFailure()
      } else if (frame.generation !== 0) return protocolFailure()
      runtimeGeneration = Number(frame.generation)
      enhancedStatement = parsed
      publish({ ...state, ...(parsed ? { enhanced: parsed } : {}) })
      return
    }
    if (frame.type === 'relay.tool_call') {
      if (
        state.status !== 'connected' ||
        !sessionId ||
        !capability ||
        !request ||
        activeTool ||
        !enhancedStatement ||
        !toolHandler ||
        frameBytes > MAX_REQUEST_BYTES + MAX_CONTROL_FRAME_BYTES ||
        !exactKeys(frame, [
          'version',
          'type',
          'session_id',
          'request_id',
          'turn_id',
          'call_id',
          'generation',
          'tool_name',
          'input',
        ]) ||
        frame.session_id !== sessionId ||
        frame.request_id !== request.id ||
        !opaque(frame.turn_id) ||
        !opaque(frame.call_id) ||
        frame.generation !== enhancedStatement.session_generation ||
        typeof frame.tool_name !== 'string' ||
        !/^[A-Za-z0-9_-]{1,128}$/.test(frame.tool_name) ||
        !frame.input ||
        typeof frame.input !== 'object' ||
        Array.isArray(frame.input)
      )
        return protocolFailure()
      const controller = new AbortController()
      activeTool = {
        requestId: frame.request_id as string,
        callId: frame.call_id as string,
        controller,
      }
      void toolHandler({
        turnId: frame.turn_id as string,
        callId: frame.call_id as string,
        generation: frame.generation as number,
        toolName: frame.tool_name,
        input: frame.input as Record<string, unknown>,
        signal: controller.signal,
      })
        .then((result) => {
          if (
            !activeTool ||
            activeTool.controller !== controller ||
            controller.signal.aborted ||
            !sessionId ||
            !capability
          )
            return
          if (
            typeof result.output !== 'string' ||
            encoder.encode(result.output).byteLength > MAX_RESPONSE_BYTES
          )
            throw new Error('tool_result_too_large')
          send({
            version: 2,
            type: 'office.tool_result',
            session_id: sessionId,
            capability,
            request_id: frame.request_id,
            turn_id: frame.turn_id,
            call_id: frame.call_id,
            generation: frame.generation,
            output: result.output,
            is_error: result.isError === true,
          })
          activeTool = undefined
        })
        .catch(() => {
          if (
            !activeTool ||
            activeTool.controller !== controller ||
            controller.signal.aborted ||
            !sessionId ||
            !capability
          )
            return
          send({
            version: 2,
            type: 'office.tool_result',
            session_id: sessionId,
            capability,
            request_id: frame.request_id,
            turn_id: frame.turn_id,
            call_id: frame.call_id,
            generation: frame.generation,
            output: 'tool_execution_failed',
            is_error: true,
          })
          activeTool = undefined
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
      else {
        abandonEnrollment()
        revoke(status)
      }
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
    connect(nextHost) {
      const abandoned = enrollment
      const abandonedBindingId = bindingOffer?.bindingId
      explicitlyDisconnected = true
      cancelRetry()
      revoke('offline')
      storedBinding = undefined
      enrollment = undefined
      bindingOffer = undefined
      if (abandoned) void bindingStore?.abort(abandoned, abandonedBindingId).catch(() => undefined)
      enhancedFallbackUsed = false
      resumeFallbackUsed = false
      retryAttempt = 0
      host = nextHost === 'unknown' ? undefined : nextHost
      if (!host) return Promise.resolve()
      explicitlyDisconnected = false
      publish({ status: 'connecting' })
      return new Promise<void>((resolve) => {
        settleConnect = resolve
        const intent = generation
        const deletion = bindingDeletionBarrier
        if (!deletion) {
          subscribeInvalidation()
          void beginConnection(intent)
          return
        }
        void (async () => {
          try {
            await deletion
          } catch {
            if (intent === generation) {
              explicitlyDisconnected = true
              revoke('offline')
            }
            return
          }
          if (intent !== generation || explicitlyDisconnected || !host) return
          subscribeInvalidation()
          await beginConnection(intent)
        })()
      })
    },
    disconnect() {
      const abandoned = enrollment
      const abandonedBindingId = bindingOffer?.bindingId
      explicitlyDisconnected = true
      cancelRetry()
      unsubscribeInvalidation?.()
      unsubscribeInvalidation = undefined
      enrollment = undefined
      bindingOffer = undefined
      if (abandoned) void bindingStore?.abort(abandoned, abandonedBindingId).catch(() => undefined)
      revoke('offline')
    },
    async forget() {
      const abandoned = enrollment
      const forgotten = storedBinding ?? pendingForgetInvalidation
      const targetHost = forgotten?.host ?? enrollment?.host ?? host
      const targetBindingId = forgotten?.bindingId ?? bindingOffer?.bindingId
      if (targetHost && targetBindingId) {
        markBindingInvalid(targetBindingId)
        pendingForgetInvalidation = {
          origin: OFFICE_RELAY_ORIGIN,
          host: targetHost,
          bindingId: targetBindingId,
        }
      }
      explicitlyDisconnected = true
      cancelRetry()
      unsubscribeInvalidation?.()
      unsubscribeInvalidation = undefined
      storedBinding = undefined
      enrollment = undefined
      bindingOffer = undefined
      revoke('offline')
      if (abandoned) await bindingStore?.abort(abandoned, targetBindingId)
      if (!targetBindingId) return
      await beginBindingDeletion(targetHost, targetBindingId, () => {
        if (pendingForgetInvalidation) {
          bindingInvalidationChannel?.broadcast({
            origin: OFFICE_RELAY_ORIGIN,
            host: pendingForgetInvalidation.host,
            bindingId: pendingForgetInvalidation.bindingId,
          })
          pendingForgetInvalidation = undefined
        }
      })
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
    setToolHandler(handler) {
      toolHandler = handler
      if (!handler) {
        activeTool?.controller.abort()
        activeTool = undefined
      }
    },
  }
  return api
}
