import type { OfficeHost } from '../office-document.js'
import { officeBridgeEndpoints } from '../../build-config.js'
import {
  parseOfficeEnhancedStatement,
  type OfficeEnhancedStatement,
} from '../agent/enhanced-session.js'

export const PC_BRIDGE_ENDPOINTS = officeBridgeEndpoints(import.meta.env)
const PAIRINGS_PATH = '/v1/office/pairings'
const MESSAGES_PATH = '/v1/office/messages'
const CONNECT_DEADLINE_MS = 120_000
const MAX_PROTOCOL_BODY_BYTES = 4096
const MAX_OPAQUE_LENGTH = 512
const HEALTH_PATH = '/v1/office/health'
const DISCOVERY_BATCH_SIZE = 8
const PROBE_TIMEOUT_MS = 400

export type PcBridgeStatus =
  'offline' | 'connecting' | 'signed_out' | 'pending' | 'rejected' | 'expired' | 'connected'
export interface PcBridgeSnapshot {
  status: PcBridgeStatus
  verificationCode?: string
  enhanced?: OfficeEnhancedStatement
}
export interface PcBridgeSession {
  snapshot(): PcBridgeSnapshot
  subscribe(listener: () => void): () => void
  connect(host: OfficeHost): Promise<void>
  disconnect(): void
  authenticatedFetch(path: typeof MESSAGES_PATH, init: RequestInit): Promise<Response>
  setToolHandler(handler: PcBridgeToolHandler | undefined): void
  handleToolFrame(event: Record<string, unknown>, signal: AbortSignal): Promise<void>
}
type PcBridgeToolHandler = (call: {
  turnId: string
  callId: string
  generation: number
  toolName: string
  input: Record<string, unknown>
  signal: AbortSignal
}) => Promise<{ output: string; isError?: boolean }>
interface Dependencies {
  endpoint?: string
  endpoints?: readonly string[]
  fetch?: typeof globalThis.fetch
  delay?: (milliseconds: number, signal: AbortSignal) => Promise<void>
  now?: () => number
  batchSize?: number
  probeTimeoutMs?: number
}

function validateEndpoint(value: string): string {
  const url = new URL(value)
  if (
    url.protocol !== 'http:' ||
    url.hostname !== '127.0.0.1' ||
    url.href !== `http://127.0.0.1:${url.port}/` ||
    !url.port
  )
    throw new Error('invalid_bridge_endpoint')
  return url.origin
}

function validateEndpoints(values: readonly string[]): readonly string[] {
  if (values.length < 1 || values.length > 128) throw new Error('invalid_bridge_endpoints')
  const endpoints = values.map(validateEndpoint)
  if (new Set(endpoints).size !== endpoints.length) throw new Error('invalid_bridge_endpoints')
  return endpoints
}
const hostLabels: Record<Exclude<OfficeHost, 'unknown'>, string> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
}
const sleep = (milliseconds: number, signal: AbortSignal) =>
  new Promise<void>((resolve, reject) => {
    const timer = setTimeout(resolve, milliseconds)
    const abort = () => {
      clearTimeout(timer)
      reject(new DOMException('Aborted', 'AbortError'))
    }
    signal.addEventListener('abort', abort, { once: true })
  })
const validOpaque = (value: unknown): value is string =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= MAX_OPAQUE_LENGTH &&
  /^[A-Za-z0-9_-]+$/.test(value)
const validExpiry = (value: unknown, maximum: number): value is number =>
  typeof value === 'number' && Number.isFinite(value) && value > 0 && value <= maximum

function abortPromise(signal: AbortSignal): Promise<never> {
  return new Promise((_resolve, reject) => {
    const rejectAbort = () => reject(new DOMException('Aborted', 'AbortError'))
    if (signal.aborted) rejectAbort()
    else signal.addEventListener('abort', rejectAbort, { once: true })
  })
}

async function boundedJson(response: Response, signal: AbortSignal): Promise<unknown> {
  const declared = Number(response.headers.get('content-length'))
  if (Number.isFinite(declared) && declared > MAX_PROTOCOL_BODY_BYTES)
    throw new Error('invalid_protocol')
  if (!response.body) throw new Error('invalid_protocol')
  const reader = response.body.getReader()
  const chunks: Uint8Array[] = []
  let size = 0
  try {
    while (true) {
      const result = await Promise.race([reader.read(), abortPromise(signal)])
      if (result.done) break
      size += result.value.byteLength
      if (size > MAX_PROTOCOL_BODY_BYTES) throw new Error('invalid_protocol')
      chunks.push(result.value)
    }
    const bytes = new Uint8Array(size)
    let offset = 0
    for (const chunk of chunks) {
      bytes.set(chunk, offset)
      offset += chunk.byteLength
    }
    return JSON.parse(new TextDecoder().decode(bytes)) as unknown
  } finally {
    if (signal.aborted || size > MAX_PROTOCOL_BODY_BYTES)
      void reader.cancel().catch(() => undefined)
    reader.releaseLock()
  }
}

interface DiscoveryDependencies {
  fetch?: typeof globalThis.fetch
  signal?: AbortSignal
  batchSize?: number
  probeTimeoutMs?: number
}

const validHealth = (value: unknown): boolean => {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const record = value as Record<string, unknown>
  return (
    Object.keys(record).length === 2 &&
    record.service === 'wiswork-office-bridge' &&
    record.version === 1
  )
}

async function discoverBatch(
  endpoints: readonly string[],
  fetcher: typeof globalThis.fetch,
  timeoutMs: number,
  outerSignal?: AbortSignal,
): Promise<string | undefined> {
  const controllers = endpoints.map(() => new AbortController())
  const results: Array<boolean | undefined> = endpoints.map(() => undefined)
  let settled = false
  return new Promise((resolve) => {
    const finish = (endpoint?: string) => {
      if (settled) return
      settled = true
      controllers.forEach((controller) => controller.abort())
      outerSignal?.removeEventListener('abort', abort)
      resolve(endpoint)
    }
    const consider = () => {
      for (let index = 0; index < results.length; index += 1) {
        if (results[index] === undefined) return
        if (results[index]) return finish(endpoints[index])
      }
      finish()
    }
    const abort = () => finish()
    if (outerSignal?.aborted) return abort()
    outerSignal?.addEventListener('abort', abort, { once: true })
    endpoints.forEach((endpoint, index) => {
      const controller = controllers[index]!
      const timer = setTimeout(() => controller.abort(), timeoutMs)
      void (async () => {
        try {
          const response = await Promise.race([
            fetcher(`${endpoint}${HEALTH_PATH}`, {
              method: 'GET',
              signal: controller.signal,
              headers: { accept: 'application/json' },
            }),
            abortPromise(controller.signal),
          ])
          results[index] =
            response.status === 200 && validHealth(await boundedJson(response, controller.signal))
        } catch {
          results[index] = false
        } finally {
          clearTimeout(timer)
          consider()
        }
      })()
    })
  })
}

export async function discoverPcBridgeEndpoint(
  endpointValues: readonly string[],
  dependencies: DiscoveryDependencies = {},
): Promise<string | undefined> {
  const endpoints = validateEndpoints(endpointValues)
  const fetcher = dependencies.fetch ?? globalThis.fetch
  const batchSize = dependencies.batchSize ?? DISCOVERY_BATCH_SIZE
  const timeoutMs = dependencies.probeTimeoutMs ?? PROBE_TIMEOUT_MS
  if (!Number.isSafeInteger(batchSize) || batchSize < 1 || batchSize > 16)
    throw new Error('invalid_batch_size')
  if (!Number.isSafeInteger(timeoutMs) || timeoutMs < 1 || timeoutMs > 500)
    throw new Error('invalid_probe_timeout')
  for (
    let offset = 0;
    offset < endpoints.length && !dependencies.signal?.aborted;
    offset += batchSize
  ) {
    const found = await discoverBatch(
      endpoints.slice(offset, offset + batchSize),
      fetcher,
      timeoutMs,
      dependencies.signal,
    )
    if (found) return found
  }
  return undefined
}

export function createPcBridgeSession(dependencies: Dependencies = {}): PcBridgeSession {
  const fixedEndpoint = dependencies.endpoint ? validateEndpoint(dependencies.endpoint) : undefined
  const configuredEndpoints = validateEndpoints(dependencies.endpoints ?? PC_BRIDGE_ENDPOINTS)
  const fetcher = dependencies.fetch ?? globalThis.fetch
  const delay = dependencies.delay ?? sleep
  const now = dependencies.now ?? Date.now
  const listeners = new Set<() => void>()
  let state: PcBridgeSnapshot = { status: 'offline' }
  let capability: string | undefined
  let controller: AbortController | undefined
  let generation = 0
  let activeEndpoint: string | undefined
  let toolHandler: PcBridgeToolHandler | undefined
  const publish = (status: PcBridgeStatus) => {
    state = { status }
    listeners.forEach((listener) => listener())
  }
  const invalidate = (status: PcBridgeStatus): number => {
    generation += 1
    capability = undefined
    activeEndpoint = undefined
    controller?.abort()
    controller = undefined
    publish(status)
    return generation
  }
  const active = (epoch: number, operation: AbortController) =>
    generation === epoch && controller === operation && !operation.signal.aborted
  const current = (epoch: number, operation: AbortController) =>
    generation === epoch && controller === operation
  const finish = (epoch: number, operation: AbortController, status: PcBridgeStatus) => {
    if (current(epoch, operation)) invalidate(status)
  }

  return {
    snapshot: () => state,
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    async connect(host) {
      const epoch = invalidate('offline')
      if (host === 'unknown') return
      const operation = new AbortController()
      controller = operation
      publish('connecting')
      const timer = setTimeout(() => operation.abort(), CONNECT_DEADLINE_MS)
      let stage: 'create' | 'poll' = 'create'
      try {
        const endpoint =
          fixedEndpoint ??
          (await discoverPcBridgeEndpoint(configuredEndpoints, {
            fetch: fetcher,
            signal: operation.signal,
            batchSize: dependencies.batchSize,
            probeTimeoutMs: dependencies.probeTimeoutMs,
          }))
        if (!endpoint) {
          finish(epoch, operation, 'offline')
          return
        }
        if (!active(epoch, operation)) return
        activeEndpoint = endpoint
        const created = await Promise.race([
          fetcher(`${endpoint}${PAIRINGS_PATH}`, {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ host_label: hostLabels[host] }),
            signal: operation.signal,
          }),
          abortPromise(operation.signal),
        ])
        if (!active(epoch, operation)) return
        if (created.status === 401 || created.status === 403) {
          finish(epoch, operation, 'signed_out')
          return
        }
        if (created.status !== 202) {
          finish(epoch, operation, 'offline')
          return
        }
        const pairing = (await boundedJson(created, operation.signal)) as Record<string, unknown>
        if (!active(epoch, operation)) return
        if (
          !validOpaque(pairing.pairing_id) ||
          !validOpaque(pairing.polling_secret) ||
          typeof pairing.verification_code !== 'string' ||
          !/^\d{6}$/.test(pairing.verification_code) ||
          !validExpiry(pairing.expires_in, 120)
        ) {
          finish(epoch, operation, 'offline')
          return
        }
        state = { status: 'pending', verificationCode: pairing.verification_code }
        listeners.forEach((listener) => listener())
        const protocolDeadline = now() + pairing.expires_in * 1000
        let pause = 500
        stage = 'poll'
        while (active(epoch, operation) && now() < protocolDeadline) {
          const response = await Promise.race([
            fetcher(`${endpoint}${PAIRINGS_PATH}/${pairing.pairing_id}`, {
              method: 'GET',
              headers: { authorization: `Pairing ${pairing.polling_secret}` },
              signal: operation.signal,
            }),
            abortPromise(operation.signal),
          ])
          if (!active(epoch, operation)) return
          if (response.status === 401 || response.status === 403) {
            finish(epoch, operation, 'signed_out')
            return
          }
          if (!response.ok) {
            finish(epoch, operation, 'offline')
            return
          }
          const result = (await boundedJson(response, operation.signal)) as Record<string, unknown>
          if (!active(epoch, operation)) return
          if (result.status === 'approved') {
            if (!validOpaque(result.capability) || !validExpiry(result.expires_in, 3600)) {
              finish(epoch, operation, 'offline')
              return
            }
            let enhanced: OfficeEnhancedStatement | undefined
            if (result.enhanced !== undefined) {
              try {
                enhanced = parseOfficeEnhancedStatement(result.enhanced)
              } catch {
                finish(epoch, operation, 'offline')
                return
              }
              if (enhanced.host !== `office-${host}` || enhanced.expires_at <= now()) {
                finish(epoch, operation, 'offline')
                return
              }
            }
            capability = result.capability
            controller = undefined
            clearTimeout(timer)
            state = { status: 'connected', ...(enhanced ? { enhanced } : {}) }
            listeners.forEach((listener) => listener())
            return
          }
          if (
            result.status === 'rejected' ||
            result.status === 'expired' ||
            result.status === 'signed_out'
          ) {
            finish(epoch, operation, result.status)
            return
          }
          if (result.status !== 'pending') {
            finish(epoch, operation, 'offline')
            return
          }
          await delay(Math.min(pause, Math.max(0, protocolDeadline - now())), operation.signal)
          pause = Math.min(Math.ceil(pause * 1.5), 3_000)
        }
        finish(epoch, operation, 'expired')
      } catch {
        finish(
          epoch,
          operation,
          operation.signal.aborted && stage === 'poll' ? 'expired' : 'offline',
        )
      } finally {
        clearTimeout(timer)
      }
    },
    disconnect() {
      invalidate('offline')
    },
    async authenticatedFetch(path, init) {
      if (path !== MESSAGES_PATH || !capability || !activeEndpoint)
        throw new Error('bridge_disconnected')
      const requestGeneration = generation
      const requestCapability = capability
      const headers = new Headers(init.headers)
      headers.set('authorization', `Bridge ${requestCapability}`)
      try {
        const response = await fetcher(`${activeEndpoint}${MESSAGES_PATH}`, { ...init, headers })
        if (
          response.status === 401 &&
          generation === requestGeneration &&
          capability === requestCapability
        )
          invalidate('signed_out')
        return response
      } catch (error) {
        throw new Error('bridge_offline', { cause: error })
      }
    },
    setToolHandler(handler) {
      toolHandler = handler
    },
    async handleToolFrame(event, signal) {
      if (!toolHandler || !capability || !activeEndpoint || !state.enhanced)
        throw new Error('bridge_disconnected')
      const requestId = event.request_id,
        turnId = event.turn_id,
        callId = event.call_id,
        toolName = event.tool_name
      if (
        ![requestId, turnId, callId, toolName].every(validOpaque) ||
        event.generation !== state.enhanced.session_generation ||
        !event.input ||
        typeof event.input !== 'object' ||
        Array.isArray(event.input)
      )
        throw new Error('invalid_tool_frame')
      const result = await toolHandler({
        turnId: turnId as string,
        callId: callId as string,
        generation: event.generation as number,
        toolName: toolName as string,
        input: event.input as Record<string, unknown>,
        signal,
      })
      const response = await fetcher(`${activeEndpoint}/v1/office/tools/results`, {
        method: 'POST',
        signal,
        headers: { authorization: `Bridge ${capability}`, 'content-type': 'application/json' },
        body: JSON.stringify({
          request_id: requestId,
          turn_id: turnId,
          call_id: callId,
          generation: event.generation,
          output: result.output,
          is_error: result.isError === true,
        }),
      })
      if (response.status !== 204) throw new Error('tool_result_rejected')
    },
  }
}
