import type { OfficeHost } from '../office-document.js'
import { officeBridgeEndpoint } from '../../build-config.js'

export const PC_BRIDGE_ENDPOINT = officeBridgeEndpoint(import.meta.env)
const PAIRINGS_PATH = '/v1/office/pairings'
const MESSAGES_PATH = '/v1/office/messages'
const CONNECT_DEADLINE_MS = 120_000
const MAX_PROTOCOL_BODY_BYTES = 4096
const MAX_OPAQUE_LENGTH = 512

export type PcBridgeStatus =
  'offline' | 'signed_out' | 'pending' | 'rejected' | 'expired' | 'connected'
export interface PcBridgeSnapshot {
  status: PcBridgeStatus
}
export interface PcBridgeSession {
  snapshot(): PcBridgeSnapshot
  subscribe(listener: () => void): () => void
  connect(host: OfficeHost): Promise<void>
  disconnect(): void
  authenticatedFetch(path: typeof MESSAGES_PATH, init: RequestInit): Promise<Response>
}
interface Dependencies {
  endpoint?: string
  fetch?: typeof globalThis.fetch
  delay?: (milliseconds: number, signal: AbortSignal) => Promise<void>
  now?: () => number
}

function validateEndpoint(value: string): string {
  const url = new URL(value)
  if (url.href !== `${PC_BRIDGE_ENDPOINT}/`) throw new Error('invalid_bridge_endpoint')
  return PC_BRIDGE_ENDPOINT
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

export function createPcBridgeSession(dependencies: Dependencies = {}): PcBridgeSession {
  const endpoint = validateEndpoint(dependencies.endpoint ?? PC_BRIDGE_ENDPOINT)
  const fetcher = dependencies.fetch ?? globalThis.fetch
  const delay = dependencies.delay ?? sleep
  const now = dependencies.now ?? Date.now
  const listeners = new Set<() => void>()
  let state: PcBridgeSnapshot = { status: 'offline' }
  let capability: string | undefined
  let controller: AbortController | undefined
  let generation = 0
  const publish = (status: PcBridgeStatus) => {
    state = { status }
    listeners.forEach((listener) => listener())
  }
  const invalidate = (status: PcBridgeStatus): number => {
    generation += 1
    capability = undefined
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
      const timer = setTimeout(() => operation.abort(), CONNECT_DEADLINE_MS)
      let stage: 'create' | 'poll' = 'create'
      try {
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
          !validExpiry(pairing.expires_in, 120)
        ) {
          finish(epoch, operation, 'offline')
          return
        }
        publish('pending')
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
            capability = result.capability
            controller = undefined
            clearTimeout(timer)
            publish('connected')
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
      if (path !== MESSAGES_PATH || !capability) throw new Error('bridge_disconnected')
      const requestGeneration = generation
      const requestCapability = capability
      const headers = new Headers(init.headers)
      headers.set('authorization', `Bridge ${requestCapability}`)
      try {
        const response = await fetcher(`${endpoint}${MESSAGES_PATH}`, { ...init, headers })
        if (
          response.status === 401 &&
          generation === requestGeneration &&
          capability === requestCapability
        )
          invalidate('signed_out')
        return response
      } catch (error) {
        if (generation === requestGeneration && capability === requestCapability)
          invalidate('offline')
        throw new Error('bridge_offline', { cause: error })
      }
    },
  }
}
