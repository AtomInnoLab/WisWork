import type { OfficeHost } from '../office-document.js'

export const PC_BRIDGE_ENDPOINT = 'http://127.0.0.1:43127'
const PAIRINGS_PATH = '/v1/office/pairings'
const MESSAGES_PATH = '/v1/office/messages'

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
    signal.addEventListener(
      'abort',
      () => {
        clearTimeout(timer)
        reject(new DOMException('Aborted', 'AbortError'))
      },
      { once: true },
    )
  })

export function createPcBridgeSession(dependencies: Dependencies = {}): PcBridgeSession {
  const endpoint = validateEndpoint(dependencies.endpoint ?? PC_BRIDGE_ENDPOINT)
  const fetcher = dependencies.fetch ?? globalThis.fetch
  const delay = dependencies.delay ?? sleep
  const now = dependencies.now ?? Date.now
  const listeners = new Set<() => void>()
  let state: PcBridgeSnapshot = { status: 'offline' }
  let capability: string | undefined
  let controller: AbortController | undefined

  const publish = (status: PcBridgeStatus) => {
    state = { status }
    listeners.forEach((listener) => listener())
  }
  const clear = (status: PcBridgeStatus) => {
    capability = undefined
    controller?.abort()
    controller = undefined
    publish(status)
  }

  return {
    snapshot: () => state,
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    async connect(host) {
      if (host === 'unknown') {
        clear('offline')
        return
      }
      clear('offline')
      const active = new AbortController()
      controller = active
      try {
        const created = await fetcher(`${endpoint}${PAIRINGS_PATH}`, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ host_label: hostLabels[host] }),
          signal: active.signal,
        })
        if (created.status === 401) {
          clear('signed_out')
          return
        }
        if (created.status !== 202) {
          clear(created.status === 403 ? 'signed_out' : 'offline')
          return
        }
        const pairing = (await created.json()) as {
          pairing_id?: unknown
          polling_secret?: unknown
          expires_in?: unknown
        }
        if (
          typeof pairing.pairing_id !== 'string' ||
          typeof pairing.polling_secret !== 'string' ||
          typeof pairing.expires_in !== 'number'
        ) {
          clear('offline')
          return
        }
        publish('pending')
        const deadline = now() + Math.min(pairing.expires_in * 1000, 120_000)
        let pause = 500
        while (!active.signal.aborted && now() < deadline) {
          const response = await fetcher(
            `${endpoint}${PAIRINGS_PATH}/${encodeURIComponent(pairing.pairing_id)}`,
            {
              method: 'GET',
              headers: { authorization: `Pairing ${pairing.polling_secret}` },
              signal: active.signal,
            },
          )
          if (response.status === 401 || response.status === 403) {
            clear('signed_out')
            return
          }
          if (!response.ok) {
            clear('offline')
            return
          }
          const result = (await response.json()) as { status?: unknown; capability?: unknown }
          if (result.status === 'approved' && typeof result.capability === 'string') {
            capability = result.capability
            controller = undefined
            publish('connected')
            return
          }
          if (
            result.status === 'rejected' ||
            result.status === 'expired' ||
            result.status === 'signed_out'
          ) {
            clear(result.status)
            return
          }
          if (result.status !== 'pending') {
            clear('offline')
            return
          }
          await delay(pause, active.signal)
          pause = Math.min(Math.ceil(pause * 1.5), 3_000)
        }
        if (!active.signal.aborted) clear('expired')
      } catch (error) {
        if (!active.signal.aborted) clear('offline')
      }
    },
    disconnect() {
      clear('offline')
    },
    async authenticatedFetch(path, init) {
      if (path !== MESSAGES_PATH || !capability) throw new Error('bridge_disconnected')
      const headers = new Headers(init.headers)
      headers.set('authorization', `Bridge ${capability}`)
      try {
        const response = await fetcher(`${endpoint}${MESSAGES_PATH}`, { ...init, headers })
        if (response.status === 401) clear('signed_out')
        return response
      } catch (error) {
        clear('offline')
        throw new Error('bridge_offline', { cause: error })
      }
    },
  }
}
