import type { OfficePairingRequest, OfficeRelayStatus } from '../shared/home-api'
import type { OfficeRelayClient } from './office-relay-client'

const DEFAULT_MAX_CLIENTS = 12

interface ChildEvents {
  onPending(pairing: OfficePairingRequest): void
  onPendingExpired(pairingId: string): void
  onStatus(status: OfficeRelayStatus): void
}

interface Slot {
  client: OfficeRelayClient | null
  status: OfficeRelayStatus
  pairings: Set<string>
}

export type OfficeRelayPool = OfficeRelayClient

export function createOfficeRelayPool(options: {
  createClient(events: ChildEvents): OfficeRelayClient
  onPending(pairing: OfficePairingRequest): void
  onPendingExpired?: (pairingId: string) => void
  onStatus?: (status: OfficeRelayStatus) => void
  maxClients?: number
}): OfficeRelayPool {
  const maximum = options.maxClients ?? DEFAULT_MAX_CLIENTS
  if (!Number.isSafeInteger(maximum) || maximum < 1 || maximum > DEFAULT_MAX_CLIENTS)
    throw new Error('invalid_relay_pool_capacity')

  const slots = new Set<Slot>()
  const pairingOwners = new Map<string, Slot>()
  let fallbackStatus: OfficeRelayStatus = 'disconnected'
  let publishedStatus: OfficeRelayStatus = 'disconnected'
  let revoking = false

  const aggregate = (): OfficeRelayStatus => {
    const statuses = [...slots].map((slot) => slot.status)
    if (statuses.includes('paired')) return 'paired'
    if (statuses.includes('awaiting_approval')) return 'awaiting_approval'
    if (statuses.includes('claiming')) return 'claiming'
    if (statuses.includes('connecting')) return 'connecting'
    return fallbackStatus
  }

  const publish = () => {
    const next = aggregate()
    if (next === publishedStatus) return
    publishedStatus = next
    options.onStatus?.(next)
  }

  const remove = (slot: Slot) => {
    if (!slots.delete(slot)) return
    for (const pairingId of slot.pairings) pairingOwners.delete(pairingId)
    slot.pairings.clear()
  }

  const revokeAll = (reason: string) => {
    if (revoking) return
    revoking = true
    const children = [...slots]
    slots.clear()
    pairingOwners.clear()
    fallbackStatus = `disconnected:${reason}` as OfficeRelayStatus
    for (const slot of children) slot.client?.revoke(reason)
    revoking = false
    publish()
  }

  return {
    async claim(code) {
      if (!/^\d{6}$/.test(code)) throw new Error('invalid_verification_code')
      if (slots.size >= maximum) throw new Error('relay_capacity_exceeded')

      // Reserve before constructing the client or awaiting auth/socket work.
      const slot: Slot = { client: null, status: 'connecting', pairings: new Set() }
      slots.add(slot)
      publish()
      try {
        slot.client = options.createClient({
          onPending(pairing) {
            if (!slots.has(slot) || pairingOwners.has(pairing.pairingId)) {
              slot.client?.revoke('protocol_violation')
              return
            }
            slot.pairings.add(pairing.pairingId)
            pairingOwners.set(pairing.pairingId, slot)
            options.onPending(pairing)
          },
          onPendingExpired(pairingId) {
            if (pairingOwners.get(pairingId) !== slot) return
            pairingOwners.delete(pairingId)
            slot.pairings.delete(pairingId)
            options.onPendingExpired?.(pairingId)
          },
          onStatus(status) {
            if (!slots.has(slot)) return
            // A freshly constructed single-session client clears its empty
            // predecessor state before its first claim. The pool has already
            // reserved this child, so this is not a child disconnect.
            if (status === 'disconnected:new_claim') return
            slot.status = status
            if (status === 'disconnected:auth_required') {
              revokeAll('auth_required')
              return
            }
            if (status === 'disconnected' || status.startsWith('disconnected:')) {
              fallbackStatus = status
              remove(slot)
            }
            publish()
          },
        })
        await slot.client.claim(code)
      } catch (error) {
        remove(slot)
        publish()
        throw error
      }
    },
    async approve(pairingId) {
      const owner = pairingOwners.get(pairingId)
      if (!owner?.client) return false
      const approved = await owner.client.approve(pairingId)
      if (approved) {
        pairingOwners.delete(pairingId)
        owner.pairings.delete(pairingId)
      }
      return approved
    },
    reject(pairingId) {
      const owner = pairingOwners.get(pairingId)
      return owner?.client ? owner.client.reject(pairingId) : false
    },
    listPending() {
      return [...slots].flatMap((slot) => slot.client?.listPending() ?? [])
    },
    status: aggregate,
    revoke: revokeAll,
  }
}
