import type { OfficePairingRequest, OfficeRelayStatus } from '../shared/home-api'
import type { OfficeRelayBinding } from './office-relay-binding-store'
import type { OfficeRelayClient } from './office-relay-client'

const DEFAULT_MAX_CLIENTS = 12

interface ChildEvents {
  onPending(pairing: OfficePairingRequest): void
  onPendingExpired(pairingId: string): void
  onStatus(status: OfficeRelayStatus): void
  onBinding(binding: OfficeRelayBinding): Promise<void>
  onBindingInvalidated(bindingId: string): void
}

interface Slot {
  client: OfficeRelayClient | null
  status: OfficeRelayStatus
  pairings: Set<string>
  binding: OfficeRelayBinding | null
  retryCount: number
  retryTimer: ReturnType<typeof setTimeout> | null
}

export interface OfficeRelayPool extends OfficeRelayClient {
  suspend(reason: string): void
  activate(): void
}

export function createOfficeRelayPool(options: {
  createClient(events: ChildEvents): OfficeRelayClient
  onPending(pairing: OfficePairingRequest): void
  onPendingExpired?: (pairingId: string) => void
  onStatus?: (status: OfficeRelayStatus) => void
  onBinding?: (binding: OfficeRelayBinding) => void | Promise<void>
  onBindingInvalidated?: (binding: OfficeRelayBinding) => void | Promise<void>
  onAuthRequired?: () => void | Promise<void>
  maxClients?: number
  baseReconnectMs?: number
  maxReconnectMs?: number
  random?: () => number
  setTimer?: typeof setTimeout
  clearTimer?: typeof clearTimeout
}): OfficeRelayPool {
  const maximum = options.maxClients ?? DEFAULT_MAX_CLIENTS
  if (!Number.isSafeInteger(maximum) || maximum < 1 || maximum > DEFAULT_MAX_CLIENTS)
    throw new Error('invalid_relay_pool_capacity')

  const slots = new Set<Slot>()
  const bindingOwners = new Map<string, Set<Slot>>()
  const pairingOwners = new Map<string, Slot>()
  const approving = new Set<string>()
  let fallbackStatus: OfficeRelayStatus = 'disconnected'
  let publishedStatus: OfficeRelayStatus = 'disconnected'
  let revoking = false
  let suspended = false
  const baseReconnectMs = options.baseReconnectMs ?? 1_000
  const maxReconnectMs = options.maxReconnectMs ?? 30_000
  const random = options.random ?? Math.random
  const setTimer = options.setTimer ?? setTimeout
  const clearTimer = options.clearTimer ?? clearTimeout

  const aggregate = (): OfficeRelayStatus => {
    const statuses = [...slots].map((slot) => slot.status)
    if (statuses.includes('paired')) return 'paired'
    if (statuses.includes('awaiting_approval')) return 'awaiting_approval'
    if (statuses.includes('waiting_for_office')) return 'waiting_for_office'
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
    if (slot.retryTimer) clearTimer(slot.retryTimer)
    slot.retryTimer = null
    if (slot.binding) {
      const owners = bindingOwners.get(slot.binding.bindingId)
      owners?.delete(slot)
      if (owners?.size === 0) bindingOwners.delete(slot.binding.bindingId)
    }
    for (const pairingId of slot.pairings) {
      pairingOwners.delete(pairingId)
      approving.delete(pairingId)
    }
    slot.pairings.clear()
  }

  const addBindingOwner = (bindingId: string, slot: Slot) => {
    const owners = bindingOwners.get(bindingId) ?? new Set<Slot>()
    owners.add(slot)
    bindingOwners.set(bindingId, owners)
  }

  const retryable = (status: OfficeRelayStatus): boolean =>
    status === 'disconnected:network_error' ||
    status === 'disconnected:relay_closed' ||
    status === 'disconnected:session_expired' ||
    status === 'disconnected:peer_unavailable' ||
    status === 'disconnected:resume_rate_limited' ||
    status === 'disconnected:resume_limit'

  const scheduleReconnect = (slot: Slot) => {
    if (revoking || !slots.has(slot) || !slot.binding || !slot.client || slot.retryTimer) return
    const exponential = Math.min(maxReconnectMs, baseReconnectMs * 2 ** slot.retryCount)
    const delay = Math.min(maxReconnectMs, Math.round(exponential * (0.75 + random() * 0.5)))
    slot.retryCount += 1
    slot.retryTimer = setTimer(() => {
      slot.retryTimer = null
      if (revoking || !slots.has(slot) || !slot.binding || !slot.client) return
      void slot.client.resume(slot.binding).catch(() => scheduleReconnect(slot))
    }, delay)
  }

  let ensureWaitingResume: (binding: OfficeRelayBinding) => Promise<void>

  const initialize = (slot: Slot): OfficeRelayClient => {
    const client = options.createClient({
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
        approving.delete(pairingId)
        options.onPendingExpired?.(pairingId)
      },
      onStatus(status) {
        if (!slots.has(slot)) return
        if (
          status === 'disconnected:new_claim' ||
          status === 'disconnected:new_resume' ||
          status === 'disconnected:new_revocation'
        )
          return
        slot.status = status
        if (status === 'disconnected:auth_required') {
          void Promise.resolve(options.onAuthRequired?.()).catch(() => undefined)
          revokeAll('auth_required')
          return
        }
        if (status === 'paired') {
          slot.retryCount = 0
          for (const pairingId of slot.pairings) {
            pairingOwners.delete(pairingId)
            approving.delete(pairingId)
          }
          slot.pairings.clear()
          if (slot.binding) void ensureWaitingResume(slot.binding)
        }
        if (
          slot.binding &&
          (status === 'disconnected:binding_unavailable' ||
            status === 'disconnected:capability_not_negotiated')
        ) {
          const invalid = slot.binding
          remove(slot)
          void Promise.resolve(options.onBindingInvalidated?.(invalid)).catch(() => undefined)
        } else if (slot.binding && retryable(status)) {
          fallbackStatus = status
          scheduleReconnect(slot)
        } else if (status === 'disconnected' || status.startsWith('disconnected:')) {
          fallbackStatus = status
          remove(slot)
        }
        publish()
      },
      async onBinding(binding) {
        if (!slots.has(slot) || suspended) {
          slot.client?.revoke('binding_lifecycle_suspended')
          throw new Error('binding_lifecycle_suspended')
        }
        await options.onBinding?.(binding)
        if (!slots.has(slot) || suspended) throw new Error('binding_lifecycle_suspended')
        slot.binding = { ...binding, capabilities: [...binding.capabilities] }
        addBindingOwner(binding.bindingId, slot)
        if (slot.status === 'paired') void ensureWaitingResume(slot.binding)
      },
      onBindingInvalidated(bindingId) {
        if (!slot.binding || slot.binding.bindingId !== bindingId) return
        const invalid = slot.binding
        for (const owner of [...(bindingOwners.get(bindingId) ?? [])]) {
          if (owner !== slot) owner.client?.revoke('binding_unavailable')
          remove(owner)
        }
        void Promise.resolve(options.onBindingInvalidated?.(invalid)).catch(() => undefined)
        publish()
      },
    })
    slot.client = client
    return client
  }

  const startResumeSlot = async (binding: OfficeRelayBinding): Promise<void> => {
    if (suspended) throw new Error('relay_suspended')
    if (slots.size >= maximum) throw new Error('relay_capacity_exceeded')
    const slot: Slot = {
      client: null,
      status: 'connecting',
      pairings: new Set(),
      binding: { ...binding, capabilities: [...binding.capabilities] },
      retryCount: 0,
      retryTimer: null,
    }
    slots.add(slot)
    addBindingOwner(binding.bindingId, slot)
    publish()
    try {
      const client = initialize(slot)
      await client.resume({ ...binding, capabilities: [...binding.capabilities] })
    } catch (error) {
      if (slots.has(slot) && retryable(slot.status)) {
        scheduleReconnect(slot)
        return
      }
      remove(slot)
      publish()
      throw error
    }
  }

  ensureWaitingResume = async (binding) => {
    if (suspended || slots.size >= maximum) return
    const owners = bindingOwners.get(binding.bindingId)
    if (owners && [...owners].some((owner) => owner.status !== 'paired')) return
    await startResumeSlot(binding).catch(() => undefined)
  }

  const revokeAll = (reason: string) => {
    if (revoking) return
    revoking = true
    const children = [...slots]
    const pendingIds = [...pairingOwners.keys()]
    slots.clear()
    bindingOwners.clear()
    pairingOwners.clear()
    approving.clear()
    fallbackStatus =
      reason === 'binding_lifecycle'
        ? 'error:binding_lifecycle'
        : (`disconnected:${reason}` as OfficeRelayStatus)
    for (const pairingId of pendingIds) options.onPendingExpired?.(pairingId)
    for (const slot of children) {
      if (slot.retryTimer) clearTimer(slot.retryTimer)
      slot.retryTimer = null
      slot.client?.revoke(reason)
    }
    revoking = false
    publish()
  }

  return {
    async claim(code) {
      if (suspended) throw new Error('relay_suspended')
      if (!/^\d{6}$/.test(code)) throw new Error('invalid_verification_code')
      if (slots.size >= maximum) throw new Error('relay_capacity_exceeded')

      // Reserve before constructing the client or awaiting auth/socket work.
      const slot: Slot = {
        client: null,
        status: 'connecting',
        pairings: new Set(),
        binding: null,
        retryCount: 0,
        retryTimer: null,
      }
      slots.add(slot)
      publish()
      try {
        const client = initialize(slot)
        await client.claim(code)
      } catch (error) {
        remove(slot)
        publish()
        throw error
      }
    },
    async resume(binding) {
      if (suspended) throw new Error('relay_suspended')
      const owners = bindingOwners.get(binding.bindingId)
      if (owners && [...owners].some((owner) => owner.status !== 'paired')) return
      await startResumeSlot(binding)
    },
    async revokeBinding(bindingId) {
      const owners = [...(bindingOwners.get(bindingId) ?? [])]
      let slot = owners[0]
      let temporary = false
      if (!slot) {
        if (slots.size >= maximum) throw new Error('relay_capacity_exceeded')
        slot = {
          client: null,
          status: 'connecting',
          pairings: new Set(),
          binding: null,
          retryCount: 0,
          retryTimer: null,
        }
        slots.add(slot)
        temporary = true
        initialize(slot)
      }
      if (slot.retryTimer) clearTimer(slot.retryTimer)
      slot.retryTimer = null
      try {
        await slot.client!.revokeBinding(bindingId)
      } finally {
        for (const owner of owners) {
          if (owner !== slot) owner.client?.revoke('binding_revoked')
          remove(owner)
        }
        if (temporary || slots.has(slot)) remove(slot)
        publish()
      }
    },
    async approve(pairingId) {
      if (approving.has(pairingId)) return false
      const owner = pairingOwners.get(pairingId)
      if (!owner?.client) return false
      const approved = await owner.client.approve(pairingId)
      if (approved) approving.add(pairingId)
      return approved
    },
    reject(pairingId) {
      if (approving.has(pairingId)) return false
      const owner = pairingOwners.get(pairingId)
      return owner?.client ? owner.client.reject(pairingId) : false
    },
    listPending() {
      return [...slots].flatMap((slot) =>
        (slot.client?.listPending() ?? []).filter((entry) => !approving.has(entry.pairingId)),
      )
    },
    status: aggregate,
    revoke(reason = 'revoked') {
      suspended = true
      revokeAll(reason)
    },
    suspend(reason) {
      suspended = true
      revokeAll(reason)
    },
    activate() {
      suspended = false
    },
  }
}
