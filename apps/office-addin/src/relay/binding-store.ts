import type { OfficeHost } from '../office-document.js'

export const OFFICE_RELAY_ORIGIN = 'https://office.8-216-134-194.sslip.io'
export const PAIRING_RESUME_FEATURE = 'pairing-resume.v1'

const SCHEMA_VERSION = 1
const DATABASE_VERSION = 2
const ENROLLMENT_LEASE_MS = 180_000
const MAX_OPAQUE_LENGTH = 512
const MAX_CAPABILITIES = 16
const DATA_CAPABILITIES = new Set(['agent.v1', 'web-search.v1', 'web-fetch.v1', 'image-search.v1'])

export const OFFICE_BINDING_DATABASE_SCHEMA = Object.freeze({
  name: 'wiswork-office-pairing',
  version: DATABASE_VERSION,
  objectStore: 'bindings',
  activeKeyPrefix: 'active:',
  enrollmentKeyPrefix: 'enrollment:',
})

type SupportedOfficeHost = Exclude<OfficeHost, 'unknown'>

const hostLabels: Record<Exclude<OfficeHost, 'unknown'>, string> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
}

export interface OfficeBindingDatabase {
  read(host: SupportedOfficeHost): Promise<{
    record: unknown
    blockedBindingId: unknown
  }>
  claim(host: SupportedOfficeHost, owner: string, expiresAt: number): Promise<boolean>
  stage(host: SupportedOfficeHost, owner: string, value: unknown): Promise<boolean>
  activate(host: SupportedOfficeHost, owner: string, bindingId: string): Promise<boolean>
  release(host: SupportedOfficeHost, owner: string, bindingId?: string): Promise<void>
  block(host: SupportedOfficeHost, bindingId: string): Promise<boolean>
  deleteBlocked(host: SupportedOfficeHost, bindingId: string): Promise<boolean>
  deleteIf(host: SupportedOfficeHost, predicate: (value: unknown) => boolean): Promise<boolean>
}

export interface OfficeBindingEnrollment {
  readonly host: Exclude<OfficeHost, 'unknown'>
  readonly capabilities: readonly string[]
  readonly publicKey: string
  readonly privateKey: CryptoKey
}

export interface OfficeStoredBinding {
  readonly schemaVersion: 1
  readonly bindingId: string
  readonly host: Exclude<OfficeHost, 'unknown'>
  readonly origin: typeof OFFICE_RELAY_ORIGIN
  readonly capabilities: readonly string[]
  readonly privateKey: CryptoKey
}

export interface OfficeBindingStore {
  load(
    host: Exclude<OfficeHost, 'unknown'>,
    capabilities: readonly string[],
  ): Promise<OfficeStoredBinding | undefined>
  createEnrollment(
    host: Exclude<OfficeHost, 'unknown'>,
    capabilities: readonly string[],
  ): Promise<OfficeBindingEnrollment>
  stage(
    enrollment: OfficeBindingEnrollment,
    bindingId: string,
    approvedCapabilities: readonly string[],
  ): Promise<void>
  activate(
    enrollment: OfficeBindingEnrollment,
    bindingId: string,
    approvedCapabilities: readonly string[],
  ): Promise<OfficeStoredBinding>
  abort(enrollment: OfficeBindingEnrollment, bindingId?: string): Promise<void>
  sign(binding: OfficeStoredBinding, challenge: string): Promise<string>
  forget(host: SupportedOfficeHost, bindingId?: string): Promise<void>
}

interface BindingStoreDependencies {
  database: OfficeBindingDatabase
  subtle: SubtleCrypto
  now?: () => number
}

const exactKeys = (value: Record<string, unknown>, keys: readonly string[]) =>
  Object.keys(value).length === keys.length && keys.every((key) => key in value)

const opaque = (value: unknown): value is string =>
  typeof value === 'string' &&
  value.length > 0 &&
  value.length <= MAX_OPAQUE_LENGTH &&
  /^[A-Za-z0-9_-]+$/.test(value)

const exactCapabilities = (value: unknown, expected: readonly string[]): value is string[] =>
  Array.isArray(value) &&
  value.length > 0 &&
  value.length <= MAX_CAPABILITIES &&
  value.length === expected.length &&
  value.every(
    (capability, index, values) =>
      typeof capability === 'string' &&
      DATA_CAPABILITIES.has(capability) &&
      values.indexOf(capability) === index &&
      capability === expected[index],
  )

function validPrivateKey(value: unknown): value is CryptoKey {
  if (!value || typeof value !== 'object') return false
  const key = value as Partial<CryptoKey>
  const algorithm = key.algorithm as Partial<EcKeyAlgorithm> | undefined
  return (
    key.type === 'private' &&
    key.extractable === false &&
    algorithm?.name === 'ECDSA' &&
    algorithm.namedCurve === 'P-256' &&
    Array.isArray(key.usages) &&
    key.usages.length === 1 &&
    key.usages[0] === 'sign'
  )
}

function encodeBase64Url(value: ArrayBuffer): string {
  const bytes = new Uint8Array(value)
  let binary = ''
  for (const byte of bytes) binary += String.fromCharCode(byte)
  return btoa(binary).replaceAll('+', '-').replaceAll('/', '_').replace(/=+$/, '')
}

export function createOfficeBindingStore(
  dependencies: BindingStoreDependencies,
): OfficeBindingStore {
  const { database, subtle } = dependencies
  const now = dependencies.now ?? Date.now

  const recordIdentity = (value: unknown): string => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      try {
        return `${typeof value}:${JSON.stringify(value)}`
      } catch {
        return `${typeof value}:${String(value)}`
      }
    }
    const record = value as Record<string, unknown>
    const identity = Object.fromEntries(
      Object.keys(record)
        .sort()
        .filter((key) => key !== 'privateKey')
        .map((key) => [key, record[key]]),
    )
    try {
      return JSON.stringify(identity)
    } catch {
      return Object.entries(identity)
        .map(([key, entry]) => `${key}:${typeof entry}:${String(entry)}`)
        .join('|')
    }
  }

  const cleanup = (host: SupportedOfficeHost, value: unknown) => {
    const identity = recordIdentity(value)
    return database.deleteIf(host, (current) => recordIdentity(current) === identity)
  }

  const forget = async (host: SupportedOfficeHost, bindingId?: string) => {
    let target = bindingId
    if (!target) {
      const snapshot = await database.read(host)
      if (opaque(snapshot.blockedBindingId)) target = snapshot.blockedBindingId
      else return
    }
    const blocked = await database.block(host, target)
    if (!blocked) throw new Error('binding_forget_conflict')
    const deleted = await database.deleteBlocked(host, target)
    if (!deleted) throw new Error('binding_storage_unavailable')
  }

  const liveBinding = (
    enrollment: OfficeBindingEnrollment,
    bindingId: string,
    approvedCapabilities: readonly string[],
  ): OfficeStoredBinding =>
    Object.freeze({
      schemaVersion: 1 as const,
      bindingId,
      host: enrollment.host,
      origin: OFFICE_RELAY_ORIGIN,
      capabilities: Object.freeze([...approvedCapabilities]),
      privateKey: enrollment.privateKey,
    })

  const validPendingRecord = (record: Record<string, unknown>, host: SupportedOfficeHost) =>
    exactKeys(record, [
      'schemaVersion',
      'bindingId',
      'host',
      'origin',
      'capabilities',
      'privateKey',
      'pending',
      'enrollmentOwner',
      'pendingExpiresAt',
    ]) &&
    record.schemaVersion === SCHEMA_VERSION &&
    opaque(record.bindingId) &&
    record.host === host &&
    record.origin === OFFICE_RELAY_ORIGIN &&
    Array.isArray(record.capabilities) &&
    exactCapabilities(record.capabilities, record.capabilities) &&
    validPrivateKey(record.privateKey) &&
    record.pending === true &&
    opaque(record.enrollmentOwner) &&
    typeof record.pendingExpiresAt === 'number' &&
    Number.isSafeInteger(record.pendingExpiresAt) &&
    record.pendingExpiresAt <= now() + ENROLLMENT_LEASE_MS

  return Object.freeze({
    async load(
      host: Exclude<OfficeHost, 'unknown'>,
      capabilities: readonly string[],
    ): Promise<OfficeStoredBinding | undefined> {
      let value: unknown
      try {
        const snapshot = await database.read(host)
        value = snapshot.record
        if (
          opaque(snapshot.blockedBindingId) &&
          value &&
          typeof value === 'object' &&
          !Array.isArray(value) &&
          (value as Record<string, unknown>).bindingId === snapshot.blockedBindingId
        )
          return undefined
      } catch {
        throw new Error('binding_storage_unavailable')
      }
      if (value === undefined) return undefined
      const record = value as Record<string, unknown>
      if (
        value &&
        typeof value === 'object' &&
        !Array.isArray(value) &&
        validPendingRecord(record, host)
      ) {
        if ((record.pendingExpiresAt as number) > now()) return undefined
        await cleanup(host, value).catch(() => undefined)
        return undefined
      }
      if (
        !value ||
        typeof value !== 'object' ||
        Array.isArray(value) ||
        !exactKeys(record, [
          'schemaVersion',
          'bindingId',
          'host',
          'origin',
          'capabilities',
          'privateKey',
        ]) ||
        record.schemaVersion !== SCHEMA_VERSION ||
        !opaque(record.bindingId) ||
        record.host !== host ||
        record.origin !== OFFICE_RELAY_ORIGIN ||
        !exactCapabilities(record.capabilities, capabilities) ||
        !validPrivateKey(record.privateKey)
      ) {
        await cleanup(host, value).catch(() => undefined)
        return undefined
      }
      try {
        const probe = await subtle.sign(
          { name: 'ECDSA', hash: 'SHA-256' },
          record.privateKey,
          new TextEncoder().encode('wiswork-office-binding-key-check-v1'),
        )
        if (probe.byteLength !== 64) throw new Error('binding_key_unusable')
      } catch {
        await cleanup(host, value).catch(() => undefined)
        return undefined
      }
      return Object.freeze({
        schemaVersion: 1 as const,
        bindingId: record.bindingId,
        host,
        origin: OFFICE_RELAY_ORIGIN,
        capabilities: Object.freeze([...(record.capabilities as string[])]),
        privateKey: record.privateKey,
      })
    },
    async createEnrollment(
      host: Exclude<OfficeHost, 'unknown'>,
      capabilities: readonly string[],
    ): Promise<OfficeBindingEnrollment> {
      if (!exactCapabilities([...capabilities], capabilities)) throw new Error('binding_invalid')
      const pair = (await subtle.generateKey({ name: 'ECDSA', namedCurve: 'P-256' }, false, [
        'sign',
        'verify',
      ])) as CryptoKeyPair
      if (!validPrivateKey(pair.privateKey)) throw new Error('binding_key_unusable')
      const raw = await subtle.exportKey('raw', pair.publicKey)
      const bytes = new Uint8Array(raw)
      if (bytes.byteLength !== 65 || bytes[0] !== 0x04) throw new Error('binding_key_unusable')
      const prepared = Object.freeze({
        host,
        capabilities: Object.freeze([...capabilities]),
        publicKey: encodeBase64Url(raw),
        privateKey: pair.privateKey,
      })
      let claimed: boolean
      try {
        claimed = await database.claim(host, prepared.publicKey, now() + ENROLLMENT_LEASE_MS)
      } catch {
        throw new Error('binding_storage_unavailable')
      }
      if (!claimed) throw new Error('binding_enrollment_in_progress')
      return prepared
    },
    async stage(
      enrollment: OfficeBindingEnrollment,
      bindingId: string,
      approvedCapabilities: readonly string[],
    ): Promise<void> {
      if (
        !opaque(bindingId) ||
        !validPrivateKey(enrollment.privateKey) ||
        !exactCapabilities([...approvedCapabilities], approvedCapabilities) ||
        approvedCapabilities.some((capability) => !enrollment.capabilities.includes(capability))
      )
        throw new Error('binding_invalid')
      let staged: boolean
      try {
        staged = await database.stage(enrollment.host, enrollment.publicKey, {
          schemaVersion: SCHEMA_VERSION,
          bindingId,
          host: enrollment.host,
          origin: OFFICE_RELAY_ORIGIN,
          capabilities: [...approvedCapabilities],
          privateKey: enrollment.privateKey,
          pending: true,
          enrollmentOwner: enrollment.publicKey,
          pendingExpiresAt: now() + ENROLLMENT_LEASE_MS,
        })
      } catch {
        throw new Error('binding_storage_unavailable')
      }
      if (!staged) {
        await database.release(enrollment.host, enrollment.publicKey).catch(() => undefined)
        throw new Error('binding_enrollment_conflict')
      }
    },
    async activate(
      enrollment: OfficeBindingEnrollment,
      bindingId: string,
      approvedCapabilities: readonly string[],
    ): Promise<OfficeStoredBinding> {
      if (
        !opaque(bindingId) ||
        !validPrivateKey(enrollment.privateKey) ||
        !exactCapabilities([...approvedCapabilities], approvedCapabilities) ||
        approvedCapabilities.some((capability) => !enrollment.capabilities.includes(capability))
      )
        throw new Error('binding_invalid')
      let activated: boolean
      try {
        activated = await database.activate(enrollment.host, enrollment.publicKey, bindingId)
      } catch {
        throw new Error('binding_storage_unavailable')
      }
      if (!activated) throw new Error('binding_enrollment_conflict')
      return liveBinding(enrollment, bindingId, approvedCapabilities)
    },
    async abort(enrollment: OfficeBindingEnrollment, bindingId?: string): Promise<void> {
      await database.release(enrollment.host, enrollment.publicKey, bindingId)
    },
    async sign(binding: OfficeStoredBinding, challenge: string): Promise<string> {
      if (!opaque(challenge) || !validPrivateKey(binding.privateKey))
        throw new Error('binding_key_unusable')
      const transcript = [
        'wiswork-office-resume-v1',
        binding.bindingId,
        challenge,
        OFFICE_RELAY_ORIGIN,
        hostLabels[binding.host],
      ].join('\n')
      const signature = await subtle.sign(
        { name: 'ECDSA', hash: 'SHA-256' },
        binding.privateKey,
        new TextEncoder().encode(transcript),
      )
      if (signature.byteLength !== 64) throw new Error('binding_key_unusable')
      return encodeBase64Url(signature)
    },
    forget,
  })
}

class IndexedDbBindingDatabase implements OfficeBindingDatabase {
  constructor(private readonly factory: IDBFactory) {}

  private activeKey(host: SupportedOfficeHost) {
    return `${OFFICE_BINDING_DATABASE_SCHEMA.activeKeyPrefix}${host}`
  }

  private enrollmentKey(host: SupportedOfficeHost) {
    return `${OFFICE_BINDING_DATABASE_SCHEMA.enrollmentKeyPrefix}${host}`
  }

  private blockedKey(host: SupportedOfficeHost) {
    return `blocked:${host}`
  }

  private request<T = unknown>(request: IDBRequest<T>): Promise<T> {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => reject(new Error('binding_storage_unavailable'))
    })
  }

  private completed(transaction: IDBTransaction): Promise<void> {
    return new Promise((resolve, reject) => {
      transaction.oncomplete = () => resolve()
      transaction.onerror = () => reject(new Error('binding_storage_unavailable'))
      transaction.onabort = () => reject(new Error('binding_storage_unavailable'))
    })
  }

  private async failTransaction(
    transaction: IDBTransaction,
    completed: Promise<void>,
  ): Promise<never> {
    try {
      transaction.abort()
    } catch {
      // The request failure may already have aborted the transaction.
    }
    await completed.catch(() => undefined)
    throw new Error('binding_storage_unavailable')
  }

  private open(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      let blocked = false
      const request = this.factory.open(
        OFFICE_BINDING_DATABASE_SCHEMA.name,
        OFFICE_BINDING_DATABASE_SCHEMA.version,
      )
      request.onupgradeneeded = (event) => {
        const database = request.result
        if (!database.objectStoreNames.contains(OFFICE_BINDING_DATABASE_SCHEMA.objectStore))
          database.createObjectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
        if ((event.oldVersion ?? 0) < 2 && request.transaction) {
          const store = request.transaction.objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
          const legacy = store.get('active')
          legacy.onsuccess = () => {
            const record = legacy.result as Record<string, unknown> | undefined
            if (
              record &&
              (record.host === 'word' || record.host === 'excel' || record.host === 'powerpoint')
            )
              store.put(record, this.activeKey(record.host))
            store.delete('active')
          }
        }
      }
      request.onsuccess = () => {
        if (blocked) request.result.close()
        else resolve(request.result)
      }
      request.onerror = () => reject(new Error('binding_storage_unavailable'))
      request.onblocked = () => {
        blocked = true
        reject(new Error('binding_storage_unavailable'))
      }
    })
  }

  async read(host: SupportedOfficeHost): Promise<{
    record: unknown
    blockedBindingId: unknown
  }> {
    const database = await this.open()
    try {
      const transaction = database.transaction(
        OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
        'readonly',
      )
      const store = transaction.objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
      const record = await this.request(store.get(this.activeKey(host)))
      const blocked = (await this.request(store.get(this.blockedKey(host)))) as
        { bindingId?: unknown } | undefined
      return { record, blockedBindingId: blocked?.bindingId }
    } finally {
      database.close()
    }
  }

  async claim(host: SupportedOfficeHost, owner: string, expiresAt: number): Promise<boolean> {
    const database = await this.open()
    try {
      const transaction = database.transaction(
        OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
        'readwrite',
      )
      const completed = this.completed(transaction)
      const store = transaction.objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
      try {
        const active = await this.request(store.get(this.activeKey(host)))
        const existing = (await this.request(store.get(this.enrollmentKey(host)))) as
          { owner?: unknown; expiresAt?: unknown } | undefined
        const validExisting =
          existing &&
          typeof existing.owner === 'string' &&
          typeof existing.expiresAt === 'number' &&
          Number.isSafeInteger(existing.expiresAt) &&
          existing.expiresAt > Date.now()
        if (active !== undefined || (validExisting && existing.owner !== owner)) {
          transaction.abort()
          await completed.catch(() => undefined)
          return false
        }
        store.put({ schemaVersion: 1, owner, expiresAt }, this.enrollmentKey(host))
        await completed
        return true
      } catch {
        return this.failTransaction(transaction, completed)
      }
    } finally {
      database.close()
    }
  }

  async stage(host: SupportedOfficeHost, owner: string, value: unknown): Promise<boolean> {
    const database = await this.open()
    try {
      const transaction = database.transaction(
        OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
        'readwrite',
      )
      const completed = this.completed(transaction)
      const store = transaction.objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
      try {
        const active = await this.request(store.get(this.activeKey(host)))
        const claim = (await this.request(store.get(this.enrollmentKey(host)))) as
          { owner?: unknown } | undefined
        if (active !== undefined || claim?.owner !== owner) {
          transaction.abort()
          await completed.catch(() => undefined)
          return false
        }
        store.put(value, this.activeKey(host))
        store.delete(this.enrollmentKey(host))
        await completed
        return true
      } catch {
        return this.failTransaction(transaction, completed)
      }
    } finally {
      database.close()
    }
  }

  async activate(host: SupportedOfficeHost, owner: string, bindingId: string): Promise<boolean> {
    const database = await this.open()
    try {
      const transaction = database.transaction(
        OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
        'readwrite',
      )
      const completed = this.completed(transaction)
      const store = transaction.objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
      try {
        const record = (await this.request(store.get(this.activeKey(host)))) as
          Record<string, unknown> | undefined
        if (
          record?.pending !== true ||
          record.enrollmentOwner !== owner ||
          record.bindingId !== bindingId
        ) {
          transaction.abort()
          await completed.catch(() => undefined)
          return false
        }
        const {
          pending: _pending,
          enrollmentOwner: _owner,
          pendingExpiresAt: _expiresAt,
          ...live
        } = record
        store.put(live, this.activeKey(host))
        await completed
        return true
      } catch {
        return this.failTransaction(transaction, completed)
      }
    } finally {
      database.close()
    }
  }

  async release(host: SupportedOfficeHost, owner: string, bindingId?: string): Promise<void> {
    const database = await this.open()
    try {
      const transaction = database.transaction(
        OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
        'readwrite',
      )
      const completed = this.completed(transaction)
      const store = transaction.objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
      try {
        const claim = (await this.request(store.get(this.enrollmentKey(host)))) as
          { owner?: unknown } | undefined
        if (claim?.owner === owner) store.delete(this.enrollmentKey(host))
        if (bindingId) {
          const record = (await this.request(store.get(this.activeKey(host)))) as
            Record<string, unknown> | undefined
          if (
            record?.pending === true &&
            record.enrollmentOwner === owner &&
            record.bindingId === bindingId
          )
            store.delete(this.activeKey(host))
        }
        await completed
      } catch {
        return this.failTransaction(transaction, completed)
      }
    } finally {
      database.close()
    }
  }

  async block(host: SupportedOfficeHost, bindingId: string): Promise<boolean> {
    const database = await this.open()
    try {
      const transaction = database.transaction(
        OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
        'readwrite',
      )
      const completed = this.completed(transaction)
      const store = transaction.objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
      try {
        const record = (await this.request(store.get(this.activeKey(host)))) as
          Record<string, unknown> | undefined
        const blocked = (await this.request(store.get(this.blockedKey(host)))) as
          { bindingId?: unknown } | undefined
        if (blocked?.bindingId === bindingId) {
          await completed
          return true
        }
        if (record?.bindingId !== bindingId) {
          transaction.abort()
          await completed.catch(() => undefined)
          return false
        }
        store.put({ schemaVersion: 1, bindingId }, this.blockedKey(host))
        await completed
        return true
      } catch {
        return this.failTransaction(transaction, completed)
      }
    } finally {
      database.close()
    }
  }

  async deleteBlocked(host: SupportedOfficeHost, bindingId: string): Promise<boolean> {
    const database = await this.open()
    try {
      const transaction = database.transaction(
        OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
        'readwrite',
      )
      const completed = this.completed(transaction)
      const store = transaction.objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
      try {
        const record = (await this.request(store.get(this.activeKey(host)))) as
          Record<string, unknown> | undefined
        const blocked = (await this.request(store.get(this.blockedKey(host)))) as
          { bindingId?: unknown } | undefined
        if (record?.bindingId !== bindingId || blocked?.bindingId !== bindingId) {
          transaction.abort()
          await completed.catch(() => undefined)
          return false
        }
        store.delete(this.activeKey(host))
        store.delete(this.enrollmentKey(host))
        store.delete(this.blockedKey(host))
        await completed
        return true
      } catch {
        return this.failTransaction(transaction, completed)
      }
    } finally {
      database.close()
    }
  }

  async deleteIf(
    host: SupportedOfficeHost,
    predicate: (value: unknown) => boolean,
  ): Promise<boolean> {
    const database = await this.open()
    try {
      const transaction = database.transaction(
        OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
        'readwrite',
      )
      const completed = this.completed(transaction)
      const store = transaction.objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
      try {
        const record = await this.request(store.get(this.activeKey(host)))
        if (!predicate(record)) {
          transaction.abort()
          await completed.catch(() => undefined)
          return false
        }
        store.delete(this.activeKey(host))
        await completed
        return true
      } catch {
        return this.failTransaction(transaction, completed)
      }
    } finally {
      database.close()
    }
  }
}

export function createIndexedDbOfficeBindingDatabase(factory: IDBFactory): OfficeBindingDatabase {
  return new IndexedDbBindingDatabase(factory)
}

export function createBrowserOfficeBindingStore(): OfficeBindingStore | undefined {
  try {
    const factory = globalThis.indexedDB
    const subtle = globalThis.crypto?.subtle
    if (!factory || !subtle) return undefined
    return createOfficeBindingStore({
      database: createIndexedDbOfficeBindingDatabase(factory),
      subtle,
    })
  } catch {
    return undefined
  }
}
