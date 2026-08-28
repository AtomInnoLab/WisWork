import type { OfficeHost } from '../office-document.js'

export const OFFICE_RELAY_ORIGIN = 'https://office.8-216-134-194.sslip.io'
export const PAIRING_RESUME_FEATURE = 'pairing-resume.v1'

const SCHEMA_VERSION = 1
const MAX_OPAQUE_LENGTH = 512
const MAX_CAPABILITIES = 16
const DATA_CAPABILITIES = new Set(['agent.v1', 'web-search.v1', 'web-fetch.v1', 'image-search.v1'])

export const OFFICE_BINDING_DATABASE_SCHEMA = Object.freeze({
  name: 'wiswork-office-pairing',
  version: SCHEMA_VERSION,
  objectStore: 'bindings',
  activeKey: 'active',
})

const hostLabels: Record<Exclude<OfficeHost, 'unknown'>, string> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
}

export interface OfficeBindingDatabase {
  read(): Promise<unknown>
  write(value: unknown): Promise<void>
  delete(): Promise<void>
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
  save(
    enrollment: OfficeBindingEnrollment,
    bindingId: string,
    approvedCapabilities: readonly string[],
  ): Promise<void>
  sign(binding: OfficeStoredBinding, challenge: string): Promise<string>
  forget(): Promise<void>
}

interface BindingStoreDependencies {
  database: OfficeBindingDatabase
  subtle: SubtleCrypto
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

  const forget = () => database.delete()

  return Object.freeze({
    async load(
      host: Exclude<OfficeHost, 'unknown'>,
      capabilities: readonly string[],
    ): Promise<OfficeStoredBinding | undefined> {
      let value: unknown
      try {
        value = await database.read()
      } catch {
        throw new Error('binding_storage_unavailable')
      }
      if (value === undefined) return undefined
      const record = value as Record<string, unknown>
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
        await forget().catch(() => undefined)
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
        await forget().catch(() => undefined)
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
      return Object.freeze({
        host,
        capabilities: Object.freeze([...capabilities]),
        publicKey: encodeBase64Url(raw),
        privateKey: pair.privateKey,
      })
    },
    async save(
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
      await database.write({
        schemaVersion: SCHEMA_VERSION,
        bindingId,
        host: enrollment.host,
        origin: OFFICE_RELAY_ORIGIN,
        capabilities: [...approvedCapabilities],
        privateKey: enrollment.privateKey,
      })
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

  private open(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      const request = this.factory.open(
        OFFICE_BINDING_DATABASE_SCHEMA.name,
        OFFICE_BINDING_DATABASE_SCHEMA.version,
      )
      request.onupgradeneeded = () => {
        const database = request.result
        if (!database.objectStoreNames.contains(OFFICE_BINDING_DATABASE_SCHEMA.objectStore))
          database.createObjectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
      }
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => reject(new Error('binding_storage_unavailable'))
      request.onblocked = () => reject(new Error('binding_storage_unavailable'))
    })
  }

  async read(): Promise<unknown> {
    const database = await this.open()
    try {
      return await new Promise((resolve, reject) => {
        const request = database
          .transaction(OFFICE_BINDING_DATABASE_SCHEMA.objectStore, 'readonly')
          .objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
          .get(OFFICE_BINDING_DATABASE_SCHEMA.activeKey)
        request.onsuccess = () => resolve(request.result)
        request.onerror = () => reject(new Error('binding_storage_unavailable'))
      })
    } finally {
      database.close()
    }
  }

  async write(value: unknown): Promise<void> {
    const database = await this.open()
    try {
      await new Promise<void>((resolve, reject) => {
        const transaction = database.transaction(
          OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
          'readwrite',
        )
        transaction
          .objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
          .put(value, OFFICE_BINDING_DATABASE_SCHEMA.activeKey)
        transaction.oncomplete = () => resolve()
        transaction.onerror = () => reject(new Error('binding_storage_unavailable'))
        transaction.onabort = () => reject(new Error('binding_storage_unavailable'))
      })
    } finally {
      database.close()
    }
  }

  async delete(): Promise<void> {
    const database = await this.open()
    try {
      await new Promise<void>((resolve, reject) => {
        const transaction = database.transaction(
          OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
          'readwrite',
        )
        transaction
          .objectStore(OFFICE_BINDING_DATABASE_SCHEMA.objectStore)
          .delete(OFFICE_BINDING_DATABASE_SCHEMA.activeKey)
        transaction.oncomplete = () => resolve()
        transaction.onerror = () => reject(new Error('binding_storage_unavailable'))
        transaction.onabort = () => reject(new Error('binding_storage_unavailable'))
      })
    } finally {
      database.close()
    }
  }
}

export function createBrowserOfficeBindingStore(): OfficeBindingStore | undefined {
  try {
    const factory = globalThis.indexedDB
    const subtle = globalThis.crypto?.subtle
    if (!factory || !subtle) return undefined
    return createOfficeBindingStore({ database: new IndexedDbBindingDatabase(factory), subtle })
  } catch {
    return undefined
  }
}
