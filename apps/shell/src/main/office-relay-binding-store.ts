import type { SafeStorageLike } from '@wiswork/auth'
import { mkdir, readFile, rename, unlink, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

const STORE_VERSION = 1
const MAX_BINDINGS_PER_ACCOUNT = 12
const MAX_TOTAL_RECORDS = 256
const MAX_FILE_BYTES = 1024 * 1024
const IDENTIFIER = /^[A-Za-z0-9_-]{8,128}$/
const HOSTS = new Set(['Word', 'Excel', 'PowerPoint'])
const CAPABILITIES = new Set(['agent.v1', 'web-search.v1', 'web-fetch.v1', 'image-search.v1'])
const SECURE_LINUX_BACKENDS = new Set(['gnome_libsecret', 'kwallet', 'kwallet5', 'kwallet6'])

export interface OfficeRelayBinding {
  bindingId: string
  accountId: string
  host: 'Word' | 'Excel' | 'PowerPoint'
  origin: string
  capabilities: string[]
  createdAt: number
}

export interface OfficeRelayBindingTombstone {
  bindingId: string
  accountId: string
  createdAt: number
}

interface BindingFile {
  version: 1
  bindings: OfficeRelayBinding[]
  tombstones: OfficeRelayBindingTombstone[]
}

export interface OfficeRelayBindingStore {
  listForAccount(accountId: string): Promise<OfficeRelayBinding[]>
  listTombstonesForAccount(accountId: string): Promise<OfficeRelayBindingTombstone[]>
  put(binding: OfficeRelayBinding): Promise<void>
  remove(accountId: string, bindingId: string): Promise<void>
  tombstoneAccount(accountId: string): Promise<void>
  acknowledgeTombstone(accountId: string, bindingId: string): Promise<void>
}

export interface OfficeRelayBindingStoreOptions {
  path: string
  platform?: NodeJS.Platform
  safeStorage: SafeStorageLike
  readFile(path: string): Promise<Uint8Array | null>
  writeFile(path: string, data: Uint8Array): Promise<void>
  rename(from: string, to: string): Promise<void>
  unlink(path: string): Promise<void>
  mkdir(path: string): Promise<void>
  now?: () => number
}

function exact(value: unknown, keys: readonly string[]): value is Record<string, unknown> {
  return Boolean(
    value &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    Object.keys(value).length === keys.length &&
    keys.every((key) => Object.hasOwn(value, key)),
  )
}

function validAccountId(value: unknown): value is string {
  return typeof value === 'string' && value.length >= 1 && value.length <= 512
}

function validOrigin(value: unknown): value is string {
  if (typeof value !== 'string' || value.length > 2_048) return false
  try {
    const url = new URL(value)
    return url.protocol === 'https:' && url.origin === value
  } catch {
    return false
  }
}

function validCapabilities(value: unknown): value is string[] {
  return (
    Array.isArray(value) &&
    value.length >= 1 &&
    value.length <= CAPABILITIES.size &&
    value.every(
      (capability, index, values) =>
        typeof capability === 'string' &&
        CAPABILITIES.has(capability) &&
        values.indexOf(capability) === index,
    )
  )
}

function validCreatedAt(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0
}

function parseBinding(value: unknown): OfficeRelayBinding | null {
  if (!exact(value, ['bindingId', 'accountId', 'host', 'origin', 'capabilities', 'createdAt']))
    return null
  if (
    typeof value.bindingId !== 'string' ||
    !IDENTIFIER.test(value.bindingId) ||
    !validAccountId(value.accountId) ||
    typeof value.host !== 'string' ||
    !HOSTS.has(value.host) ||
    !validOrigin(value.origin) ||
    !validCapabilities(value.capabilities) ||
    !validCreatedAt(value.createdAt)
  )
    return null
  return value as unknown as OfficeRelayBinding
}

function parseTombstone(value: unknown): OfficeRelayBindingTombstone | null {
  if (!exact(value, ['bindingId', 'accountId', 'createdAt'])) return null
  if (
    typeof value.bindingId !== 'string' ||
    !IDENTIFIER.test(value.bindingId) ||
    !validAccountId(value.accountId) ||
    !validCreatedAt(value.createdAt)
  )
    return null
  return value as unknown as OfficeRelayBindingTombstone
}

function parseFile(value: unknown): BindingFile | null {
  if (!exact(value, ['version', 'bindings', 'tombstones']) || value.version !== STORE_VERSION)
    return null
  if (!Array.isArray(value.bindings) || !Array.isArray(value.tombstones)) return null
  const bindings = value.bindings.map(parseBinding)
  const tombstones = value.tombstones.map(parseTombstone)
  if (bindings.some((entry) => !entry) || tombstones.some((entry) => !entry)) return null
  if (bindings.length + tombstones.length > MAX_TOTAL_RECORDS) return null
  const liveKeys = new Set<string>()
  const tombstoneKeys = new Set<string>()
  const liveCounts = new Map<string, number>()
  for (const binding of bindings as OfficeRelayBinding[]) {
    const key = `${binding.accountId}\0${binding.bindingId}`
    if (liveKeys.has(key)) return null
    liveKeys.add(key)
    const count = (liveCounts.get(binding.accountId) ?? 0) + 1
    if (count > MAX_BINDINGS_PER_ACCOUNT) return null
    liveCounts.set(binding.accountId, count)
  }
  for (const tombstone of tombstones as OfficeRelayBindingTombstone[]) {
    const key = `${tombstone.accountId}\0${tombstone.bindingId}`
    if (tombstoneKeys.has(key) || liveKeys.has(key)) return null
    tombstoneKeys.add(key)
  }
  return {
    version: STORE_VERSION,
    bindings: bindings as OfficeRelayBinding[],
    tombstones: tombstones as OfficeRelayBindingTombstone[],
  }
}

function parentPath(path: string): string {
  const at = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return at < 1 ? '.' : path.slice(0, at)
}

function cloneBinding(binding: OfficeRelayBinding): OfficeRelayBinding {
  return { ...binding, capabilities: [...binding.capabilities] }
}

export function createOfficeRelayBindingStore(
  options: OfficeRelayBindingStoreOptions,
): OfficeRelayBindingStore {
  let operation = Promise.resolve()
  const now = options.now ?? Date.now

  const requireEncryption = () => {
    if (!options.safeStorage.isEncryptionAvailable()) throw new Error('secure_storage_unavailable')
    if ((options.platform ?? process.platform) === 'linux') {
      const backend = options.safeStorage.getSelectedStorageBackend?.()
      if (!backend || !SECURE_LINUX_BACKENDS.has(backend))
        throw new Error('secure_storage_unavailable')
    }
  }

  const load = async (): Promise<BindingFile> => {
    requireEncryption()
    const encrypted = await options.readFile(options.path)
    if (!encrypted) return { version: STORE_VERSION, bindings: [], tombstones: [] }
    if (encrypted.byteLength > MAX_FILE_BYTES) {
      await options.unlink(options.path).catch(() => undefined)
      return { version: STORE_VERSION, bindings: [], tombstones: [] }
    }
    try {
      const parsed = parseFile(JSON.parse(options.safeStorage.decryptString(encrypted)))
      if (!parsed) throw new Error('invalid')
      return parsed
    } catch {
      await options.unlink(options.path).catch(() => undefined)
      return { version: STORE_VERSION, bindings: [], tombstones: [] }
    }
  }

  const save = async (file: BindingFile): Promise<void> => {
    requireEncryption()
    await options.mkdir(parentPath(options.path))
    const temporary = `${options.path}.tmp`
    try {
      await options.writeFile(temporary, options.safeStorage.encryptString(JSON.stringify(file)))
      await options.rename(temporary, options.path)
    } catch (error) {
      await options.unlink(temporary).catch(() => undefined)
      throw error
    }
  }

  const serialize = <T>(action: () => Promise<T>): Promise<T> => {
    const next = operation.then(action, action)
    operation = next.then(
      () => undefined,
      () => undefined,
    )
    return next
  }

  return {
    listForAccount(accountId) {
      if (!validAccountId(accountId)) return Promise.reject(new Error('invalid_account_id'))
      return serialize(async () =>
        (await load()).bindings
          .filter((binding) => binding.accountId === accountId)
          .map(cloneBinding),
      )
    },
    listTombstonesForAccount(accountId) {
      if (!validAccountId(accountId)) return Promise.reject(new Error('invalid_account_id'))
      return serialize(async () =>
        (await load()).tombstones
          .filter((tombstone) => tombstone.accountId === accountId)
          .map((tombstone) => ({ ...tombstone })),
      )
    },
    put(binding) {
      const parsed = parseBinding(binding)
      if (!parsed) return Promise.reject(new Error('invalid_office_binding'))
      return serialize(async () => {
        const file = await load()
        const existing = file.bindings.findIndex(
          (entry) => entry.accountId === parsed.accountId && entry.bindingId === parsed.bindingId,
        )
        if (
          existing < 0 &&
          file.bindings.filter((entry) => entry.accountId === parsed.accountId).length >=
            MAX_BINDINGS_PER_ACCOUNT
        )
          throw new Error('office_binding_limit')
        if (existing < 0 && file.bindings.length + file.tombstones.length >= MAX_TOTAL_RECORDS)
          throw new Error('office_binding_store_limit')
        if (existing >= 0) file.bindings[existing] = cloneBinding(parsed)
        else file.bindings.push(cloneBinding(parsed))
        file.tombstones = file.tombstones.filter(
          (entry) => entry.accountId !== parsed.accountId || entry.bindingId !== parsed.bindingId,
        )
        await save(file)
      })
    },
    remove(accountId, bindingId) {
      if (!validAccountId(accountId) || !IDENTIFIER.test(bindingId))
        return Promise.reject(new Error('invalid_office_binding'))
      return serialize(async () => {
        const file = await load()
        file.bindings = file.bindings.filter(
          (entry) => entry.accountId !== accountId || entry.bindingId !== bindingId,
        )
        await save(file)
      })
    },
    tombstoneAccount(accountId) {
      if (!validAccountId(accountId)) return Promise.reject(new Error('invalid_account_id'))
      return serialize(async () => {
        const file = await load()
        const active = file.bindings.filter((binding) => binding.accountId === accountId)
        file.bindings = file.bindings.filter((binding) => binding.accountId !== accountId)
        const tombstones = new Map(
          file.tombstones
            .filter((entry) => entry.accountId === accountId)
            .map((entry) => [entry.bindingId, entry]),
        )
        for (const binding of active) {
          tombstones.set(binding.bindingId, {
            bindingId: binding.bindingId,
            accountId,
            createdAt: now(),
          })
        }
        file.tombstones = [
          ...file.tombstones.filter((entry) => entry.accountId !== accountId),
          ...tombstones.values(),
        ]
        await save(file)
      })
    },
    acknowledgeTombstone(accountId, bindingId) {
      if (!validAccountId(accountId) || !IDENTIFIER.test(bindingId))
        return Promise.reject(new Error('invalid_office_binding'))
      return serialize(async () => {
        const file = await load()
        file.tombstones = file.tombstones.filter(
          (entry) => entry.accountId !== accountId || entry.bindingId !== bindingId,
        )
        await save(file)
      })
    },
  }
}

export function createElectronOfficeRelayBindingStore(options: {
  userDataPath: string
  safeStorage: SafeStorageLike
  platform?: NodeJS.Platform
}): OfficeRelayBindingStore {
  return createOfficeRelayBindingStore({
    path: join(options.userDataPath, 'office-pairings.enc'),
    safeStorage: options.safeStorage,
    ...(options.platform ? { platform: options.platform } : {}),
    async readFile(path) {
      try {
        return await readFile(path)
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null
        throw error
      }
    },
    writeFile: (path, data) => writeFile(path, data, { mode: 0o600 }),
    rename,
    unlink,
    mkdir: (path) => mkdir(path, { recursive: true, mode: 0o700 }).then(() => undefined),
  })
}
