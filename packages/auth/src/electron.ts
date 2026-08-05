import { readFile, mkdir, rename, unlink, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { authConfigFromEnv } from './config'
import { createAuthClient, type AuthClient } from './oauth'
import { AuthError, type AuthSession, type SessionStore } from './session'

export interface SafeStorageLike {
  isEncryptionAvailable(): boolean
  getSelectedStorageBackend?(): string
  encryptString(value: string): Uint8Array
  decryptString(value: Uint8Array): string
}

export interface EncryptedStoreOptions {
  path: string
  platform?: NodeJS.Platform
  safeStorage: SafeStorageLike
  readFile(path: string): Promise<Uint8Array | null>
  writeFile(path: string, data: Uint8Array): Promise<void>
  rename(from: string, to: string): Promise<void>
  unlink(path: string): Promise<void>
  mkdir(path: string): Promise<void>
}

function parentPath(path: string): string {
  const at = Math.max(path.lastIndexOf('/'), path.lastIndexOf('\\'))
  return at < 1 ? '.' : path.slice(0, at)
}

export function createEncryptedSessionStore(options: EncryptedStoreOptions): SessionStore {
  const requireEncryption = () => {
    if (!options.safeStorage.isEncryptionAvailable())
      throw new AuthError('secure_storage_unavailable')
    if ((options.platform ?? process.platform) === 'linux') {
      const backend = options.safeStorage.getSelectedStorageBackend?.()
      const secureBackends = new Set(['gnome_libsecret', 'kwallet', 'kwallet5', 'kwallet6'])
      if (!backend || !secureBackends.has(backend))
        throw new AuthError('secure_storage_unavailable')
    }
  }
  return {
    async load() {
      const encrypted = await options.readFile(options.path)
      if (!encrypted) return null
      requireEncryption()
      try {
        const parsed = JSON.parse(options.safeStorage.decryptString(encrypted)) as AuthSession
        if (
          !parsed ||
          typeof parsed.accessToken !== 'string' ||
          typeof parsed.refreshToken !== 'string' ||
          typeof parsed.userId !== 'string'
        )
          throw new Error('invalid')
        return parsed
      } catch {
        await options.unlink(options.path).catch(() => undefined)
        return null
      }
    },
    async save(session) {
      requireEncryption()
      await options.mkdir(parentPath(options.path))
      const temporary = `${options.path}.tmp`
      await options.writeFile(temporary, options.safeStorage.encryptString(JSON.stringify(session)))
      await options.rename(temporary, options.path)
    },
    async clear() {
      await options.unlink(options.path).catch(() => undefined)
    },
  }
}

export function extractCallbackUrl(input: string | readonly string[]): string | null {
  const values = typeof input === 'string' ? [input] : input
  for (const value of values) {
    if (typeof value !== 'string' || value.length > 4_096) continue
    try {
      const url = new URL(value)
      if (url.protocol === 'wiswork:' && url.hostname === 'oauth' && url.pathname === '/callback')
        return value
    } catch {
      /* argv contains many non-URL values */
    }
  }
  return null
}

export interface ElectronAuthRuntimeOptions {
  userDataPath: string
  safeStorage: SafeStorageLike
  openExternal(url: string): Promise<unknown>
  fetch?: typeof globalThis.fetch
  env?: NodeJS.ProcessEnv
}

export interface ElectronAuthRuntime {
  client: AuthClient
  beginLogin(): Promise<boolean>
  consumeDeepLink(input: string | readonly string[]): Promise<boolean>
}

/** Shared bootstrap used by the unified shell and every standalone editor. */
export function createElectronAuthRuntime(
  options: ElectronAuthRuntimeOptions,
): ElectronAuthRuntime {
  const sessionPath = join(options.userDataPath, 'auth-session.enc')
  const store = createEncryptedSessionStore({
    path: sessionPath,
    safeStorage: options.safeStorage,
    async readFile(path) {
      try {
        return await readFile(path)
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null
        throw error
      }
    },
    writeFile,
    rename,
    unlink,
    mkdir: (path) => mkdir(path, { recursive: true }).then(() => undefined),
  })
  const client = createAuthClient({
    store,
    ...(options.fetch ? { fetch: options.fetch } : {}),
    config: authConfigFromEnv(options.env),
  })
  return {
    client,
    async beginLogin() {
      const request = client.createAuthorizationRequest()
      await options.openExternal(request.url)
      return true
    },
    async consumeDeepLink(input) {
      const callback = extractCallbackUrl(input)
      if (!callback) return false
      await client.consumeCallback(callback)
      return true
    },
  }
}

export interface AuthProtocolRoutingOptions {
  registerProtocolClient(protocol: string): boolean
  onOpenUrl(handler: (event: { preventDefault(): void }, url: string) => void): void
  onSecondInstance(handler: (argv: readonly string[]) => void): void
  initialArgv: readonly string[]
  consume(input: string | readonly string[]): Promise<unknown>
}

export function registerAuthProtocolRouting(options: AuthProtocolRoutingOptions): {
  protocolRegistered: boolean
  route(input: string | readonly string[]): boolean
} {
  const route = (input: string | readonly string[]): boolean => {
    const callback = extractCallbackUrl(input)
    if (!callback) return false
    void options.consume(callback).catch(() => undefined)
    return true
  }
  const protocolRegistered = options.registerProtocolClient('wiswork')
  options.onOpenUrl((event, url) => {
    if (!extractCallbackUrl(url)) return
    event.preventDefault()
    route(url)
  })
  options.onSecondInstance((argv) => {
    route(argv)
  })
  route(options.initialArgv)
  return { protocolRegistered, route }
}

let defaultElectronAuthRuntime: ElectronAuthRuntime | null = null

export function initializeElectronAuthRuntime(
  options: ElectronAuthRuntimeOptions,
): ElectronAuthRuntime {
  defaultElectronAuthRuntime = createElectronAuthRuntime(options)
  return defaultElectronAuthRuntime
}

export function getElectronAuthRuntimeOrNull(): ElectronAuthRuntime | null {
  return defaultElectronAuthRuntime
}

export function getElectronAuthRuntime(): ElectronAuthRuntime {
  if (!defaultElectronAuthRuntime) throw new AuthError('auth_not_initialized')
  return defaultElectronAuthRuntime
}
