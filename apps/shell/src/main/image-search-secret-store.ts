import { mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from 'node:fs'
import { dirname } from 'node:path'

interface SafeStorageLike {
  isEncryptionAvailable(): boolean
  getSelectedStorageBackend?(): string
  encryptString(value: string): Uint8Array
  decryptString(value: Uint8Array): string
}

export interface ImageSearchSecretStatus {
  configured: boolean
  source: 'environment' | 'stored' | 'none'
}

const SECURE_LINUX_BACKENDS = new Set(['gnome_libsecret', 'kwallet', 'kwallet5', 'kwallet6'])

export function createImageSearchSecretStore(options: {
  path: string
  safeStorage: SafeStorageLike
  platform?: NodeJS.Platform
  env?: NodeJS.ProcessEnv
}) {
  const env = options.env ?? process.env
  const requireEncryption = () => {
    if (!options.safeStorage.isEncryptionAvailable()) throw new Error('secure_storage_unavailable')
    if ((options.platform ?? process.platform) === 'linux') {
      const backend = options.safeStorage.getSelectedStorageBackend?.()
      if (!backend || !SECURE_LINUX_BACKENDS.has(backend))
        throw new Error('secure_storage_unavailable')
    }
  }
  const load = (): string => {
    let encrypted: Uint8Array
    try {
      encrypted = readFileSync(options.path)
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return ''
      throw error
    }
    requireEncryption()
    return options.safeStorage.decryptString(encrypted).trim()
  }
  return {
    load,
    resolve: () => env.SERPAPI_API_KEY?.trim() || load(),
    status(): ImageSearchSecretStatus {
      if (env.SERPAPI_API_KEY?.trim()) return { configured: true, source: 'environment' }
      return load() ? { configured: true, source: 'stored' } : { configured: false, source: 'none' }
    },
    save(value: string): void {
      const key = value.trim()
      if (!key || key.length > 4096) throw new Error('invalid_image_search_key')
      requireEncryption()
      mkdirSync(dirname(options.path), { recursive: true })
      const temporary = `${options.path}.tmp`
      try {
        writeFileSync(temporary, options.safeStorage.encryptString(key), { mode: 0o600 })
        renameSync(temporary, options.path)
      } catch (error) {
        rmSync(temporary, { force: true })
        throw error
      }
    },
    clear(): void {
      rmSync(options.path, { force: true })
    },
  }
}
