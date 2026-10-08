import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { createImageSearchSecretStore } from '../src/main/image-search-secret-store'

const directories: string[] = []

function makeStore(overrides: Record<string, unknown> = {}) {
  const directory = mkdtempSync(join(tmpdir(), 'image-search-secret-'))
  directories.push(directory)
  const path = join(directory, 'image-search-key.enc')
  return {
    path,
    store: createImageSearchSecretStore({
      path,
      platform: 'darwin',
      safeStorage: {
        isEncryptionAvailable: () => true,
        encryptString: (value) => Buffer.from(value.split('').reverse().join('')),
        decryptString: (value) => Buffer.from(value).toString().split('').reverse().join(''),
      },
      ...overrides,
    }),
  }
}

afterEach(() => {
  for (const directory of directories.splice(0)) rmSync(directory, { recursive: true, force: true })
})

describe('image search secret store', () => {
  it('encrypts the key and exposes only non-sensitive status', () => {
    const configured = makeStore()
    configured.store.save('secret-serpapi-key')

    expect(readFileSync(configured.path, 'utf8')).not.toContain('secret-serpapi-key')
    expect(configured.store.status()).toEqual({ configured: true, source: 'stored' })
    expect(configured.store.status()).not.toHaveProperty('key')
    expect(configured.store.load()).toBe('secret-serpapi-key')
  })

  it('clears the persisted key', () => {
    const configured = makeStore()
    configured.store.save('secret-serpapi-key')
    configured.store.clear()
    expect(configured.store.load()).toBe('')
    expect(configured.store.status()).toEqual({ configured: false, source: 'none' })
  })

  it('keeps the environment key as an unreadable override', () => {
    const configured = makeStore({ env: { SERPAPI_API_KEY: 'environment-secret' } })
    expect(configured.store.resolve()).toBe('environment-secret')
    expect(configured.store.status()).toEqual({ configured: true, source: 'environment' })
  })

  it('fails closed when secure storage is unavailable', () => {
    const configured = makeStore({
      safeStorage: {
        isEncryptionAvailable: () => false,
        encryptString: () => Buffer.alloc(0),
        decryptString: () => '',
      },
    })
    expect(() => configured.store.save('secret')).toThrow('secure_storage_unavailable')
    expect(configured.store.load()).toBe('')
    expect(configured.store.status()).toEqual({ configured: false, source: 'none' })
  })

  it('fails closed when an existing key cannot be decrypted securely', () => {
    const configured = makeStore()
    configured.store.save('secret')
    const unavailable = makeStore({
      path: configured.path,
      safeStorage: {
        isEncryptionAvailable: () => false,
        encryptString: () => Buffer.alloc(0),
        decryptString: () => '',
      },
    })
    expect(() => unavailable.store.load()).toThrow('secure_storage_unavailable')
  })
})
