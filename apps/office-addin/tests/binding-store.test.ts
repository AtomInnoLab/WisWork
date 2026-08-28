import { describe, expect, it } from 'vitest'
import {
  OFFICE_BINDING_DATABASE_SCHEMA,
  OFFICE_RELAY_ORIGIN,
  createOfficeBindingStore,
  type OfficeBindingDatabase,
} from '../src/relay/binding-store.js'

class MemoryBindingDatabase implements OfficeBindingDatabase {
  value: unknown
  deletes = 0

  async read(): Promise<unknown> {
    return this.value
  }

  async write(value: unknown): Promise<void> {
    this.value = value
  }

  async delete(): Promise<void> {
    this.deletes += 1
    this.value = undefined
  }
}

describe('Office persistent binding store', () => {
  it('generates and stores only a non-exportable P-256 private CryptoKey', async () => {
    const database = new MemoryBindingDatabase()
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle })
    const enrollment = await store.createEnrollment('word', ['agent.v1', 'web-search.v1'])

    expect(enrollment.publicKey).toMatch(/^[A-Za-z0-9_-]+$/)
    const publicBytes = Buffer.from(enrollment.publicKey, 'base64url')
    expect(publicBytes).toHaveLength(65)
    expect(publicBytes[0]).toBe(0x04)
    expect(enrollment.privateKey.type).toBe('private')
    expect(enrollment.privateKey.extractable).toBe(false)
    expect(enrollment.privateKey.algorithm).toMatchObject({ name: 'ECDSA', namedCurve: 'P-256' })
    await expect(crypto.subtle.exportKey('pkcs8', enrollment.privateKey)).rejects.toThrow()

    await store.save(enrollment, 'binding_12345678', ['agent.v1'])
    expect(database.value).toEqual({
      schemaVersion: 1,
      bindingId: 'binding_12345678',
      host: 'word',
      origin: OFFICE_RELAY_ORIGIN,
      capabilities: ['agent.v1'],
      privateKey: enrollment.privateKey,
    })
    expect(JSON.stringify(database.value)).not.toContain('token')
    expect(JSON.stringify(database.value)).not.toContain('capability_')
    await expect(store.save(enrollment, 'binding_12345678', ['web-fetch.v1'])).rejects.toThrow(
      'binding_invalid',
    )
    expect(OFFICE_BINDING_DATABASE_SCHEMA).toEqual({
      name: 'wiswork-office-pairing',
      version: 1,
      objectStore: 'bindings',
      activeKey: 'active',
    })
  })

  it('loads exact matching records and deletes corrupt, host, capability, and key mismatches', async () => {
    const database = new MemoryBindingDatabase()
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle })
    const enrollment = await store.createEnrollment('word', ['agent.v1'])
    await store.save(enrollment, 'binding_12345678', ['agent.v1'])
    const validRecord = database.value as object

    await expect(store.load('word', ['agent.v1'])).resolves.toMatchObject({
      schemaVersion: 1,
      bindingId: 'binding_12345678',
      host: 'word',
    })

    for (const invalid of [
      { ...validRecord, schemaVersion: 2 },
      { ...validRecord, host: 'excel' },
      { ...validRecord, capabilities: ['web-fetch.v1'] },
      { ...validRecord, privateKey: {} },
      {
        ...validRecord,
        privateKey: {
          type: 'private',
          extractable: false,
          algorithm: { name: 'ECDSA', namedCurve: 'P-256' },
          usages: ['sign'],
        },
      },
      { ...validRecord, unexpected: true },
      { ...validRecord, bindingId: 'x'.repeat(513) },
    ]) {
      database.value = invalid
      const before = database.deletes
      await expect(store.load('word', ['agent.v1'])).resolves.toBeUndefined()
      expect(database.deletes).toBe(before + 1)
    }
  })

  it('signs the fixed domain-separated resume transcript as a raw P-256 signature', async () => {
    const database = new MemoryBindingDatabase()
    const inputs: Uint8Array[] = []
    const subtle = {
      ...crypto.subtle,
      generateKey: crypto.subtle.generateKey.bind(crypto.subtle),
      exportKey: crypto.subtle.exportKey.bind(crypto.subtle),
      sign: async (algorithm: AlgorithmIdentifier, key: CryptoKey, data: BufferSource) => {
        inputs.push(new Uint8Array(data instanceof ArrayBuffer ? data : data.buffer))
        return crypto.subtle.sign(algorithm, key, data)
      },
    } as SubtleCrypto
    const store = createOfficeBindingStore({ database, subtle })
    const enrollment = await store.createEnrollment('powerpoint', ['agent.v1'])
    await store.save(enrollment, 'binding_12345678', ['agent.v1'])
    const binding = await store.load('powerpoint', ['agent.v1'])

    const signature = await store.sign(binding!, 'challenge_12345678')

    expect(new TextDecoder().decode(inputs.at(-1))).toBe(
      'wiswork-office-resume-v1\nbinding_12345678\nchallenge_12345678\nhttps://office.8-216-134-194.sslip.io\nPowerPoint',
    )
    expect(Buffer.from(signature, 'base64url')).toHaveLength(64)
  })

  it('forgets the active record explicitly', async () => {
    const database = new MemoryBindingDatabase()
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle })
    const enrollment = await store.createEnrollment('excel', ['agent.v1'])
    await store.save(enrollment, 'binding_12345678', ['agent.v1'])

    await store.forget()

    expect(database.value).toBeUndefined()
    expect(database.deletes).toBe(1)
  })

  it('reports unavailable persistence instead of treating it as an empty durable store', async () => {
    const database: OfficeBindingDatabase = {
      read: async () => {
        throw new Error('blocked')
      },
      write: async () => undefined,
      delete: async () => undefined,
    }
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle })

    await expect(store.load('word', ['agent.v1'])).rejects.toThrow('binding_storage_unavailable')
  })
})
