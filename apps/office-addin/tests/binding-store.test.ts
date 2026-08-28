import { describe, expect, it } from 'vitest'
import {
  OFFICE_BINDING_DATABASE_SCHEMA,
  OFFICE_RELAY_ORIGIN,
  createIndexedDbOfficeBindingDatabase,
  createOfficeBindingStore,
  type OfficeBindingDatabase,
} from '../src/relay/binding-store.js'

class MemoryBindingDatabase implements OfficeBindingDatabase {
  readonly values = new Map<string, unknown>()
  readonly claims = new Map<string, { owner: string; expiresAt: number }>()
  readonly blocked = new Map<string, string>()
  deletes = 0

  async read(host: 'word' | 'excel' | 'powerpoint') {
    return { record: this.values.get(host), blockedBindingId: this.blocked.get(host) }
  }

  async claim(
    host: 'word' | 'excel' | 'powerpoint',
    owner: string,
    expiresAt: number,
  ): Promise<boolean> {
    const existing = this.claims.get(host)
    if (this.values.has(host) || (existing && existing.expiresAt > Date.now())) return false
    this.claims.set(host, { owner, expiresAt })
    return true
  }

  async stage(
    host: 'word' | 'excel' | 'powerpoint',
    owner: string,
    value: unknown,
  ): Promise<boolean> {
    if (this.values.has(host) || this.claims.get(host)?.owner !== owner) return false
    this.values.set(host, value)
    this.claims.delete(host)
    return true
  }

  async activate(
    host: 'word' | 'excel' | 'powerpoint',
    owner: string,
    bindingId: string,
  ): Promise<boolean> {
    const record = this.values.get(host) as Record<string, unknown> | undefined
    if (
      record?.bindingId !== bindingId ||
      record.pending !== true ||
      record.enrollmentOwner !== owner
    )
      return false
    const {
      pending: _pending,
      enrollmentOwner: _owner,
      pendingExpiresAt: _expiresAt,
      ...live
    } = record
    this.values.set(host, live)
    return true
  }

  async release(
    host: 'word' | 'excel' | 'powerpoint',
    owner: string,
    bindingId?: string,
  ): Promise<void> {
    if (this.claims.get(host)?.owner === owner) this.claims.delete(host)
    const record = this.values.get(host) as Record<string, unknown> | undefined
    if (
      record?.pending === true &&
      record.enrollmentOwner === owner &&
      record.bindingId === bindingId
    )
      this.values.delete(host)
  }

  async block(host: 'word' | 'excel' | 'powerpoint', bindingId: string): Promise<boolean> {
    const record = this.values.get(host) as Record<string, unknown> | undefined
    if (this.blocked.get(host) === bindingId) return true
    if (record?.bindingId !== bindingId) return false
    this.blocked.set(host, bindingId)
    return true
  }

  async deleteBlocked(host: 'word' | 'excel' | 'powerpoint', bindingId: string): Promise<boolean> {
    const record = this.values.get(host) as Record<string, unknown> | undefined
    if (record?.bindingId !== bindingId || this.blocked.get(host) !== bindingId) return false
    this.deletes += 1
    this.values.delete(host)
    this.claims.delete(host)
    this.blocked.delete(host)
    return true
  }

  async deleteIf(
    host: 'word' | 'excel' | 'powerpoint',
    predicate: (value: unknown) => boolean,
  ): Promise<boolean> {
    const record = this.values.get(host)
    if (!predicate(record)) return false
    this.deletes += 1
    this.values.delete(host)
    return true
  }
}

class DeterministicIdbFactory {
  readonly values = new Map<IDBValidKey, unknown>()
  readonly transactions: IDBTransactionMode[] = []
  private created = false
  private version = 0
  private writeTail = Promise.resolve()

  constructor(legacyRecord?: unknown) {
    if (legacyRecord !== undefined) {
      this.created = true
      this.version = 1
      this.values.set('active', legacyRecord)
    }
  }

  open(_name: string, requestedVersion = 1): IDBOpenDBRequest {
    const request = {} as IDBOpenDBRequest
    const database = {
      objectStoreNames: { contains: () => this.created },
      createObjectStore: () => {
        this.created = true
      },
      close: () => undefined,
      transaction: (_store: string, mode: IDBTransactionMode) => {
        this.transactions.push(mode)
        const transaction = {} as IDBTransaction
        const writes = mode === 'readwrite' || mode === 'versionchange'
        const start = writes ? this.writeTail : Promise.resolve()
        let release: () => void = () => undefined
        if (writes) {
          const completed = new Promise<void>((resolve) => (release = resolve))
          this.writeTail = start.then(() => completed)
        }
        let finished = false
        let pending = 0
        const run = (operation: () => void) => void start.then(() => queueMicrotask(operation))
        const finish = (kind: 'complete' | 'abort') => {
          if (finished) return
          finished = true
          queueMicrotask(() => {
            if (kind === 'complete') transaction.oncomplete?.(new Event('complete'))
            else transaction.onabort?.(new Event('abort'))
            release()
          })
        }
        const completeWhenIdle = () =>
          queueMicrotask(() => queueMicrotask(() => pending === 0 && finish('complete')))
        transaction.abort = () => run(() => finish('abort'))
        transaction.objectStore = () =>
          ({
            get: (key: IDBValidKey) => {
              const result = {} as IDBRequest
              pending += 1
              run(() => {
                Object.defineProperty(result, 'result', { value: this.values.get(key) })
                result.onsuccess?.(new Event('success'))
                pending -= 1
                completeWhenIdle()
              })
              return result
            },
            put: (value: unknown, key: IDBValidKey) => {
              pending += 1
              run(() => {
                this.values.set(key, value)
                pending -= 1
                completeWhenIdle()
              })
              return {} as IDBRequest
            },
            delete: (key: IDBValidKey) => {
              pending += 1
              run(() => {
                this.values.delete(key)
                pending -= 1
                completeWhenIdle()
              })
              return {} as IDBRequest
            },
          }) as IDBObjectStore
        return transaction
      },
    } as unknown as IDBDatabase
    queueMicrotask(() => {
      Object.defineProperty(request, 'result', { value: database })
      if (this.version < requestedVersion) {
        const oldVersion = this.version
        const upgrade = database.transaction(
          OFFICE_BINDING_DATABASE_SCHEMA.objectStore,
          'versionchange',
        )
        Object.defineProperty(request, 'transaction', {
          value: upgrade,
        })
        request.onupgradeneeded?.({ oldVersion } as IDBVersionChangeEvent)
        this.version = requestedVersion
        upgrade.oncomplete = () => request.onsuccess?.(new Event('success'))
        return
      }
      request.onsuccess?.(new Event('success'))
    })
    return request
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

    await store.stage(enrollment, 'binding_12345678', ['agent.v1'])
    expect(database.values.get('word')).toEqual({
      schemaVersion: 1,
      bindingId: 'binding_12345678',
      host: 'word',
      origin: OFFICE_RELAY_ORIGIN,
      capabilities: ['agent.v1'],
      privateKey: enrollment.privateKey,
      pending: true,
      enrollmentOwner: enrollment.publicKey,
      pendingExpiresAt: expect.any(Number),
    })
    await expect(store.load('word', ['agent.v1'])).resolves.toBeUndefined()
    expect(database.values.has('word')).toBe(true)

    const activated = await store.activate(enrollment, 'binding_12345678', ['agent.v1'])
    expect(activated).toMatchObject({ bindingId: 'binding_12345678', host: 'word' })
    expect(database.values.get('word')).toEqual({
      schemaVersion: 1,
      bindingId: 'binding_12345678',
      host: 'word',
      origin: OFFICE_RELAY_ORIGIN,
      capabilities: ['agent.v1'],
      privateKey: enrollment.privateKey,
    })
    expect(JSON.stringify(database.values.get('word'))).not.toContain('token')
    expect(JSON.stringify(database.values.get('word'))).not.toContain('capability_')
    await expect(store.stage(enrollment, 'binding_12345678', ['web-fetch.v1'])).rejects.toThrow(
      'binding_invalid',
    )
    expect(OFFICE_BINDING_DATABASE_SCHEMA).toEqual({
      name: 'wiswork-office-pairing',
      version: 2,
      objectStore: 'bindings',
      activeKeyPrefix: 'active:',
      enrollmentKeyPrefix: 'enrollment:',
    })
  })

  it('loads exact matching records and deletes corrupt, host, capability, and key mismatches', async () => {
    const database = new MemoryBindingDatabase()
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle })
    const enrollment = await store.createEnrollment('word', ['agent.v1'])
    await store.stage(enrollment, 'binding_12345678', ['agent.v1'])
    await store.activate(enrollment, 'binding_12345678', ['agent.v1'])
    const validRecord = database.values.get('word') as object

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
      database.values.set('word', invalid)
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
    await store.stage(enrollment, 'binding_12345678', ['agent.v1'])
    await store.activate(enrollment, 'binding_12345678', ['agent.v1'])
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
    await store.stage(enrollment, 'binding_12345678', ['agent.v1'])
    await store.activate(enrollment, 'binding_12345678', ['agent.v1'])

    await store.forget('excel', 'binding_12345678')

    expect(database.values.get('excel')).toBeUndefined()
    expect(database.deletes).toBe(1)
  })

  it('reports unavailable persistence instead of treating it as an empty durable store', async () => {
    const database: OfficeBindingDatabase = {
      read: async () => {
        throw new Error('blocked')
      },
      claim: async () => false,
      stage: async () => false,
      activate: async () => false,
      release: async () => undefined,
      block: async () => false,
      deleteBlocked: async () => false,
      deleteIf: async () => false,
    }
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle })

    await expect(store.load('word', ['agent.v1'])).rejects.toThrow('binding_storage_unavailable')
  })

  it('persists and reloads a non-exportable CryptoKey through the IndexedDB adapter contract', async () => {
    const factory = new DeterministicIdbFactory()
    const first = createOfficeBindingStore({
      database: createIndexedDbOfficeBindingDatabase(factory as unknown as IDBFactory),
      subtle: crypto.subtle,
    })
    const enrollment = await first.createEnrollment('word', ['agent.v1'])
    await first.stage(enrollment, 'binding_12345678', ['agent.v1'])
    await first.activate(enrollment, 'binding_12345678', ['agent.v1'])

    const reloaded = createOfficeBindingStore({
      database: createIndexedDbOfficeBindingDatabase(factory as unknown as IDBFactory),
      subtle: crypto.subtle,
    })
    const record = await reloaded.load('word', ['agent.v1'])

    expect(record?.privateKey).toBe(enrollment.privateKey)
    expect(record?.privateKey.extractable).toBe(false)
    expect(await reloaded.sign(record!, 'challenge_12345678')).toMatch(/^[A-Za-z0-9_-]+$/)
    await reloaded.forget('word', 'binding_12345678')
    await expect(reloaded.load('word', ['agent.v1'])).resolves.toBeUndefined()
    expect(factory.transactions).toContain('readonly')
  })

  it('migrates the legacy active record into only its host slot without exporting the key', async () => {
    const pair = (await crypto.subtle.generateKey({ name: 'ECDSA', namedCurve: 'P-256' }, false, [
      'sign',
      'verify',
    ])) as CryptoKeyPair
    const legacy = {
      schemaVersion: 1,
      bindingId: 'binding_legacy_12345678',
      host: 'excel',
      origin: OFFICE_RELAY_ORIGIN,
      capabilities: ['agent.v1'],
      privateKey: pair.privateKey,
    }
    const factory = new DeterministicIdbFactory(legacy)
    const store = createOfficeBindingStore({
      database: createIndexedDbOfficeBindingDatabase(factory as unknown as IDBFactory),
      subtle: crypto.subtle,
    })

    await expect(store.load('excel', ['agent.v1'])).resolves.toMatchObject({
      bindingId: 'binding_legacy_12345678',
      host: 'excel',
    })
    await expect(store.load('word', ['agent.v1'])).resolves.toBeUndefined()
    expect(factory.values.has('active')).toBe(false)
    expect(factory.values.get('active:excel')).toEqual(legacy)
    await expect(crypto.subtle.exportKey('pkcs8', pair.privateKey)).rejects.toThrow()
  })

  it('allows only one concurrent same-host claim through the IndexedDB adapter contract', async () => {
    const factory = new DeterministicIdbFactory()
    const database = createIndexedDbOfficeBindingDatabase(factory as unknown as IDBFactory)

    const claims = await Promise.all([
      database.claim('word', 'owner_one', Date.now() + 60_000),
      database.claim('word', 'owner_two', Date.now() + 60_000),
    ])

    expect(claims.filter(Boolean)).toHaveLength(1)
    expect(claims.filter((claimed) => !claimed)).toHaveLength(1)
  })

  it('keeps independent live bindings for Word, Excel, and PowerPoint in one origin', async () => {
    const database = new MemoryBindingDatabase()
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle })

    for (const host of ['word', 'excel', 'powerpoint'] as const) {
      const next = await store.createEnrollment(host, ['agent.v1'])
      await store.stage(next, `binding_${host}_12345678`, ['agent.v1'])
      await store.activate(next, `binding_${host}_12345678`, ['agent.v1'])
    }

    await expect(store.load('word', ['agent.v1'])).resolves.toMatchObject({ host: 'word' })
    await expect(store.load('excel', ['agent.v1'])).resolves.toMatchObject({ host: 'excel' })
    await expect(store.load('powerpoint', ['agent.v1'])).resolves.toMatchObject({
      host: 'powerpoint',
    })
    await store.forget('excel', 'binding_excel_12345678')
    await expect(store.load('word', ['agent.v1'])).resolves.toBeDefined()
    await expect(store.load('excel', ['agent.v1'])).resolves.toBeUndefined()
    await expect(store.load('powerpoint', ['agent.v1'])).resolves.toBeDefined()
  })

  it('leases first enrollment per host and releases a losing or abandoned enrollment safely', async () => {
    const database = new MemoryBindingDatabase()
    const first = createOfficeBindingStore({ database, subtle: crypto.subtle })
    const second = createOfficeBindingStore({ database, subtle: crypto.subtle })
    const enrollment = await first.createEnrollment('word', ['agent.v1'])

    await expect(second.createEnrollment('word', ['agent.v1'])).rejects.toThrow(
      'binding_enrollment_in_progress',
    )
    await expect(second.createEnrollment('excel', ['agent.v1'])).resolves.toMatchObject({
      host: 'excel',
    })

    await first.abort(enrollment)
    await expect(second.createEnrollment('word', ['agent.v1'])).resolves.toMatchObject({
      host: 'word',
    })
  })

  it('uses compare-and-set so a staged binding cannot overwrite an existing host binding', async () => {
    const database = new MemoryBindingDatabase()
    const first = createOfficeBindingStore({ database, subtle: crypto.subtle })
    const enrollment = await first.createEnrollment('word', ['agent.v1'])
    database.values.set('word', {
      schemaVersion: 1,
      bindingId: 'binding_existing_12345678',
      host: 'word',
      origin: OFFICE_RELAY_ORIGIN,
      capabilities: ['agent.v1'],
      privateKey: enrollment.privateKey,
    })

    await expect(first.stage(enrollment, 'binding_loser_12345678', ['agent.v1'])).rejects.toThrow(
      'binding_enrollment_conflict',
    )
    expect((database.values.get('word') as { bindingId: string }).bindingId).toBe(
      'binding_existing_12345678',
    )
  })

  it('expires a stranded pending key so a crashed enrollment cannot block the host forever', async () => {
    const database = new MemoryBindingDatabase()
    let now = 1_000
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle, now: () => now })
    const enrollment = await store.createEnrollment('word', ['agent.v1'])
    await store.stage(enrollment, 'binding_pending_12345678', ['agent.v1'])

    await expect(store.load('word', ['agent.v1'])).resolves.toBeUndefined()
    expect(database.values.has('word')).toBe(true)

    now += 180_001
    await expect(store.load('word', ['agent.v1'])).resolves.toBeUndefined()
    expect(database.values.has('word')).toBe(false)
    await expect(store.createEnrollment('word', ['agent.v1'])).resolves.toMatchObject({
      host: 'word',
    })
  })

  it('rejects a pending expiry beyond the bounded enrollment window', async () => {
    const database = new MemoryBindingDatabase()
    const now = 1_000
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle, now: () => now })
    const enrollment = await store.createEnrollment('word', ['agent.v1'])
    await store.stage(enrollment, 'binding_future_12345678', ['agent.v1'])
    ;(database.values.get('word') as { pendingExpiresAt: number }).pendingExpiresAt = now + 180_001

    await expect(store.load('word', ['agent.v1'])).resolves.toBeUndefined()
    expect(database.values.has('word')).toBe(false)
  })

  it('durably blocks resume before deletion and clears the marker only after a retry succeeds', async () => {
    const database = new MemoryBindingDatabase()
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle })
    const enrollment = await store.createEnrollment('word', ['agent.v1'])
    await store.stage(enrollment, 'binding_blocked_12345678', ['agent.v1'])
    await store.activate(enrollment, 'binding_blocked_12345678', ['agent.v1'])
    const deleteBlocked = database.deleteBlocked.bind(database)
    database.deleteBlocked = async () => {
      throw new Error('disk_failure')
    }

    await expect(store.forget('word', 'binding_blocked_12345678')).rejects.toThrow('disk_failure')
    expect(database.values.has('word')).toBe(true)
    expect(database.blocked.get('word')).toBe('binding_blocked_12345678')

    const reloaded = createOfficeBindingStore({ database, subtle: crypto.subtle })
    await expect(reloaded.load('word', ['agent.v1'])).resolves.toBeUndefined()
    database.deleteBlocked = deleteBlocked
    await reloaded.forget('word')
    expect(database.values.has('word')).toBe(false)
    expect(database.blocked.has('word')).toBe(false)
  })

  it('does not let stale corrupt cleanup delete a newly committed live record', async () => {
    const database = new MemoryBindingDatabase()
    database.values.set('word', {
      schemaVersion: 1,
      bindingId: 'binding_corrupt_12345678',
      host: 'word',
      origin: OFFICE_RELAY_ORIGIN,
      capabilities: ['agent.v1'],
      privateKey: {},
    })
    const pair = (await crypto.subtle.generateKey({ name: 'ECDSA', namedCurve: 'P-256' }, false, [
      'sign',
      'verify',
    ])) as CryptoKeyPair
    const replacement = {
      schemaVersion: 1,
      bindingId: 'binding_replacement_12345678',
      host: 'word',
      origin: OFFICE_RELAY_ORIGIN,
      capabilities: ['agent.v1'],
      privateKey: pair.privateKey,
    }
    const deleteIf = database.deleteIf.bind(database)
    database.deleteIf = async (host, predicate) => {
      database.values.set(host, replacement)
      return deleteIf(host, predicate)
    }
    const store = createOfficeBindingStore({ database, subtle: crypto.subtle })

    await expect(store.load('word', ['agent.v1'])).resolves.toBeUndefined()
    expect(database.values.get('word')).toBe(replacement)
    expect(database.deletes).toBe(0)
  })
})
