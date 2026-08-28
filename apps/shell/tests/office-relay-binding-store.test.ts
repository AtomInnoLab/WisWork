import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  createOfficeRelayBindingStore,
  createElectronOfficeRelayBindingStore,
  type OfficeRelayBinding,
} from '../src/main/office-relay-binding-store'

const temporaryRoots: string[] = []

afterEach(async () => {
  await Promise.all(
    temporaryRoots.splice(0).map((root) => rm(root, { recursive: true, force: true })),
  )
})

const wordBinding = (overrides: Partial<OfficeRelayBinding> = {}): OfficeRelayBinding => ({
  bindingId: 'binding_word_12345678',
  accountId: 'account-one',
  host: 'Word',
  origin: 'https://office.8-216-134-194.sslip.io',
  capabilities: ['agent.v1'],
  createdAt: 1_700_000_000_000,
  ...overrides,
})

function harness(initial: unknown = null) {
  let encrypted =
    initial === null ? null : Buffer.from(`encrypted:${JSON.stringify(initial)}`, 'utf8')
  const writes: Array<{ path: string; data: Uint8Array }> = []
  const rename = vi.fn(async (from: string, to: string) => {
    expect(from).toBe('/profile/office-pairings.enc.tmp')
    expect(to).toBe('/profile/office-pairings.enc')
    encrypted = writes.at(-1)?.data ?? null
  })
  const unlink = vi.fn(async () => {
    encrypted = null
  })
  const store = createOfficeRelayBindingStore({
    path: '/profile/office-pairings.enc',
    platform: 'darwin',
    safeStorage: {
      isEncryptionAvailable: () => true,
      encryptString: (value) => Buffer.from(`encrypted:${value}`, 'utf8'),
      decryptString: (value) => Buffer.from(value).toString('utf8').slice('encrypted:'.length),
    },
    readFile: async () => encrypted,
    writeFile: async (path, data) => {
      writes.push({ path, data })
    },
    rename,
    unlink,
    mkdir: vi.fn(async () => undefined),
  })
  return { store, writes, rename, unlink, encrypted: () => encrypted }
}

describe('Office relay encrypted binding store', () => {
  it('creates the production store at office-pairings.enc under Electron userData', async () => {
    const root = await mkdtemp(join(tmpdir(), 'wiswork-office-pairings-'))
    temporaryRoots.push(root)
    const store = createElectronOfficeRelayBindingStore({
      userDataPath: root,
      platform: 'darwin',
      safeStorage: {
        isEncryptionAvailable: () => true,
        encryptString: (value) => Buffer.from(`encrypted:${value}`),
        decryptString: (value) => Buffer.from(value).toString().slice('encrypted:'.length),
      },
    })
    await store.put(wordBinding())
    await expect(readFile(join(root, 'office-pairings.enc'), 'utf8')).resolves.toContain(
      'encrypted:',
    )
  })

  it('uses safeStorage and an atomic temporary rename in a file separate from auth-session.enc', async () => {
    const { store, writes, rename } = harness()
    await store.put(wordBinding())

    expect(writes).toHaveLength(1)
    expect(writes[0]!.path).toBe('/profile/office-pairings.enc.tmp')
    expect(Buffer.from(writes[0]!.data).toString('utf8')).toContain('encrypted:')
    expect(rename).toHaveBeenCalledOnce()
    await expect(store.listForAccount('account-one')).resolves.toEqual([wordBinding()])
  })

  it('fails closed on unavailable or insecure Linux safeStorage', async () => {
    const base = {
      path: '/profile/office-pairings.enc',
      readFile: async () => null,
      writeFile: async () => undefined,
      rename: async () => undefined,
      unlink: async () => undefined,
      mkdir: async () => undefined,
    }
    const unavailable = createOfficeRelayBindingStore({
      ...base,
      safeStorage: {
        isEncryptionAvailable: () => false,
        encryptString: () => new Uint8Array(),
        decryptString: () => '',
      },
    })
    await expect(unavailable.listForAccount('account-one')).rejects.toThrow(
      'secure_storage_unavailable',
    )

    const basicText = createOfficeRelayBindingStore({
      ...base,
      platform: 'linux',
      safeStorage: {
        isEncryptionAvailable: () => true,
        getSelectedStorageBackend: () => 'basic_text',
        encryptString: () => new Uint8Array(),
        decryptString: () => '',
      },
    })
    await expect(basicText.put(wordBinding())).rejects.toThrow('secure_storage_unavailable')
  })

  it.each([
    {},
    { version: 2, bindings: [], tombstones: [] },
    { version: 1, bindings: [], tombstones: [], extra: true },
    { version: 1, bindings: [wordBinding({ host: 'Outlook' as 'Word' })], tombstones: [] },
    {
      version: 1,
      bindings: [wordBinding({ capabilities: ['pairing-resume.v1'] })],
      tombstones: [],
    },
    { version: 1, bindings: [wordBinding({ accountId: '' })], tombstones: [] },
  ])(
    'deletes an encrypted file whose decrypted payload violates the exact schema',
    async (value) => {
      const { store, unlink } = harness(value)
      await expect(store.listForAccount('account-one')).resolves.toEqual([])
      expect(unlink).toHaveBeenCalledWith('/profile/office-pairings.enc')
    },
  )

  it('returns only exact-account records and enforces twelve live bindings per account', async () => {
    const { store } = harness()
    for (let index = 0; index < 12; index += 1) {
      await store.put(
        wordBinding({
          bindingId: `binding_account_one_${String(index).padStart(2, '0')}`,
          createdAt: 1_700_000_000_000 + index,
        }),
      )
    }
    await store.put(wordBinding({ bindingId: 'binding_account_two_00', accountId: 'account-two' }))
    await expect(
      store.put(wordBinding({ bindingId: 'binding_account_one_overflow' })),
    ).rejects.toThrow('office_binding_limit')
    expect(await store.listForAccount('account-one')).toHaveLength(12)
    expect(await store.listForAccount('account-two')).toEqual([
      wordBinding({ bindingId: 'binding_account_two_00', accountId: 'account-two' }),
    ])
  })

  it('turns cleared active records into account-scoped encrypted tombstones and removes them on ack', async () => {
    const { store } = harness()
    await store.put(wordBinding())
    await store.put(wordBinding({ bindingId: 'binding_other_12345678', accountId: 'account-two' }))

    await store.tombstoneAccount('account-one')
    await expect(store.listForAccount('account-one')).resolves.toEqual([])
    await expect(store.listForAccount('account-two')).resolves.toHaveLength(1)
    await expect(store.listTombstonesForAccount('account-one')).resolves.toEqual([
      {
        bindingId: 'binding_word_12345678',
        accountId: 'account-one',
        createdAt: expect.any(Number),
      },
    ])

    await store.acknowledgeTombstone('account-one', 'binding_word_12345678')
    await expect(store.listTombstonesForAccount('account-one')).resolves.toEqual([])
  })

  it('removes only the invalidated binding without creating a revocation tombstone', async () => {
    const { store } = harness()
    await store.put(wordBinding())
    await store.put(wordBinding({ bindingId: 'binding_excel_12345678', host: 'Excel' }))
    await store.remove('account-one', 'binding_word_12345678')
    await expect(store.listForAccount('account-one')).resolves.toEqual([
      wordBinding({ bindingId: 'binding_excel_12345678', host: 'Excel' }),
    ])
    await expect(store.listTombstonesForAccount('account-one')).resolves.toEqual([])
  })
})
