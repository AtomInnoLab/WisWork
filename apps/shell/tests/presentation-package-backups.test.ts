import { createHash } from 'node:crypto'
import {
  mkdtempSync,
  rmSync,
  symlinkSync,
  mkdirSync,
  writeFileSync,
  renameSync,
  readFileSync,
  existsSync,
  constants,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import { createPresentationPackageBackupService } from '../src/main/presentation-package-backups'
import { createPresentationService } from '../src/main/presentation-service'
import {
  savePackageBackup,
  readPackageBackup,
} from '../../office-addin/src/skills/powerpoint/presentation-package-backup'
const race = vi.hoisted(() => ({
  open: undefined as undefined | ((path: string, flags: number) => void),
  syncOpen: undefined as undefined | ((path: string, flags: number) => void),
  sync: undefined as undefined | ((path: string) => void),
}))
vi.mock('node:fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs')>()
  return {
    ...actual,
    openSync: (
      path: Parameters<typeof actual.openSync>[0],
      flags: Parameters<typeof actual.openSync>[1],
      mode?: number,
    ) => {
      race.syncOpen?.(String(path), Number(flags))
      return actual.openSync(path, flags, mode)
    },
  }
})
vi.mock('node:fs/promises', async (importOriginal) => {
  const actual = await importOriginal<typeof import('node:fs/promises')>()
  return {
    ...actual,
    open: async (
      path: Parameters<typeof actual.open>[0],
      flags: Parameters<typeof actual.open>[1],
      mode?: number,
    ) => {
      race.open?.(String(path), Number(flags))
      const handle = await actual.open(path, flags, mode)
      const sync = handle.sync.bind(handle)
      handle.sync = async () => {
        await sync()
        race.sync?.(String(path))
      }
      return handle
    },
  }
})

const roots: string[] = []
afterEach(() => {
  race.open = undefined
  race.syncOpen = undefined
  race.sync = undefined
})
afterEach(() => roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true })))
const sha = (v: string | Buffer) => createHash('sha256').update(v).digest('hex')
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'master-backup-'))
  roots.push(root)
  const scope = { documentId: 'document-1', changeId: 'change-1', key: 'snapshot' },
    raw = Buffer.from('native snapshot exact bytes')
  const service = createPresentationPackageBackupService({ userDataPath: root })
  const call = (
    operation: string,
    extra: Record<string, unknown> = {},
    signal = new AbortController().signal,
  ) => service({ operation, ...scope, ...extra }, signal)
  const begin = () => call('package_backup_begin', { sha256: sha(raw), sizeBytes: raw.length })
  return { root, scope, raw, call, begin }
}
it('persists exact blobs and recovers chunk/finish ACK replay after restart', async () => {
  const f = fixture()
  await f.begin()
  const chunk = { offset: 0, base64: f.raw.toString('base64') }
  await f.call('package_backup_chunk', chunk)
  await f.call('package_backup_chunk', chunk)
  expect(await f.call('package_backup_finish')).toMatchObject({
    ...f.scope,
    status: 'ready',
    receivedBytes: f.raw.length,
  })
  expect(await f.begin()).toMatchObject({ status: 'ready' })
  expect(await f.call('package_backup_finish')).toMatchObject({ status: 'ready' })
  const read = (await createPresentationPackageBackupService({ userDataPath: f.root })(
    { operation: 'package_backup_read', ...f.scope, offset: 0, length: f.raw.length },
    new AbortController().signal,
  )) as { base64: string }
  expect(Buffer.from(read.base64, 'base64')).toEqual(f.raw)
})
it('releases only the exact backup identity and returns capacity', async () => {
  const f = fixture()
  await f.begin()
  await f.call('package_backup_chunk', { offset: 0, base64: f.raw.toString('base64') })
  await f.call('package_backup_finish')
  await expect(
    f.call('package_backup_release', { sha256: sha('wrong'), sizeBytes: f.raw.length }),
  ).rejects.toThrow('presentation_package_backup_invalid')
  expect(await f.call('package_backup_status')).toMatchObject({ status: 'ready' })
  const released = await createPresentationService({ userDataPath: f.root })(
    {
      operation: 'package_backup_release',
      ...f.scope,
      sha256: sha(f.raw),
      sizeBytes: f.raw.length,
    },
    new AbortController().signal,
  )
  expect(JSON.parse(Buffer.from(released).toString())).toMatchObject({ released: true })
  await expect(f.call('package_backup_status')).rejects.toThrow(
    'presentation_package_backup_invalid',
  )
  expect(
    await createPresentationPackageBackupService({ userDataPath: f.root })(
      {
        operation: 'package_backup_list',
        documentId: f.scope.documentId,
        changeId: f.scope.changeId,
      },
      new AbortController().signal,
    ),
  ).toMatchObject({ backups: [] })
})
it('sweeps interrupted releases after the directory left document quota', async () => {
  const f = fixture()
  await f.begin()
  const root = join(f.root, 'presentation-package-backups'),
    active = join(
      root,
      sha(f.scope.documentId),
      sha(JSON.stringify([f.scope.changeId, f.scope.key])),
    ),
    released = join(root, '.released-00000000-0000-4000-8000-000000000001')
  renameSync(active, released)
  await f.begin()
  expect(existsSync(released)).toBe(false)
  const partial = join(root, '.released-00000000-0000-4000-8000-000000000002')
  renameSync(active, partial)
  rmSync(join(partial, 'blob'))
  await f.begin()
  expect(existsSync(partial)).toBe(false)
  expect(await f.call('package_backup_status')).toMatchObject({ status: 'uploading' })
})
it('refuses a substituted released directory without touching its target', async () => {
  const f = fixture()
  await f.begin()
  const outside = join(f.root, 'outside-release-target')
  mkdirSync(outside)
  writeFileSync(join(outside, 'sentinel'), 'keep')
  symlinkSync(
    outside,
    join(f.root, 'presentation-package-backups', '.released-00000000-0000-4000-8000-000000000003'),
    'dir',
  )
  await expect(f.begin()).rejects.toThrow('presentation_package_backup_invalid')
  expect(readFileSync(join(outside, 'sentinel'), 'utf8')).toBe('keep')
})
it('rejects altered scopes, overlap, malformed requests, unready reads and corrupted bytes', async () => {
  const f = fixture()
  await f.begin()
  await expect(
    f.call('package_backup_begin', { sha256: sha('other'), sizeBytes: f.raw.length }),
  ).rejects.toThrow('presentation_package_backup_invalid')
  await expect(f.call('package_backup_read', { offset: 0, length: 1 })).rejects.toThrow(
    'presentation_package_backup_invalid',
  )
  await f.call('package_backup_chunk', { offset: 0, base64: f.raw.toString('base64') })
  await expect(
    f.call('package_backup_chunk', {
      offset: 0,
      base64: Buffer.from('different').toString('base64'),
    }),
  ).rejects.toThrow('presentation_package_backup_invalid')
  await expect(f.call('package_backup_status', { key: '../snapshot' })).rejects.toThrow(
    'presentation_package_backup_invalid',
  )
  await expect(
    f.call('package_backup_begin', { sizeBytes: 8 * 1024 * 1024 + 1, sha256: sha('x') }),
  ).rejects.toThrow('presentation_package_backup_invalid')
  await f.call('package_backup_finish')
  writeFileSync(
    join(
      f.root,
      'presentation-package-backups',
      sha(f.scope.documentId),
      sha(JSON.stringify([f.scope.changeId, f.scope.key])),
      'blob',
    ),
    Buffer.alloc(f.raw.length),
  )
  await expect(f.call('package_backup_status')).rejects.toThrow(
    'presentation_package_backup_invalid',
  )
})
it('supports over sixteen blobs and lists without truncation through PC service', async () => {
  const f = fixture()
  for (let i = 0; i < 20; i++)
    await f.call('package_backup_begin', { key: `page-${i}`, sha256: sha('a'), sizeBytes: 1 })
  const response = await createPresentationService({ userDataPath: f.root })(
    {
      operation: 'package_backup_list',
      documentId: f.scope.documentId,
      changeId: f.scope.changeId,
    },
    new AbortController().signal,
  )
  expect(JSON.parse(Buffer.from(response).toString()).backups).toHaveLength(20)
})
it('pages a document backup inventory without requiring a remembered change ID', async () => {
  const f = fixture(),
    service = createPresentationService({ userDataPath: f.root })
  for (let i = 0; i < 66; i++)
    await f.call('package_backup_begin', {
      changeId: `change-${i}`,
      sha256: sha(f.raw),
      sizeBytes: f.raw.length,
    })
  const inventory = async (after: string) =>
    JSON.parse(
      Buffer.from(
        await service(
          { operation: 'package_backup_inventory', documentId: f.scope.documentId, after },
          new AbortController().signal,
        ),
      ).toString(),
    ) as { backups: { changeId: string; key: string }[]; nextCursor?: string }
  const first = await inventory('')
  expect(first.backups).toHaveLength(64)
  expect(first.nextCursor).toMatch(/^[a-f0-9]{64}$/)
  const second = await inventory(first.nextCursor!)
  expect(second.backups).toHaveLength(2)
  expect(second.nextCursor).toBeUndefined()
  expect(new Set([...first.backups, ...second.backups].map((b) => b.changeId)).size).toBe(66)
  const invalid = await service(
    { operation: 'package_backup_inventory', documentId: f.scope.documentId, after: '../' },
    new AbortController().signal,
  )
  expect(JSON.parse(Buffer.from(invalid).toString())).toEqual({
    error: 'presentation_package_backup_invalid',
  })
})
it('serializes conflicts and snapshots caller fields before awaiting', async () => {
  const f = fixture(),
    service = createPresentationPackageBackupService({ userDataPath: f.root })
  const body = {
    operation: 'package_backup_begin',
    ...f.scope,
    sha256: sha(f.raw),
    sizeBytes: f.raw.length,
  }
  const first = service(body, new AbortController().signal)
  body.key = 'image-1'
  await first
  expect(await f.call('package_backup_status')).toMatchObject({ key: 'snapshot' })
  const settled = await Promise.allSettled([
    f.call('package_backup_begin', { key: 'page-1', sha256: sha('a'), sizeBytes: 1 }),
    f.call('package_backup_begin', { key: 'page-1', sha256: sha('b'), sizeBytes: 1 }),
  ])
  expect(settled.filter((r) => r.status === 'fulfilled')).toHaveLength(1)
})
it('refuses symlink roots and cancellation', async () => {
  const f = fixture(),
    target = join(f.root, 'outside')
  mkdirSync(target)
  symlinkSync(target, join(f.root, 'presentation-package-backups'), 'dir')
  await expect(f.begin()).rejects.toThrow('presentation_package_backup_invalid')
  const abort = new AbortController()
  abort.abort()
  await expect(f.call('package_backup_status', {}, abort.signal)).rejects.toThrow('aborted')
})
it('reserves the full 2 GiB document quota without blocking existing status at capacity', async () => {
  const f = fixture()
  await f.begin()
  const document = join(f.root, 'presentation-package-backups', sha(f.scope.documentId))
  // Synthetic metadata fills reserved capacity without allocating 2 GiB test data.
  for (let i = 0; i < 256; i++) {
    const m = {
      documentId: f.scope.documentId,
      changeId: 'capacity',
      key: `page-${i}`,
      sha256: sha('x'),
      sizeBytes: i === 255 ? 8 * 1024 * 1024 - f.raw.length : 8 * 1024 * 1024,
      status: 'uploading',
    }
    const dir = join(document, sha(JSON.stringify([m.changeId, m.key])))
    mkdirSync(dir)
    writeFileSync(join(dir, 'metadata.json'), JSON.stringify(m))
  }
  await expect(
    f.call('package_backup_begin', { key: 'image-1', sha256: sha('a'), sizeBytes: 1 }),
  ).rejects.toThrow('presentation_package_backup_capacity')
  expect(await f.call('package_backup_status')).toMatchObject({ key: 'snapshot', receivedBytes: 0 })
})
it('rejects symlink blob data and exact request field expansion', async () => {
  const f = fixture()
  await f.begin()
  await expect(f.call('package_backup_status', { extra: true })).rejects.toThrow(
    'presentation_package_backup_invalid',
  )
  const dir = join(
    f.root,
    'presentation-package-backups',
    sha(f.scope.documentId),
    sha(JSON.stringify([f.scope.changeId, f.scope.key])),
  )
  rmSync(join(dir, 'blob'))
  writeFileSync(join(f.root, 'foreign'), 'secret')
  symlinkSync(join(f.root, 'foreign'), join(dir, 'blob'))
  await expect(f.call('package_backup_status')).rejects.toThrow(
    'presentation_package_backup_invalid',
  )
})

it('roundtrips browser client through actual PC service and immutable ready receipts', async () => {
  const f = fixture(),
    service = createPresentationService({ userDataPath: f.root })
  const request = async (body: unknown, signal?: AbortSignal) =>
    new Response(Buffer.from(await service(body, signal ?? new AbortController().signal)), {
      status: 200,
    })
  const bytes = Uint8Array.from({ length: 300000 }, (_, i) => i % 251)
  const scope = { request, documentId: f.scope.documentId, changeId: f.scope.changeId }
  const backup = await savePackageBackup({ ...scope, key: 'image-1', bytes })
  expect(await readPackageBackup({ ...scope, backup })).toEqual(bytes)
  expect(await savePackageBackup({ ...scope, key: 'image-1', bytes })).toEqual(backup)
})

function substitute(f: ReturnType<typeof fixture>) {
  const parent = join(
    f.root,
    'presentation-package-backups',
    sha(f.scope.documentId),
    sha(JSON.stringify([f.scope.changeId, f.scope.key])),
  )
  const foreign = join(f.root, 'foreign')
  mkdirSync(foreign)
  writeFileSync(join(foreign, 'blob'), Buffer.alloc(0))
  writeFileSync(join(foreign, 'metadata.json'), 'foreign')
  renameSync(parent, parent + '-old')
  symlinkSync(foreign, parent, 'dir')
  return foreign
}
it.each(['chunk', 'read', 'finish', 'publish'])(
  'refuses persistent parent substitution during %s without foreign bytes',
  async (scenario) => {
    const f = fixture()
    await f.begin()
    if (scenario !== 'chunk') {
      await f.call('package_backup_chunk', { offset: 0, base64: f.raw.toString('base64') })
    }
    if (scenario === 'read') await f.call('package_backup_finish')
    let foreign: string | undefined
    if (scenario === 'publish')
      race.sync = (path) => {
        if (path.endsWith('.tmp')) {
          race.sync = undefined
          foreign = substitute(f)
        }
      }
    else
      race.open = (path, flags) => {
        const trigger =
          scenario === 'chunk'
            ? flags === (constants.O_RDWR | constants.O_APPEND | constants.O_NOFOLLOW)
            : scenario === 'read'
              ? path.endsWith('/blob')
              : path.endsWith('.tmp')
        if (trigger) {
          race.open = undefined
          foreign = substitute(f)
          if (scenario === 'read') writeFileSync(join(foreign, 'blob'), f.raw)
        }
      }
    const operation =
      scenario === 'chunk'
        ? 'package_backup_chunk'
        : scenario === 'read'
          ? 'package_backup_read'
          : 'package_backup_finish'
    const extra =
      scenario === 'chunk'
        ? { offset: 0, base64: f.raw.toString('base64') }
        : scenario === 'read'
          ? { offset: 0, length: f.raw.length }
          : {}
    await expect(f.call(operation, extra)).rejects.toThrow('presentation_package_backup_invalid')
    expect(foreign).toBeDefined()
    expect(readFileSync(join(foreign!, 'blob'))).toEqual(
      scenario === 'read' ? f.raw : Buffer.alloc(0),
    )
    expect(readFileSync(join(foreign!, 'metadata.json')).toString()).toBe('foreign')
  },
)
it('keeps verified own blobs readable without scanning unrelated metadata; allocation still validates all reservations', async () => {
  const f = fixture()
  await f.begin()
  await f.call('package_backup_chunk', { offset: 0, base64: f.raw.toString('base64') })
  await f.call('package_backup_finish')
  await f.call('package_backup_begin', { key: 'page-1', sha256: sha('a'), sizeBytes: 1 })
  const unrelated = join(
    f.root,
    'presentation-package-backups',
    sha(f.scope.documentId),
    sha(JSON.stringify([f.scope.changeId, 'page-1'])),
    'metadata.json',
  )
  writeFileSync(unrelated, 'corrupt unrelated metadata')
  expect(await f.call('package_backup_status')).toMatchObject({ status: 'ready' })
  expect(
    await f.call('package_backup_chunk', { offset: 0, base64: f.raw.toString('base64') }),
  ).toMatchObject({ status: 'ready' })
  expect(await f.call('package_backup_finish')).toMatchObject({ status: 'ready' })
  expect(await f.call('package_backup_read', { offset: 0, length: f.raw.length })).toMatchObject({
    base64: f.raw.toString('base64'),
  })
  await expect(
    f.call('package_backup_begin', { key: 'page-2', sha256: sha('b'), sizeBytes: 1 }),
  ).rejects.toThrow('presentation_package_backup_invalid')
  await expect(
    createPresentationPackageBackupService({ userDataPath: f.root })(
      {
        operation: 'package_backup_list',
        documentId: f.scope.documentId,
        changeId: f.scope.changeId,
      },
      new AbortController().signal,
    ),
  ).rejects.toThrow('presentation_package_backup_invalid')
})
it('checks full bounded quota metadata without one asynchronous descriptor open per reservation', async () => {
  const f = fixture()
  for (let i = 0; i < 4; i++)
    await f.call('package_backup_begin', { key: `page-${i}`, sha256: sha('a'), sizeBytes: 1 })
  let asynchronousMetadataOpens = 0
  race.open = (path) => {
    if (path.endsWith('/metadata.json')) asynchronousMetadataOpens++
  }
  await f.call('package_backup_begin', { key: 'page-4', sha256: sha('a'), sizeBytes: 1 })
  expect(asynchronousMetadataOpens).toBe(0)
})

it('refuses parent substitution at synchronous metadata open even with an identical valid record', async () => {
  const f = fixture()
  await f.begin()
  const original = join(
    f.root,
    'presentation-package-backups',
    sha(f.scope.documentId),
    sha(JSON.stringify([f.scope.changeId, f.scope.key])),
    'metadata.json',
  )
  const value = readFileSync(original)
  let foreign: string | undefined
  race.syncOpen = (path) => {
    if (path.endsWith('/metadata.json')) {
      race.syncOpen = undefined
      foreign = substitute(f)
      writeFileSync(join(foreign, 'metadata.json'), value)
    }
  }
  await expect(f.call('package_backup_status')).rejects.toThrow(
    'presentation_package_backup_invalid',
  )
  expect(readFileSync(join(foreign!, 'metadata.json'))).toEqual(value)
})

it('isolates master/package namespaces and rejects arbitrary storage directories', async () => {
  const { createPresentationMasterBackupService } =
    await import('../src/main/presentation-master-backups')
  const root = mkdtempSync(join(tmpdir(), 'package-isolation-'))
  roots.push(root)
  const signal = new AbortController().signal
  const master = createPresentationMasterBackupService({ userDataPath: root })
  const packages = createPresentationPackageBackupService({ userDataPath: root })
  const identity = { documentId: 'deck', changeId: 'change', key: 'snapshot' }
  await master(
    {
      ...identity,
      operation: 'master_backup_begin',
      sha256: createHash('sha256').update('m').digest('hex'),
      sizeBytes: 1,
    },
    signal,
  )
  await expect(
    packages({ ...identity, operation: 'package_backup_status' }, signal),
  ).rejects.toThrow('presentation_package_backup_invalid')
  await expect(
    packages({ ...identity, operation: 'master_backup_status' }, signal),
  ).rejects.toThrow('presentation_package_backup_invalid')
  await expect(master({ ...identity, operation: 'package_backup_status' }, signal)).rejects.toThrow(
    'presentation_master_backup_invalid',
  )
  expect(() =>
    createPresentationMasterBackupService({
      userDataPath: root,
      storageDirectory: '../outside' as never,
    }),
  ).toThrow('presentation_master_backup_invalid')
})
