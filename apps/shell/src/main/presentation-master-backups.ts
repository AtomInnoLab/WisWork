import { createHash, randomUUID } from 'node:crypto'
import {
  constants,
  closeSync,
  openSync,
  readFileSync,
  fstatSync,
  lstatSync,
  renameSync,
  rmSync,
  rmdirSync,
  type Stats,
} from 'node:fs'
import { mkdir, open, readdir } from 'node:fs/promises'
import { join, resolve } from 'node:path'

const MAX_BLOB = 8 * 1024 * 1024
const CHUNK = 128 * 1024
const MAX_BLOBS = 4096
const MAX_DOCUMENT = 2 * 1024 * 1024 * 1024
const locks = new Map<string, Promise<void>>()
const sweepLocks = new Map<string, Promise<void>>()
const hash = (value: string | Uint8Array) => createHash('sha256').update(value).digest('hex')
function invalid(): never {
  throw new Error('presentation_master_backup_invalid')
}
const check = (signal: AbortSignal) => {
  if (signal.aborted) throw new Error('aborted')
}
const integer = (v: unknown, min: number, max: number): v is number =>
  typeof v === 'number' && Number.isSafeInteger(v) && v >= min && v <= max
const digest = (v: unknown): v is string => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
const key = (v: unknown): v is string =>
  typeof v === 'string' &&
  v.length <= 128 &&
  /^(snapshot|page-[0-9]+|image-[0-9]+|receipt-[0-9]+)$/.test(v)
const id = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
interface Metadata {
  documentId: string
  changeId: string
  key: string
  sha256: string
  sizeBytes: number
  status: 'uploading' | 'ready'
}
const metadataKeys = ['documentId', 'changeId', 'key', 'sha256', 'sizeBytes', 'status']
function validMetadata(v: unknown): v is Metadata {
  if (!v || typeof v !== 'object' || Array.isArray(v)) return false
  const m = v as Record<string, unknown>
  return (
    Object.keys(m).length === metadataKeys.length &&
    metadataKeys.every((k) => Object.hasOwn(m, k)) &&
    typeof m.documentId === 'string' &&
    !!m.documentId &&
    m.documentId.length <= 2048 &&
    id(m.changeId) &&
    key(m.key) &&
    digest(m.sha256) &&
    integer(m.sizeBytes, 1, MAX_BLOB) &&
    (m.status === 'ready' || m.status === 'uploading')
  )
}
type Parents = { path: string; stat: Stats }[]
const sameFile = (a: Stats, b: Stats) => a.dev === b.dev && a.ino === b.ino
function guard(parents: Parents) {
  for (const parent of parents) {
    const current = lstatSync(parent.path)
    if (!current.isDirectory() || current.isSymbolicLink() || !sameFile(current, parent.stat))
      invalid()
  }
}
function leaf(path: string, parents: Parents) {
  guard(parents)
  const stat = lstatSync(path)
  if (!stat.isFile() || stat.isSymbolicLink()) invalid()
  return stat
}
function held(path: string, fd: number, expected: Stats, parents: Parents) {
  const stat = fstatSync(fd)
  if (!stat.isFile() || !sameFile(stat, expected) || !sameFile(leaf(path, parents), expected))
    invalid()
}
async function directory(path: string, parents: Parents): Promise<Parents> {
  guard(parents)
  try {
    await mkdir(path, { mode: 0o700 })
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code !== 'EEXIST') throw e
  }
  guard(parents)
  const stat = lstatSync(path)
  if (!stat.isDirectory() || stat.isSymbolicLink()) invalid()
  return [...parents, { path, stat }]
}
async function syncDirectory(parents: Parents) {
  guard(parents)
  if (process.platform === 'win32') return
  const parent = parents[parents.length - 1]!
  const f = await open(parent.path, constants.O_RDONLY | constants.O_NOFOLLOW)
  try {
    guard(parents)
    if (!sameFile(fstatSync(f.fd), parent.stat)) invalid()
    await f.sync()
    guard(parents)
  } finally {
    await f.close()
    guard(parents)
  }
}
async function bytes(path: string, max: number, parents: Parents) {
  const expected = leaf(path, parents)
  const f = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW)
  try {
    held(path, f.fd, expected, parents)
    if (fstatSync(f.fd).size > max) invalid()
    const raw = await f.readFile()
    held(path, f.fd, expected, parents)
    if (raw.length > max) invalid()
    return raw
  } finally {
    await f.close()
    if (!sameFile(leaf(path, parents), expected)) invalid()
  }
}
async function atomic(path: string, value: string | Buffer, parents: Parents) {
  const temporary = `${path}.${randomUUID()}.tmp`
  guard(parents)
  try {
    const f = await open(
      temporary,
      constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW,
      0o600,
    )
    try {
      const expected = leaf(temporary, parents)
      // Final checks are synchronous; the immediately invoked write uses the verified held descriptor.
      held(temporary, f.fd, expected, parents)
      await f.writeFile(value)
      await f.sync()
      held(temporary, f.fd, expected, parents)
    } finally {
      await f.close()
    }
    guard(parents)
    // Keep ancestor validation and publication in the same turn, without an asynchronous path lookup gap.
    renameSync(temporary, path)
    guard(parents)
    await syncDirectory(parents)
  } finally {
    // A substituted parent must never turn cleanup into a deletion in a foreign directory.
    try {
      guard(parents)
      rmSync(temporary, { force: true })
    } catch {
      /* retain an unreachable temporary on failed identity proof */
    }
  }
}
function metadata(path: string, parents: Parents): { value: Metadata; parents: Parents } {
  guard(parents)
  const stat = lstatSync(path)
  if (!stat.isDirectory() || stat.isSymbolicLink()) invalid()
  const scoped = [...parents, { path, stat }]
  const file = join(path, 'metadata.json')
  const expected = leaf(file, scoped)
  const fd = openSync(file, constants.O_RDONLY | constants.O_NOFOLLOW)
  let raw: Buffer
  try {
    held(file, fd, expected, scoped)
    if (fstatSync(fd).size > 4096) invalid()
    raw = readFileSync(fd)
    held(file, fd, expected, scoped)
    if (raw.length > 4096) invalid()
  } finally {
    closeSync(fd)
    if (!sameFile(leaf(file, scoped), expected)) invalid()
  }
  const value: unknown = JSON.parse(raw.toString())
  if (!validMetadata(value)) invalid()
  return { value, parents: scoped }
}
async function sweepReleased(root: string, rootParents: Parents, signal: AbortSignal) {
  const previous = sweepLocks.get(root) ?? Promise.resolve()
  let release!: () => void
  const tail = new Promise<void>((resolve) => {
    release = resolve
  })
  sweepLocks.set(root, tail)
  await previous
  try {
    const entries = await readdir(root)
    guard(rootParents)
    for (const entry of entries) {
      if (
        !/^\.released-[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(
          entry,
        )
      )
        continue
      check(signal)
      const path = join(root, entry),
        stat = lstatSync(path)
      if (!stat.isDirectory() || stat.isSymbolicLink()) invalid()
      const parents = [...rootParents, { path, stat }],
        members = await readdir(path)
      guard(parents)
      if (members.some((member) => !['blob', 'metadata.json'].includes(member))) invalid()
      if (members.includes('metadata.json')) metadata(path, rootParents)
      else if (members.length) invalid()
      if (members.includes('blob')) {
        const blob = join(path, 'blob')
        if (leaf(blob, parents).size > MAX_BLOB) invalid()
        rmSync(blob)
      }
      if (members.includes('metadata.json')) {
        const file = join(path, 'metadata.json')
        leaf(file, parents)
        rmSync(file)
      }
      guard(rootParents)
      rmdirSync(path)
    }
    await syncDirectory(rootParents)
  } finally {
    release()
    if (sweepLocks.get(root) === tail) sweepLocks.delete(root)
  }
}
export function createPresentationMasterBackupService(options: {
  userDataPath: string
  storageDirectory?: 'presentation-master-backups' | 'presentation-package-backups'
}) {
  const userData = resolve(options.userDataPath)
  const storageDirectory = options.storageDirectory ?? 'presentation-master-backups'
  if (!['presentation-master-backups', 'presentation-package-backups'].includes(storageDirectory))
    invalid()
  const root = join(userData, storageDirectory)
  return async (input: Record<string, unknown>, signal: AbortSignal): Promise<unknown> => {
    check(signal)
    let body: Record<string, unknown>
    try {
      body = structuredClone(input)
    } catch {
      return invalid()
    }
    const fields: Record<string, string[]> = {
      master_backup_begin: ['key', 'sha256', 'sizeBytes'],
      master_backup_chunk: ['key', 'offset', 'base64'],
      master_backup_finish: ['key'],
      master_backup_status: ['key'],
      master_backup_read: ['key', 'offset', 'length'],
      master_backup_list: [],
      master_backup_release: ['key', 'sha256', 'sizeBytes'],
    }
    const operation = body.operation
    if (typeof operation !== 'string' || !Object.hasOwn(fields, operation)) invalid()
    const expected = ['operation', 'documentId', 'changeId', ...fields[operation]!]
    if (
      Object.keys(body).length !== expected.length ||
      expected.some((k) => !Object.hasOwn(body, k)) ||
      typeof body.documentId !== 'string' ||
      !body.documentId ||
      body.documentId.length > 2048 ||
      !id(body.changeId) ||
      Buffer.byteLength(JSON.stringify(body)) > 256 * 1024 ||
      (operation !== 'master_backup_list' && !key(body.key))
    )
      invalid()
    if (
      operation === 'master_backup_begin' &&
      (!digest(body.sha256) || !integer(body.sizeBytes, 1, MAX_BLOB))
    )
      invalid()
    if (
      operation === 'master_backup_release' &&
      (!digest(body.sha256) || !integer(body.sizeBytes, 1, MAX_BLOB))
    )
      invalid()
    const documentId = body.documentId,
      changeId = body.changeId,
      blobKey = body.key as string
    const document = join(root, hash(documentId)),
      blobHash = hash(JSON.stringify([changeId, blobKey])),
      dir = join(document, blobHash)
    const previous = locks.get(document) ?? Promise.resolve()
    let release!: () => void
    const tail = new Promise<void>((resolve) => {
      release = resolve
    })
    locks.set(document, tail)
    await previous
    try {
      check(signal)
      const userStat = lstatSync(userData)
      if (!userStat.isDirectory() || userStat.isSymbolicLink()) invalid()
      const rootParents = await directory(root, [{ path: userData, stat: userStat }])
      await sweepReleased(root, rootParents, signal)
      const documentParents = await directory(document, rootParents)
      const entries = await readdir(document)
      guard(documentParents)
      if (entries.length > MAX_BLOBS || entries.some((e) => !digest(e))) invalid()
      // ponytail: quota allocation scans at most 4096 reservations; nonallocating operations verify only their own blob.
      const all: Metadata[] = []
      const scopes = new Map<string, Parents>()
      let reserved = 0
      const quotaScan = operation === 'master_backup_begin' || operation === 'master_backup_list'
      const inspectedEntries = quotaScan ? entries : entries.filter((entry) => entry === blobHash)
      for (const entry of inspectedEntries) {
        check(signal)
        const result = metadata(join(document, entry), documentParents)
        const m = result.value
        scopes.set(entry, result.parents)
        if (m.documentId !== documentId || hash(JSON.stringify([m.changeId, m.key])) !== entry)
          invalid()
        reserved += m.sizeBytes
        all.push(m)
      }
      if (reserved > MAX_DOCUMENT) invalid()
      if (operation === 'master_backup_list') {
        const backups = []
        for (const m of all.filter((m) => m.changeId === changeId)) {
          const raw = await bytes(
            join(document, hash(JSON.stringify([m.changeId, m.key])), 'blob'),
            MAX_BLOB,
            scopes.get(hash(JSON.stringify([m.changeId, m.key])))!,
          )
          check(signal)
          if (
            raw.length > m.sizeBytes ||
            (m.status === 'ready' && (raw.length !== m.sizeBytes || hash(raw) !== m.sha256))
          )
            invalid()
          backups.push({ ...m, receivedBytes: raw.length })
        }
        return { documentId, changeId, backups }
      }
      const existing = all.find((m) => m.changeId === changeId && m.key === blobKey)
      if (!existing) {
        if (operation !== 'master_backup_begin') invalid()
        if (entries.length >= MAX_BLOBS || reserved + (body.sizeBytes as number) > MAX_DOCUMENT)
          throw new Error('presentation_master_backup_capacity')
        const m: Metadata = {
          documentId,
          changeId,
          key: blobKey,
          sha256: body.sha256 as string,
          sizeBytes: body.sizeBytes as number,
          status: 'uploading',
        }
        if (Buffer.byteLength(JSON.stringify(m)) > 4096) invalid()
        const staging = join(root, `.tmp-${randomUUID()}`)
        const stagingParents = await directory(staging, rootParents)
        try {
          await atomic(join(staging, 'blob'), Buffer.alloc(0), stagingParents)
          await atomic(join(staging, 'metadata.json'), JSON.stringify(m), stagingParents)
          check(signal)
          guard(stagingParents)
          guard(documentParents)
          renameSync(staging, dir)
          guard(documentParents)
          await syncDirectory(documentParents)
        } finally {
          try {
            guard(stagingParents)
            rmSync(staging, { recursive: true, force: true })
          } catch {
            /* no cleanup through substituted ancestors */
          }
        }
        return { ...m, receivedBytes: 0 }
      }
      let m = existing
      const blobParents = scopes.get(blobHash)!
      const path = join(dir, 'blob'),
        raw = await bytes(path, MAX_BLOB, blobParents),
        received = raw.length
      if (
        received > m.sizeBytes ||
        (m.status === 'ready' && (received !== m.sizeBytes || hash(raw) !== m.sha256))
      )
        invalid()
      check(signal)
      if (operation === 'master_backup_release') {
        if (body.sha256 !== m.sha256 || body.sizeBytes !== m.sizeBytes) invalid()
        const members = await readdir(dir)
        guard(blobParents)
        if (members.length !== 2 || !members.includes('blob') || !members.includes('metadata.json'))
          invalid()
        leaf(path, blobParents)
        leaf(join(dir, 'metadata.json'), blobParents)
        check(signal)
        const releasedPath = join(root, `.released-${randomUUID()}`)
        guard(blobParents)
        guard(rootParents)
        renameSync(dir, releasedPath)
        await syncDirectory(documentParents)
        try {
          await sweepReleased(root, rootParents, signal)
        } catch {
          // The source reservation is already gone. A later storage sweep can remove
          // this isolated release directory without touching an active backup.
        }
        return { documentId, changeId, key: blobKey, released: true }
      }
      if (operation === 'master_backup_begin') {
        if (body.sha256 !== m.sha256 || body.sizeBytes !== m.sizeBytes) invalid()
        return { ...m, receivedBytes: received }
      }
      if (operation === 'master_backup_chunk') {
        if (
          !integer(body.offset, 0, m.sizeBytes) ||
          typeof body.base64 !== 'string' ||
          body.base64.length > Math.ceil(CHUNK / 3) * 4
        )
          invalid()
        const chunk = Buffer.from(body.base64, 'base64'),
          offset = body.offset
        if (
          !chunk.length ||
          chunk.length > CHUNK ||
          chunk.toString('base64') !== body.base64 ||
          offset + chunk.length > m.sizeBytes ||
          offset > received
        )
          invalid()
        const overlap = Math.min(chunk.length, received - offset)
        if (!raw.subarray(offset, offset + overlap).equals(chunk.subarray(0, overlap))) invalid()
        if (overlap < chunk.length) {
          const expected = leaf(path, blobParents)
          const f = await open(path, constants.O_RDWR | constants.O_APPEND | constants.O_NOFOLLOW)
          try {
            held(path, f.fd, expected, blobParents)
            const s = fstatSync(f.fd)
            if (!s.isFile() || s.size !== received) invalid()
            check(signal)
            await f.writeFile(chunk.subarray(overlap))
            await f.sync()
            held(path, f.fd, expected, blobParents)
          } finally {
            await f.close()
            if (!sameFile(leaf(path, blobParents), expected)) invalid()
          }
        }
        return { ...m, receivedBytes: Math.max(received, offset + chunk.length) }
      }
      if (operation === 'master_backup_finish') {
        if (received !== m.sizeBytes || hash(raw) !== m.sha256) invalid()
        if (m.status !== 'ready') {
          m = { ...m, status: 'ready' }
          check(signal)
          await atomic(join(dir, 'metadata.json'), JSON.stringify(m), blobParents)
        }
      }
      guard(blobParents)
      if (operation === 'master_backup_read') {
        if (
          m.status !== 'ready' ||
          !integer(body.offset, 0, m.sizeBytes) ||
          !integer(body.length, 1, CHUNK) ||
          body.offset + body.length > m.sizeBytes
        )
          invalid()
        return {
          documentId,
          changeId,
          key: blobKey,
          sha256: m.sha256,
          sizeBytes: m.sizeBytes,
          offset: body.offset,
          base64: raw.subarray(body.offset, body.offset + body.length).toString('base64'),
        }
      }
      check(signal)
      return { ...m, receivedBytes: received }
    } catch (error) {
      if (
        error instanceof Error &&
        [
          'aborted',
          'presentation_master_backup_invalid',
          'presentation_master_backup_capacity',
        ].includes(error.message)
      )
        throw error
      return invalid()
    } finally {
      release()
      if (locks.get(document) === tail) locks.delete(document)
    }
  }
}
