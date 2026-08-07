import { createHash, randomBytes } from 'node:crypto'
import { constants } from 'node:fs'
import {
  lstat,
  mkdir,
  mkdtemp,
  open,
  readFile,
  rename,
  rm,
  unlink,
  writeFile,
} from 'node:fs/promises'
import { join } from 'node:path'
import { AtomicWriteCommittedError, atomicWriteFile } from './atomic-write.js'
import { ProjectPathPolicy } from './path-policy.js'
import type { LatexProject } from './project.js'

const INDEX_SCHEMA_VERSION = 1
const MANIFEST_SCHEMA_VERSION = 1

export interface SnapshotStoreOptions {
  maxBytes?: number
  maxSnapshots?: number
  now?: () => number
  /** @internal deterministic storage failure hooks for tests. */
  storageHooks?: {
    writeIndex?: (path: string, data: Uint8Array) => Promise<void>
    removeSnapshotDirectory?: (path: string) => Promise<void>
  }
}

export interface SnapshotSummary {
  id: string
  projectId: string
  createdAt: number
  byteLength: number
  paths: string[]
}

type SnapshotEntry =
  | { kind: 'text'; path: string; sha256: string; byteLength: number; file: string }
  | { kind: 'absent'; path: string }

interface SnapshotManifest {
  schemaVersion: typeof MANIFEST_SCHEMA_VERSION
  id: string
  projectId: string
  createdAt: number
  entries: SnapshotEntry[]
}

interface SnapshotIndex {
  schemaVersion: typeof INDEX_SCHEMA_VERSION
  snapshots: SnapshotSummary[]
  pendingDeletes: SnapshotSummary[]
  currentRollback: Record<string, string>
}

export interface SnapshotRestoreOptions {
  /** Required before deleting a file represented by an absent marker. */
  expectedAbsentHashes?: ReadonlyMap<string, string>
  /** Require each text entry to still match this hash before restoration. */
  expectedCurrentHashes?: ReadonlyMap<string, string>
}

export class SnapshotStore {
  private readonly root: string
  private readonly maxBytes: number
  private readonly maxSnapshots: number
  private readonly now: () => number
  private readonly writeIndex: (path: string, data: Uint8Array) => Promise<void>
  private readonly removeSnapshotDirectory: (path: string) => Promise<void>

  constructor(cacheRoot: string, options: SnapshotStoreOptions = {}) {
    this.root = join(cacheRoot, 'snapshots')
    this.maxBytes = options.maxBytes ?? 64 * 1024 * 1024
    this.maxSnapshots = options.maxSnapshots ?? 20
    this.now = options.now ?? Date.now
    this.writeIndex = options.storageHooks?.writeIndex ?? atomicWriteFile
    this.removeSnapshotDirectory =
      options.storageHooks?.removeSnapshotDirectory ??
      ((path) => rm(path, { recursive: true, force: true }))
    if (!Number.isSafeInteger(this.maxBytes) || this.maxBytes <= 0) {
      throw new Error('Snapshot maxBytes must be a positive safe integer')
    }
    if (!Number.isSafeInteger(this.maxSnapshots) || this.maxSnapshots <= 0) {
      throw new Error('Snapshot maxSnapshots must be a positive safe integer')
    }
  }

  async create(
    projectId: string,
    project: LatexProject,
    relativePaths: readonly string[],
  ): Promise<SnapshotSummary> {
    validateProjectId(projectId)
    await mkdir(this.root, { recursive: true })
    const policy = await ProjectPathPolicy.open(project.rootPath)
    const normalizedPaths = relativePaths.map((path) => policy.normalize(path))
    if (new Set(normalizedPaths).size !== normalizedPaths.length) {
      throw new Error('Snapshot contains duplicate normalized paths')
    }
    const id = randomBytes(16).toString('hex')
    const temporary = await mkdtemp(join(this.root, '.snapshot-'))
    const entries: SnapshotEntry[] = []
    let byteLength = 0
    try {
      await mkdir(join(temporary, 'files'))
      for (const [index, path] of normalizedPaths.entries()) {
        if (!(await textExists(policy, path))) {
          entries.push({ kind: 'absent', path })
          continue
        }
        const text = await project.readText(path)
        const bytes = Buffer.from(text, 'utf8')
        const file = `files/${index}.txt`
        await writeFile(join(temporary, file), bytes, { flag: 'wx', mode: 0o600 })
        entries.push({
          kind: 'text',
          path,
          sha256: digest(bytes),
          byteLength: bytes.length,
          file,
        })
        byteLength += bytes.length
      }
      const createdAt = this.now()
      const manifest: SnapshotManifest = {
        schemaVersion: MANIFEST_SCHEMA_VERSION,
        id,
        projectId,
        createdAt,
        entries,
      }
      const manifestBytes = Buffer.from(`${JSON.stringify(manifest, null, 2)}\n`)
      await writeFile(join(temporary, 'manifest.json'), manifestBytes, { flag: 'wx', mode: 0o600 })
      byteLength += manifestBytes.length
      const summary: SnapshotSummary = {
        id,
        projectId,
        createdAt,
        byteLength,
        paths: normalizedPaths,
      }
      await rename(temporary, this.snapshotPath(id))
      try {
        await this.addAndPrune(summary)
      } catch (error) {
        let indexed: boolean
        try {
          indexed = (await this.readIndex()).snapshots.some((item) => item.id === id)
        } catch (readbackError) {
          throw new AggregateError(
            [error, readbackError],
            'Snapshot index update failed and its commit state could not be determined',
            { cause: readbackError },
          )
        }
        if (indexed) return summary
        await rm(this.snapshotPath(id), { recursive: true, force: true })
        throw error
      }
      return summary
    } catch (error) {
      await rm(temporary, { recursive: true, force: true })
      throw error
    }
  }

  async list(projectId: string): Promise<SnapshotSummary[]> {
    validateProjectId(projectId)
    const index = await this.readIndex()
    return index.snapshots.filter((snapshot) => snapshot.projectId === projectId)
  }

  async setCurrentRollback(projectId: string, snapshotId: string | null): Promise<void> {
    validateProjectId(projectId)
    await this.mutateIndex((index) => {
      if (snapshotId === null) {
        delete index.currentRollback[projectId]
        return
      }
      validateSnapshotId(snapshotId)
      if (!index.snapshots.some((item) => item.id === snapshotId && item.projectId === projectId)) {
        throw new Error(`Snapshot not found for project: ${snapshotId}`)
      }
      index.currentRollback[projectId] = snapshotId
    })
  }

  async getCurrentRollback(projectId: string): Promise<string | undefined> {
    validateProjectId(projectId)
    return (await this.readIndex()).currentRollback[projectId]
  }

  async restore(
    projectId: string,
    snapshotId: string,
    project: LatexProject,
    options: SnapshotRestoreOptions = {},
  ): Promise<void> {
    validateProjectId(projectId)
    validateSnapshotId(snapshotId)
    const manifest = await this.readManifest(snapshotId)
    if (manifest.projectId !== projectId) throw new Error('Snapshot belongs to another project')
    const policy = await ProjectPathPolicy.open(project.rootPath)
    const actions: (
      | { kind: 'text'; path: string; text: string; expectedHash?: string }
      | { kind: 'absent'; path: string; expectedHash: string }
    )[] = []
    for (const entry of manifest.entries) {
      if (entry.kind === 'text') {
        const bytes = await readFile(join(this.snapshotPath(snapshotId), entry.file))
        if (bytes.length !== entry.byteLength || digest(bytes) !== entry.sha256) {
          throw new Error(`Snapshot payload failed integrity check: ${entry.path}`)
        }
        const text = new TextDecoder('utf-8', { fatal: true }).decode(bytes)
        const expectedHash = options.expectedCurrentHashes?.get(entry.path)
        if (options.expectedCurrentHashes) {
          if (!expectedHash || !(await textExists(policy, entry.path))) {
            throw new Error(`Project changed before snapshot restore: ${entry.path}`)
          }
          const currentHash = digest(Buffer.from(await project.readText(entry.path), 'utf8'))
          if (currentHash === entry.sha256) continue
          if (currentHash !== expectedHash) {
            throw new Error(`Project changed before snapshot restore: ${entry.path}`)
          }
        }
        actions.push({ kind: 'text', path: entry.path, text, expectedHash })
      } else {
        const expectedHash = options.expectedAbsentHashes?.get(entry.path)
        const exists = await textExists(policy, entry.path)
        if (exists) {
          if (!expectedHash) {
            throw new Error(
              `Snapshot absent marker requires an expected current hash: ${entry.path}`,
            )
          }
          const current = await project.readText(entry.path)
          if (digest(Buffer.from(current, 'utf8')) !== expectedHash) {
            throw new Error(`Project changed before restoring absent file: ${entry.path}`)
          }
        }
        if (expectedHash) actions.push({ kind: 'absent', path: entry.path, expectedHash })
      }
    }

    for (const action of actions) {
      if (action.kind === 'text') {
        await project.saveText(action.path, action.text, {
          expectedSha256: action.expectedHash,
        })
      } else {
        await project.deleteText(action.path, {
          expectedSha256: action.expectedHash,
          transactionId: snapshotId,
        })
      }
    }
  }

  async getFileHashes(
    projectId: string,
    snapshotId: string,
  ): Promise<ReadonlyMap<string, string | null>> {
    validateProjectId(projectId)
    validateSnapshotId(snapshotId)
    const manifest = await this.readManifest(snapshotId)
    if (manifest.projectId !== projectId) throw new Error('Snapshot belongs to another project')
    return new Map(
      manifest.entries.map((entry) => [entry.path, entry.kind === 'text' ? entry.sha256 : null]),
    )
  }

  async verifyRestored(
    projectId: string,
    snapshotId: string,
    project: LatexProject,
  ): Promise<void> {
    const manifest = await this.readManifest(snapshotId)
    if (manifest.projectId !== projectId) throw new Error('Snapshot belongs to another project')
    const policy = await ProjectPathPolicy.open(project.rootPath)
    for (const entry of manifest.entries) {
      const exists = await textExists(policy, entry.path)
      if (entry.kind === 'absent') {
        if (exists) throw new Error(`Project changed after undo: ${entry.path}`)
      } else {
        if (
          !exists ||
          digest(Buffer.from(await project.readText(entry.path), 'utf8')) !== entry.sha256
        ) {
          throw new Error(`Project changed after undo: ${entry.path}`)
        }
      }
    }
  }

  private async addAndPrune(summary: SnapshotSummary): Promise<void> {
    try {
      await this.mutateIndex((index) => {
        const pendingCount = index.pendingDeletes.length
        const pendingBytes = index.pendingDeletes.reduce((sum, item) => sum + item.byteLength, 0)
        index.snapshots.push(summary)
        const isOver = () =>
          index.snapshots.length + pendingCount > this.maxSnapshots ||
          index.snapshots.reduce((sum, item) => sum + item.byteLength, pendingBytes) > this.maxBytes
        while (isOver()) {
          const candidate = index.snapshots
            .filter(
              (item) => item.id !== summary.id && index.currentRollback[item.projectId] !== item.id,
            )
            .sort((left, right) => left.createdAt - right.createdAt)[0]
          if (!candidate)
            throw new Error('Snapshot cleanup is pending and quota cannot accept another snapshot')
          index.snapshots = index.snapshots.filter((item) => item.id !== candidate.id)
          index.pendingDeletes.push(candidate)
        }
      })
    } catch (error) {
      const committed = (await this.readIndex()).snapshots.some((item) => item.id === summary.id)
      if (!committed) throw error
    }
    await this.flushPendingDeletes()
  }

  private async readManifest(snapshotId: string): Promise<SnapshotManifest> {
    let value: unknown
    try {
      value = JSON.parse(
        await readFile(join(this.snapshotPath(snapshotId), 'manifest.json'), 'utf8'),
      )
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
        throw new Error(`Snapshot not found: ${snapshotId}`, { cause: error })
      }
      throw error
    }
    if (!isManifest(value) || value.id !== snapshotId) throw new Error('Invalid snapshot manifest')
    return value
  }

  private snapshotPath(snapshotId: string): string {
    validateSnapshotId(snapshotId)
    return join(this.root, snapshotId)
  }

  private async readIndex(): Promise<SnapshotIndex> {
    try {
      const value: unknown = JSON.parse(await readFile(join(this.root, 'index.json'), 'utf8'))
      if (!isIndex(value)) throw new Error('Invalid snapshot index')
      return { ...value, pendingDeletes: value.pendingDeletes ?? [] }
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
        return {
          schemaVersion: INDEX_SCHEMA_VERSION,
          snapshots: [],
          pendingDeletes: [],
          currentRollback: {},
        }
      }
      throw error
    }
  }

  private async mutateIndex(change: (index: SnapshotIndex) => void): Promise<void> {
    await mkdir(this.root, { recursive: true })
    const lockPath = join(this.root, 'index.lock')
    const lock = await acquireIndexLock(lockPath)
    try {
      const index = await this.readIndex()
      const cleaned = await this.retryPendingDeletes(index)
      try {
        change(index)
      } catch (error) {
        if (cleaned) await this.persistIndex(index)
        throw error
      }
      await this.persistIndex(index)
    } finally {
      await releaseIndexLock(lockPath, lock)
    }
  }

  private async flushPendingDeletes(): Promise<void> {
    if ((await this.readIndex()).pendingDeletes.length === 0) return
    await this.mutateIndex(() => undefined)
  }

  private async retryPendingDeletes(index: SnapshotIndex): Promise<boolean> {
    if (index.pendingDeletes.length === 0) return false
    const remaining: SnapshotSummary[] = []
    for (const snapshot of index.pendingDeletes) {
      try {
        await this.removeSnapshotDirectory(this.snapshotPath(snapshot.id))
      } catch {
        remaining.push(snapshot)
      }
    }
    const changed = remaining.length !== index.pendingDeletes.length
    index.pendingDeletes = remaining
    return changed
  }

  private async persistIndex(index: SnapshotIndex): Promise<void> {
    const data = Buffer.from(`${JSON.stringify(index, null, 2)}\n`)
    try {
      await this.writeIndex(join(this.root, 'index.json'), data)
    } catch (error) {
      if (
        error instanceof AtomicWriteCommittedError &&
        JSON.stringify(await this.readIndex()) === JSON.stringify(index)
      ) {
        return
      }
      throw error
    }
  }
}

interface HeldIndexLock {
  handle: import('node:fs/promises').FileHandle
  token: string
  dev: number
  ino: number
}

interface IndexLockRecord {
  schemaVersion: 1
  pid: number
  token: string
  createdAt: number
}

const EMPTY_LOCK_STALE_MS = 30_000

async function acquireIndexLock(path: string): Promise<HeldIndexLock> {
  const token = randomBytes(16).toString('hex')
  const noFollow = 'O_NOFOLLOW' in constants ? constants.O_NOFOLLOW : 0
  for (let attempt = 0; attempt < 2; attempt += 1) {
    let handle: import('node:fs/promises').FileHandle | undefined
    let ownedIdentity: Pick<HeldIndexLock, 'dev' | 'ino'> | undefined
    try {
      handle = await open(
        path,
        constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | noFollow,
        0o600,
      )
      const info = await handle.stat()
      ownedIdentity = { dev: info.dev, ino: info.ino }
      const record: IndexLockRecord = {
        schemaVersion: 1,
        pid: process.pid,
        token,
        createdAt: Date.now(),
      }
      await handle.writeFile(`${JSON.stringify(record)}\n`)
      await handle.sync()
      return { handle, token, ...ownedIdentity }
    } catch (error) {
      await handle?.close().catch(() => undefined)
      if (ownedIdentity) {
        try {
          await unlinkOwnedIndexLock(path, ownedIdentity)
        } catch (cleanupError) {
          throw new AggregateError(
            [error, cleanupError],
            'Snapshot index lock acquisition and cleanup both failed',
            { cause: cleanupError },
          )
        }
      }
      if (
        attempt === 0 &&
        (error as NodeJS.ErrnoException).code === 'EEXIST' &&
        (await removeStaleIndexLock(path))
      ) {
        continue
      }
      throw error
    }
  }
  throw new Error('Unable to acquire snapshot index lock')
}

async function unlinkOwnedIndexLock(
  path: string,
  identity: Pick<HeldIndexLock, 'dev' | 'ino'>,
): Promise<void> {
  try {
    const info = await lstat(path)
    if (info.dev === identity.dev && info.ino === identity.ino) await unlink(path)
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
  }
}

async function removeStaleIndexLock(path: string): Promise<boolean> {
  let before
  let contents: string
  try {
    before = await lstat(path)
    if (!before.isFile() || before.isSymbolicLink()) return false
    contents = await readFile(path, 'utf8')
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return true
    throw error
  }

  const record = parseIndexLock(contents)
  const stale = record
    ? !isProcessAlive(record.pid)
    : Date.now() - before.mtimeMs >= EMPTY_LOCK_STALE_MS
  if (!stale) return false

  try {
    const after = await lstat(path)
    if (after.dev !== before.dev || after.ino !== before.ino) return false
    if ((await readFile(path, 'utf8')) !== contents) return false
    await unlink(path)
    return true
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return true
    throw error
  }
}

async function releaseIndexLock(path: string, lock: HeldIndexLock): Promise<void> {
  await lock.handle.close().catch(() => undefined)
  try {
    const info = await lstat(path)
    if (info.dev !== lock.dev || info.ino !== lock.ino) return
    const record = parseIndexLock(await readFile(path, 'utf8'))
    if (record?.token === lock.token && record.pid === process.pid) await unlink(path)
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
  }
}

function parseIndexLock(contents: string): IndexLockRecord | undefined {
  try {
    const value: unknown = JSON.parse(contents)
    if (!value || typeof value !== 'object') return undefined
    const record = value as Partial<IndexLockRecord>
    if (
      record.schemaVersion !== 1 ||
      !Number.isSafeInteger(record.pid) ||
      (record.pid ?? 0) <= 0 ||
      typeof record.token !== 'string' ||
      !/^[a-f0-9]{32}$/.test(record.token) ||
      !Number.isSafeInteger(record.createdAt)
    ) {
      return undefined
    }
    return record as IndexLockRecord
  } catch {
    return undefined
  }
}

function isProcessAlive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return (error as NodeJS.ErrnoException).code !== 'ESRCH'
  }
}
async function textExists(policy: ProjectPathPolicy, path: string): Promise<boolean> {
  try {
    await policy.resolveExisting(path, 'file')
    return true
  } catch (error) {
    if (error instanceof Error && /does not exist/.test(error.message)) return false
    throw error
  }
}

function digest(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex')
}

function validateProjectId(projectId: string): void {
  if (!isSafeProjectId(projectId)) throw new Error('Invalid project ID')
}

function validateSnapshotId(snapshotId: string): void {
  if (!/^[a-f0-9]{32}$/.test(snapshotId)) throw new Error('Invalid snapshot ID')
}

function isManifest(value: unknown): value is SnapshotManifest {
  if (!value || typeof value !== 'object') return false
  const item = value as Partial<SnapshotManifest>
  if (
    item.schemaVersion !== MANIFEST_SCHEMA_VERSION ||
    !isSafeSnapshotId(item.id) ||
    !isSafeProjectId(item.projectId) ||
    !Number.isSafeInteger(item.createdAt) ||
    !Array.isArray(item.entries)
  ) {
    return false
  }

  const paths = new Set<string>()
  const files = new Set<string>()
  for (const [index, valueEntry] of item.entries.entries()) {
    if (!valueEntry || typeof valueEntry !== 'object') return false
    const entry = valueEntry as Partial<SnapshotEntry>
    if (!isSafeSnapshotPath(entry.path) || paths.has(entry.path)) return false
    paths.add(entry.path)
    if (entry.kind === 'absent') continue
    if (
      entry.kind !== 'text' ||
      typeof entry.sha256 !== 'string' ||
      !/^[a-f0-9]{64}$/.test(entry.sha256) ||
      !Number.isSafeInteger(entry.byteLength) ||
      (entry.byteLength ?? -1) < 0 ||
      entry.file !== `files/${index}.txt` ||
      files.has(entry.file)
    ) {
      return false
    }
    files.add(entry.file)
  }
  return true
}

function isIndex(value: unknown): value is SnapshotIndex {
  if (!value || typeof value !== 'object') return false
  const item = value as Partial<SnapshotIndex>
  if (
    item.schemaVersion !== INDEX_SCHEMA_VERSION ||
    !Array.isArray(item.snapshots) ||
    (item.pendingDeletes !== undefined && !Array.isArray(item.pendingDeletes)) ||
    !item.currentRollback ||
    typeof item.currentRollback !== 'object' ||
    Array.isArray(item.currentRollback)
  ) {
    return false
  }
  const ids = new Set<string>()
  if (
    !item.snapshots.every((summary) => {
      if (!isSnapshotSummary(summary) || ids.has(summary.id)) return false
      ids.add(summary.id)
      return true
    })
  ) {
    return false
  }
  const pendingDeletes = item.pendingDeletes ?? []
  if (
    !pendingDeletes.every((summary) => {
      if (!isSnapshotSummary(summary) || ids.has(summary.id)) return false
      ids.add(summary.id)
      return true
    })
  ) {
    return false
  }
  return Object.entries(item.currentRollback).every(
    ([projectId, snapshotId]) =>
      isSafeProjectId(projectId) &&
      isSafeSnapshotId(snapshotId) &&
      item.snapshots?.some(
        (summary) => summary.id === snapshotId && summary.projectId === projectId,
      ),
  )
}

function isSnapshotSummary(value: unknown): value is SnapshotSummary {
  if (!value || typeof value !== 'object') return false
  const summary = value as Partial<SnapshotSummary>
  return (
    isSafeSnapshotId(summary.id) &&
    isSafeProjectId(summary.projectId) &&
    Number.isSafeInteger(summary.createdAt) &&
    Number.isSafeInteger(summary.byteLength) &&
    (summary.byteLength ?? -1) >= 0 &&
    Array.isArray(summary.paths) &&
    summary.paths.every(isSafeSnapshotPath) &&
    new Set(summary.paths).size === summary.paths.length
  )
}

function isSafeProjectId(value: unknown): value is string {
  return (
    typeof value === 'string' &&
    /^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/.test(value) &&
    !['__proto__', 'constructor', 'prototype'].includes(value)
  )
}

function isSafeSnapshotId(value: unknown): value is string {
  return typeof value === 'string' && /^[a-f0-9]{32}$/.test(value)
}

function isSafeSnapshotPath(value: unknown): value is string {
  if (
    typeof value !== 'string' ||
    !value ||
    value.includes('\0') ||
    value.includes('\\') ||
    value.startsWith('/') ||
    /^[a-zA-Z]:[\\/]/.test(value)
  ) {
    return false
  }
  return value.split('/').every((segment) => segment !== '' && segment !== '.' && segment !== '..')
}
