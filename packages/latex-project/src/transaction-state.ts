import { randomBytes } from 'node:crypto'
import { constants } from 'node:fs'
import { lstat, mkdir, open, readFile, readdir, unlink } from 'node:fs/promises'
import { extname, join } from 'node:path'
import { AtomicWriteCommittedError, atomicWriteFile } from './atomic-write.js'

const JOURNAL_SCHEMA_VERSION = 1
const PROJECT_STATE_SCHEMA_VERSION = 1
const LOCK_SCHEMA_VERSION = 1
const DEFAULT_EMPTY_LOCK_STALE_MS = 30_000
const TEXT_EXTENSIONS = new Set([
  '.bib',
  '.bst',
  '.cls',
  '.csv',
  '.ltx',
  '.md',
  '.sty',
  '.tex',
  '.txt',
])

export type ProjectTransactionOperation = 'apply' | 'undo'
export type ProjectTransactionPhase = 'prepared' | 'committed' | 'restoring'

export interface ProjectTransactionFile {
  path: string
  beforeSha256: string | null
  afterSha256: string
}

export interface ProjectTransactionJournal {
  schemaVersion: typeof JOURNAL_SCHEMA_VERSION
  id: string
  operation: ProjectTransactionOperation
  phase: ProjectTransactionPhase
  projectId: string
  projectRevision: number
  nextProjectRevision: number
  snapshotId: string
  previousRollback: string | null
  files: ProjectTransactionFile[]
}

export interface ProjectTransactionStateOptions {
  emptyLockStaleMs?: number
  now?: () => number
  isProcessAlive?: (pid: number) => boolean
}

interface ProjectRevisionRecord {
  schemaVersion: typeof PROJECT_STATE_SCHEMA_VERSION
  projectId: string
  revision: number
}

interface ProjectLockRecord {
  schemaVersion: typeof LOCK_SCHEMA_VERSION
  pid: number
  token: string
  createdAt: number
}

interface HeldProjectLock {
  handle: import('node:fs/promises').FileHandle
  token: string
  dev: number
  ino: number
}

export class ProjectRevisionConflictError extends Error {
  readonly code = 'LATEX_PROJECT_REVISION_CONFLICT'

  constructor(projectId: string) {
    super(`Project transaction revision changed: ${projectId}`)
    this.name = 'ProjectRevisionConflictError'
  }
}

export class ProjectTransactionState {
  private readonly journalsRoot: string
  private readonly projectsRoot: string
  private readonly locksRoot: string
  private readonly emptyLockStaleMs: number
  private readonly now: () => number
  private readonly processAlive: (pid: number) => boolean

  constructor(cacheRoot: string, options: ProjectTransactionStateOptions = {}) {
    const root = join(cacheRoot, 'proposals')
    this.journalsRoot = join(root, 'transactions')
    this.projectsRoot = join(root, 'projects')
    this.locksRoot = join(root, 'project-locks')
    this.emptyLockStaleMs = options.emptyLockStaleMs ?? DEFAULT_EMPTY_LOCK_STALE_MS
    this.now = options.now ?? Date.now
    this.processAlive = options.isProcessAlive ?? isProcessAlive
    if (!Number.isSafeInteger(this.emptyLockStaleMs) || this.emptyLockStaleMs <= 0) {
      throw new Error('emptyLockStaleMs must be a positive safe integer')
    }
  }

  async withProjectLock<T>(projectId: string, action: () => Promise<T>): Promise<T> {
    validateProjectId(projectId)
    await this.ensureDirectories()
    const path = this.lockPath(projectId)
    const lock = await this.acquireLock(path)
    try {
      return await action()
    } finally {
      await releaseLock(path, lock)
    }
  }

  async readRevision(projectId: string): Promise<number> {
    validateProjectId(projectId)
    try {
      const value: unknown = JSON.parse(await readFile(this.projectStatePath(projectId), 'utf8'))
      if (!isRevisionRecord(value) || value.projectId !== projectId) {
        throw new Error('Invalid project transaction revision state')
      }
      return value.revision
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return 0
      throw error
    }
  }

  async advanceRevision(
    projectId: string,
    expectedRevision: number,
    nextRevision: number,
  ): Promise<void> {
    validateProjectId(projectId)
    validateRevision(expectedRevision)
    validateRevision(nextRevision)
    if (nextRevision !== expectedRevision + 1) {
      throw new Error('Project transaction revision must advance by exactly one')
    }
    const currentRevision = await this.readRevision(projectId)
    if (currentRevision === nextRevision) return
    if (currentRevision !== expectedRevision) {
      throw new ProjectRevisionConflictError(projectId)
    }
    await mkdir(this.projectsRoot, { recursive: true })
    const record: ProjectRevisionRecord = {
      schemaVersion: PROJECT_STATE_SCHEMA_VERSION,
      projectId,
      revision: nextRevision,
    }
    const data = Buffer.from(`${JSON.stringify(record, null, 2)}\n`)
    try {
      await atomicWriteFile(this.projectStatePath(projectId), data)
    } catch (error) {
      if (
        error instanceof AtomicWriteCommittedError &&
        (await this.readRevision(projectId)) === nextRevision
      ) {
        return
      }
      throw error
    }
  }

  async listJournals(projectId: string): Promise<ProjectTransactionJournal[]> {
    validateProjectId(projectId)
    let entries
    try {
      entries = await readdir(this.journalsRoot, { withFileTypes: true })
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return []
      throw error
    }
    const journals: ProjectTransactionJournal[] = []
    for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name))) {
      if (!entry.isFile() || !entry.name.endsWith('.json')) continue
      const id = entry.name.slice(0, -'.json'.length)
      validateTransactionId(id)
      const journal = await this.readJournal(id)
      if (journal.projectId === projectId) journals.push(journal)
    }
    return journals
  }

  async readJournal(id: string): Promise<ProjectTransactionJournal> {
    validateTransactionId(id)
    let value: unknown
    try {
      value = JSON.parse(await readFile(this.journalPath(id), 'utf8'))
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
        throw new Error(`Project transaction journal not found: ${id}`, { cause: error })
      }
      throw error
    }
    if (!isJournal(value) || value.id !== id) {
      throw new Error(`Invalid project transaction journal: ${id}`)
    }
    return value
  }

  async writeJournal(journal: ProjectTransactionJournal): Promise<void> {
    if (!isJournal(journal)) throw new Error('Invalid project transaction journal')
    await mkdir(this.journalsRoot, { recursive: true })
    const existing = await this.readOptionalJournal(journal.id)
    assertPhaseTransition(existing, journal)
    const data = Buffer.from(`${JSON.stringify(journal, null, 2)}\n`)
    try {
      await atomicWriteFile(this.journalPath(journal.id), data)
    } catch (error) {
      if (error instanceof AtomicWriteCommittedError) {
        const persisted = await this.readOptionalJournal(journal.id)
        if (persisted && journalsEqual(persisted, journal)) return
      }
      throw error
    }
  }

  async deleteJournal(id: string): Promise<void> {
    validateTransactionId(id)
    try {
      await unlink(this.journalPath(id))
      await syncDirectory(this.journalsRoot)
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
    }
  }

  private async readOptionalJournal(id: string): Promise<ProjectTransactionJournal | undefined> {
    try {
      return await this.readJournal(id)
    } catch (error) {
      if (
        error instanceof Error &&
        error.message === `Project transaction journal not found: ${id}`
      ) {
        return undefined
      }
      throw error
    }
  }

  private async ensureDirectories(): Promise<void> {
    await Promise.all(
      [this.journalsRoot, this.projectsRoot, this.locksRoot].map((path) =>
        mkdir(path, { recursive: true }),
      ),
    )
  }

  private journalPath(id: string): string {
    validateTransactionId(id)
    return join(this.journalsRoot, `${id}.json`)
  }

  private projectStatePath(projectId: string): string {
    validateProjectId(projectId)
    return join(this.projectsRoot, `${projectId}.json`)
  }

  private lockPath(projectId: string): string {
    validateProjectId(projectId)
    return join(this.locksRoot, `${projectId}.lock`)
  }

  private async acquireLock(path: string): Promise<HeldProjectLock> {
    const token = randomBytes(16).toString('hex')
    const noFollow = 'O_NOFOLLOW' in constants ? constants.O_NOFOLLOW : 0
    for (let attempt = 0; attempt < 2; attempt += 1) {
      let handle: import('node:fs/promises').FileHandle | undefined
      let ownedIdentity: Pick<HeldProjectLock, 'dev' | 'ino'> | undefined
      try {
        handle = await open(
          path,
          constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | noFollow,
          0o600,
        )
        const info = await handle.stat()
        ownedIdentity = { dev: info.dev, ino: info.ino }
        const record: ProjectLockRecord = {
          schemaVersion: LOCK_SCHEMA_VERSION,
          pid: process.pid,
          token,
          createdAt: this.now(),
        }
        await handle.writeFile(`${JSON.stringify(record)}\n`)
        await handle.sync()
        return { handle, token, ...ownedIdentity }
      } catch (error) {
        await handle?.close().catch(() => undefined)
        if (ownedIdentity) await unlinkOwned(path, ownedIdentity)
        if (
          attempt === 0 &&
          (error as NodeJS.ErrnoException).code === 'EEXIST' &&
          (await this.removeStaleLock(path))
        ) {
          continue
        }
        throw error
      }
    }
    throw new Error('Unable to acquire project transaction lock')
  }

  private async removeStaleLock(path: string): Promise<boolean> {
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
    const record = parseLock(contents)
    const stale = record
      ? !this.processAlive(record.pid)
      : this.now() - before.mtimeMs >= this.emptyLockStaleMs
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
}

function assertPhaseTransition(
  existing: ProjectTransactionJournal | undefined,
  next: ProjectTransactionJournal,
): void {
  if (!existing) {
    if (
      (next.operation === 'apply' && next.phase === 'prepared') ||
      (next.operation === 'undo' && next.phase === 'restoring')
    ) {
      return
    }
    throw new Error('Project transaction journal must start before its commit point')
  }
  if (
    existing.id !== next.id ||
    existing.operation !== next.operation ||
    existing.projectId !== next.projectId ||
    existing.projectRevision !== next.projectRevision ||
    existing.nextProjectRevision !== next.nextProjectRevision ||
    existing.snapshotId !== next.snapshotId ||
    existing.previousRollback !== next.previousRollback ||
    JSON.stringify(existing.files) !== JSON.stringify(next.files)
  ) {
    throw new Error('Project transaction journal identity changed')
  }
  const valid =
    journalsEqual(existing, next) ||
    (existing.operation === 'apply' && existing.phase === 'prepared' && next.phase === 'committed')
  if (!valid) throw new Error('Invalid project transaction journal phase transition')
}

function isJournal(value: unknown): value is ProjectTransactionJournal {
  if (!value || typeof value !== 'object') return false
  const journal = value as Partial<ProjectTransactionJournal>
  if (
    journal.schemaVersion !== JOURNAL_SCHEMA_VERSION ||
    !isSafeTransactionId(journal.id) ||
    !isSafeProjectId(journal.projectId) ||
    !isSnapshotId(journal.snapshotId) ||
    (journal.previousRollback !== null && !isSnapshotId(journal.previousRollback)) ||
    !isRevision(journal.projectRevision) ||
    !isRevision(journal.nextProjectRevision) ||
    journal.nextProjectRevision !== journal.projectRevision + 1 ||
    !Array.isArray(journal.files) ||
    journal.files.length === 0 ||
    journal.files.length > 100
  ) {
    return false
  }
  if (!(
    (journal.operation === 'apply' &&
      (journal.phase === 'prepared' || journal.phase === 'committed')) ||
    (journal.operation === 'undo' && journal.phase === 'restoring')
  )) {
    return false
  }
  const paths = new Set<string>()
  return journal.files.every((file) => {
    if (!file || typeof file !== 'object') return false
    if (
      !isCanonicalTextPath(file.path) ||
      paths.has(file.path) ||
      (file.beforeSha256 !== null && !isSha256(file.beforeSha256)) ||
      !isSha256(file.afterSha256)
    ) {
      return false
    }
    paths.add(file.path)
    return true
  })
}

function isRevisionRecord(value: unknown): value is ProjectRevisionRecord {
  if (!value || typeof value !== 'object') return false
  const record = value as Partial<ProjectRevisionRecord>
  return (
    record.schemaVersion === PROJECT_STATE_SCHEMA_VERSION &&
    isSafeProjectId(record.projectId) &&
    isRevision(record.revision)
  )
}

function parseLock(contents: string): ProjectLockRecord | undefined {
  try {
    const value: unknown = JSON.parse(contents)
    if (!value || typeof value !== 'object') return undefined
    const record = value as Partial<ProjectLockRecord>
    if (
      record.schemaVersion !== LOCK_SCHEMA_VERSION ||
      !Number.isSafeInteger(record.pid) ||
      (record.pid ?? 0) <= 0 ||
      typeof record.token !== 'string' ||
      !/^[a-f0-9]{32}$/.test(record.token) ||
      !Number.isSafeInteger(record.createdAt)
    ) {
      return undefined
    }
    return record as ProjectLockRecord
  } catch {
    return undefined
  }
}

async function releaseLock(path: string, lock: HeldProjectLock): Promise<void> {
  await lock.handle.close().catch(() => undefined)
  try {
    const info = await lstat(path)
    if (info.dev !== lock.dev || info.ino !== lock.ino) return
    const record = parseLock(await readFile(path, 'utf8'))
    if (record?.token === lock.token && record.pid === process.pid) await unlink(path)
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
  }
}

async function unlinkOwned(
  path: string,
  identity: Pick<HeldProjectLock, 'dev' | 'ino'>,
): Promise<void> {
  try {
    const info = await lstat(path)
    if (info.dev === identity.dev && info.ino === identity.ino) await unlink(path)
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
  }
}

async function syncDirectory(path: string): Promise<void> {
  let handle
  try {
    handle = await open(path, constants.O_RDONLY)
    await handle.sync()
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code ?? ''
    if (
      process.platform !== 'win32' ||
      !new Set(['EACCES', 'EINVAL', 'EISDIR', 'ENOTSUP', 'EPERM']).has(code)
    ) {
      throw error
    }
  } finally {
    await handle?.close().catch(() => undefined)
  }
}

function journalsEqual(left: ProjectTransactionJournal, right: ProjectTransactionJournal): boolean {
  return JSON.stringify(left) === JSON.stringify(right)
}

function validateProjectId(id: string): void {
  if (!isSafeProjectId(id)) throw new Error('Invalid project ID')
}

function validateTransactionId(id: string): void {
  if (!isSafeTransactionId(id)) throw new Error('Invalid project transaction ID')
}

function validateRevision(revision: number): void {
  if (!isRevision(revision)) throw new Error('Invalid project transaction revision')
}

function isSafeProjectId(value: unknown): value is string {
  return (
    typeof value === 'string' &&
    /^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/.test(value) &&
    !['__proto__', 'constructor', 'prototype'].includes(value)
  )
}

function isSafeTransactionId(value: unknown): value is string {
  return (
    typeof value === 'string' &&
    /^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/.test(value) &&
    !['__proto__', 'constructor', 'prototype'].includes(value)
  )
}

function isSnapshotId(value: unknown): value is string {
  return typeof value === 'string' && /^[a-f0-9]{32}$/.test(value)
}

function isRevision(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isSha256(value: unknown): value is string {
  return typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
}

function isCanonicalTextPath(value: unknown): value is string {
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
  const segments = value.split('/')
  if (segments.some((segment) => !segment || segment === '.' || segment === '..')) return false
  const lower = value.toLowerCase()
  return lower === 'tectonic.toml' || TEXT_EXTENSIONS.has(extname(lower))
}

function isProcessAlive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return (error as NodeJS.ErrnoException).code !== 'ESRCH'
  }
}
