import { createHash } from 'node:crypto'
import { access, mkdir, readFile, rename, unlink } from 'node:fs/promises'
import { extname, join } from 'node:path'
import { atomicWriteFile } from './atomic-write.js'
import { ProjectPathPolicy } from './path-policy.js'
import type { LatexProject } from './project.js'
import { SnapshotStore } from './snapshot.js'
import { DEFAULT_MAX_TEXT_BYTES } from './types.js'

const PROPOSAL_SCHEMA_VERSION = 1
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

export interface ProposalFile {
  path: string
  beforeSha256: string | null
  afterText: string
}

export interface EditProposal {
  id: string
  projectId: string
  expiresAt: number
  files: ProposalFile[]
}

interface StoredProposal extends EditProposal {
  schemaVersion: typeof PROPOSAL_SCHEMA_VERSION
}

interface RollbackRecord {
  schemaVersion: 1
  snapshotId: string
  projectId: string
  expectedAbsentHashes: Record<string, string>
  expectedCurrentHashes: Record<string, string>
}

export interface ProposalStoreOptions {
  maxFileBytes?: number
  maxFiles?: number
  now?: () => number
  /** @internal deterministic transaction failure hook for tests. */
  writeText?: (project: LatexProject, path: string, text: string) => Promise<unknown>
}

export interface AppliedProposal {
  proposalId: string
  snapshotId: string
}

export class ProposalStore {
  private readonly proposalsRoot: string
  private readonly consumedRoot: string
  private readonly rollbackRoot: string
  private readonly undoRoot: string
  private readonly maxFileBytes: number
  private readonly maxFiles: number
  private readonly now: () => number
  private readonly writeText: (
    project: LatexProject,
    path: string,
    text: string,
  ) => Promise<unknown>

  constructor(
    cacheRoot: string,
    private readonly snapshots: SnapshotStore,
    options: ProposalStoreOptions = {},
  ) {
    const root = join(cacheRoot, 'proposals')
    this.proposalsRoot = join(root, 'pending')
    this.consumedRoot = join(root, 'consumed')
    this.rollbackRoot = join(root, 'rollbacks')
    this.undoRoot = join(root, 'undone')
    this.maxFileBytes = options.maxFileBytes ?? DEFAULT_MAX_TEXT_BYTES
    this.maxFiles = options.maxFiles ?? 100
    this.now = options.now ?? Date.now
    this.writeText = options.writeText ?? ((project, path, text) => project.saveText(path, text))
    if (!Number.isSafeInteger(this.maxFileBytes) || this.maxFileBytes <= 0) {
      throw new Error('Proposal maxFileBytes must be a positive safe integer')
    }
    if (!Number.isSafeInteger(this.maxFiles) || this.maxFiles <= 0) {
      throw new Error('Proposal maxFiles must be a positive safe integer')
    }
  }

  async create(proposal: EditProposal): Promise<void> {
    const stored = validateProposal(proposal, this.maxFileBytes, this.maxFiles, this.now())
    await this.ensureDirectories()
    if (await exists(this.consumedPath(stored.id))) throw new Error('Proposal was already consumed')
    try {
      await atomicWriteFile(
        this.pendingPath(stored.id),
        Buffer.from(`${JSON.stringify(stored, null, 2)}\n`),
        {
          validateBeforeRename: async () => {
            if (await exists(this.pendingPath(stored.id)))
              throw new Error('Proposal already exists')
            if (await exists(this.consumedPath(stored.id)))
              throw new Error('Proposal was already consumed')
          },
        },
      )
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'EEXIST')
        throw new Error('Proposal already exists')
      throw error
    }
  }

  async apply(
    proposalId: string,
    projectId: string,
    project: LatexProject,
  ): Promise<AppliedProposal> {
    validateProposalId(proposalId)
    validateProjectId(projectId)
    const proposal = await this.claim(proposalId)
    if (proposal.projectId !== projectId) throw new Error('Proposal belongs to another project')
    if (proposal.expiresAt <= this.now()) throw new Error('Proposal has expired')

    const policy = await ProjectPathPolicy.open(project.rootPath)
    const prepared = await Promise.all(
      proposal.files.map(async (file) => {
        const path = policy.normalize(file.path)
        const before = await readOptionalText(project, policy, path)
        const beforeHash = before === undefined ? null : digest(before)
        if (beforeHash !== file.beforeSha256) {
          throw new Error(`Proposal baseline conflict: ${path}`)
        }
        return { ...file, path, afterSha256: digest(file.afterText) }
      }),
    )

    const previousRollback = await this.snapshots.getCurrentRollback(projectId)
    const snapshot = await this.snapshots.create(
      projectId,
      project,
      prepared.map((file) => file.path),
    )
    const expectedAbsentHashes = new Map(
      prepared
        .filter((file) => file.beforeSha256 === null)
        .map((file) => [file.path, file.afterSha256] as const),
    )
    await assertPreparedBaselines(project, policy, prepared)
    const rollbackPath = this.rollbackPath(snapshot.id)
    const record: RollbackRecord = {
      schemaVersion: 1,
      snapshotId: snapshot.id,
      projectId,
      expectedAbsentHashes: Object.fromEntries(expectedAbsentHashes),
      expectedCurrentHashes: Object.fromEntries(
        prepared.map((file) => [file.path, file.afterSha256]),
      ),
    }

    try {
      for (const file of prepared) await this.writeText(project, file.path, file.afterText)
      await atomicWriteFile(rollbackPath, Buffer.from(`${JSON.stringify(record, null, 2)}\n`))
      await this.snapshots.setCurrentRollback(projectId, snapshot.id)
      return { proposalId, snapshotId: snapshot.id }
    } catch (transactionError) {
      const cleanupErrors: unknown[] = []
      try {
        await this.snapshots.restore(projectId, snapshot.id, project, { expectedAbsentHashes })
      } catch (rollbackError) {
        cleanupErrors.push(rollbackError)
      }
      try {
        await this.snapshots.setCurrentRollback(projectId, previousRollback ?? null)
      } catch (rollbackPointError) {
        cleanupErrors.push(rollbackPointError)
      }
      try {
        await unlink(rollbackPath)
      } catch (unlinkError) {
        if ((unlinkError as NodeJS.ErrnoException).code !== 'ENOENT')
          cleanupErrors.push(unlinkError)
      }
      if (cleanupErrors.length > 0) {
        throw new AggregateError(
          [transactionError, ...cleanupErrors],
          'Proposal transaction failed and cleanup was incomplete',
        )
      }
      throw transactionError
    }
  }

  async undo(
    projectId: string,
    snapshotId: string,
    project: LatexProject,
  ): Promise<{ snapshotId: string; restored: boolean }> {
    validateProjectId(projectId)
    validateSnapshotId(snapshotId)
    const markerPath = this.undoPath(snapshotId)
    if (await undoMarkerExists(markerPath)) {
      await this.snapshots.verifyRestored(projectId, snapshotId, project)
      if ((await this.snapshots.getCurrentRollback(projectId)) === snapshotId) {
        await this.snapshots.setCurrentRollback(projectId, null)
      }
      return { snapshotId, restored: false }
    }
    if ((await this.snapshots.getCurrentRollback(projectId)) !== snapshotId) {
      throw new Error('Snapshot is not the current rollback point')
    }
    try {
      await this.snapshots.verifyRestored(projectId, snapshotId, project)
      await atomicWriteFile(
        markerPath,
        Buffer.from(`${JSON.stringify({ schemaVersion: 1, projectId, snapshotId })}\n`),
      )
      await this.snapshots.setCurrentRollback(projectId, null)
      return { snapshotId, restored: false }
    } catch (error) {
      if (!isProjectChangedAfterUndoError(error)) throw error
    }
    const record = await this.readRollback(snapshotId)
    if (record.projectId !== projectId) throw new Error('Rollback belongs to another project')
    const snapshot = (await this.snapshots.list(projectId)).find((item) => item.id === snapshotId)
    const expectedCurrentHashes = new Map(Object.entries(record.expectedCurrentHashes))
    if (
      !snapshot ||
      snapshot.paths.length !== expectedCurrentHashes.size ||
      snapshot.paths.some((path) => !expectedCurrentHashes.has(path))
    ) {
      throw new Error('Invalid rollback record paths')
    }
    const policy = await ProjectPathPolicy.open(project.rootPath)
    await assertExpectedCurrentHashes(project, policy, expectedCurrentHashes)
    await this.snapshots.restore(projectId, snapshotId, project, {
      expectedAbsentHashes: new Map(Object.entries(record.expectedAbsentHashes)),
    })
    await atomicWriteFile(
      markerPath,
      Buffer.from(`${JSON.stringify({ schemaVersion: 1, projectId, snapshotId })}\n`),
    )
    await this.snapshots.setCurrentRollback(projectId, null)
    return { snapshotId, restored: true }
  }

  private async claim(proposalId: string): Promise<StoredProposal> {
    await this.ensureDirectories()
    const pending = this.pendingPath(proposalId)
    const consumed = this.consumedPath(proposalId)
    try {
      await rename(pending, consumed)
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT' && (await exists(consumed))) {
        throw new Error('Proposal was already consumed')
      }
      throw new Error(`Proposal not found: ${proposalId}`, { cause: error })
    }
    return readStoredProposal(consumed, this.maxFileBytes, this.maxFiles)
  }

  private async readRollback(snapshotId: string): Promise<RollbackRecord> {
    const value: unknown = JSON.parse(await readFile(this.rollbackPath(snapshotId), 'utf8'))
    if (!isRollbackRecord(value) || value.snapshotId !== snapshotId) {
      throw new Error('Invalid rollback record')
    }
    return value
  }

  private async ensureDirectories(): Promise<void> {
    await Promise.all(
      [this.proposalsRoot, this.consumedRoot, this.rollbackRoot, this.undoRoot].map((path) =>
        mkdir(path, { recursive: true }),
      ),
    )
  }

  private pendingPath(id: string): string {
    validateProposalId(id)
    return join(this.proposalsRoot, `${id}.json`)
  }

  private consumedPath(id: string): string {
    validateProposalId(id)
    return join(this.consumedRoot, `${id}.json`)
  }

  private rollbackPath(id: string): string {
    validateSnapshotId(id)
    return join(this.rollbackRoot, `${id}.json`)
  }

  private undoPath(id: string): string {
    validateSnapshotId(id)
    return join(this.undoRoot, `${id}.json`)
  }
}

function validateProposal(
  proposal: EditProposal,
  maxFileBytes: number,
  maxFiles: number,
  now: number,
): StoredProposal {
  if (!proposal || typeof proposal !== 'object') throw new Error('Invalid proposal')
  validateProposalId(proposal.id)
  validateProjectId(proposal.projectId)
  if (!Number.isSafeInteger(proposal.expiresAt) || proposal.expiresAt <= now) {
    throw new Error('Proposal expiry must be in the future')
  }
  if (
    !Array.isArray(proposal.files) ||
    proposal.files.length === 0 ||
    proposal.files.length > maxFiles
  ) {
    throw new Error('Proposal file count is invalid')
  }
  const paths = new Set<string>()
  const files = proposal.files.map((file) => {
    if (!file || typeof file !== 'object') throw new Error('Invalid proposal file')
    const path = normalizePortablePath(file.path)
    if (paths.has(path)) throw new Error('Proposal contains duplicate normalized paths')
    paths.add(path)
    if (!isTextPath(path)) throw new Error(`Unsupported text file type: ${path}`)
    if (file.beforeSha256 !== null && !/^[a-f0-9]{64}$/.test(file.beforeSha256)) {
      throw new Error(`Invalid proposal baseline hash: ${path}`)
    }
    if (typeof file.afterText !== 'string') throw new Error('Proposal deletion is not supported')
    assertValidUtf8String(file.afterText, path)
    if (Buffer.byteLength(file.afterText, 'utf8') > maxFileBytes) {
      throw new Error(`Proposal text exceeds size limit: ${path}`)
    }
    return { path, beforeSha256: file.beforeSha256, afterText: file.afterText }
  })
  return {
    schemaVersion: PROPOSAL_SCHEMA_VERSION,
    id: proposal.id,
    projectId: proposal.projectId,
    expiresAt: proposal.expiresAt,
    files,
  }
}

async function readStoredProposal(
  path: string,
  maxFileBytes: number,
  maxFiles: number,
): Promise<StoredProposal> {
  const value: unknown = JSON.parse(await readFile(path, 'utf8'))
  if (!value || typeof value !== 'object') throw new Error('Invalid stored proposal')
  const item = value as Partial<StoredProposal>
  if (item.schemaVersion !== PROPOSAL_SCHEMA_VERSION) throw new Error('Invalid stored proposal')
  return validateProposal(value as EditProposal, maxFileBytes, maxFiles, Number.MIN_SAFE_INTEGER)
}

async function readOptionalText(
  project: LatexProject,
  policy: ProjectPathPolicy,
  path: string,
): Promise<string | undefined> {
  try {
    await policy.resolveExisting(path, 'file')
  } catch (error) {
    if (error instanceof Error && /does not exist/.test(error.message)) return undefined
    throw error
  }
  return project.readText(path)
}

function normalizePortablePath(path: unknown): string {
  if (typeof path !== 'string' || !path || path.includes('\0'))
    throw new Error('Invalid proposal path')
  if (path.startsWith('/') || path.startsWith('\\') || /^[a-zA-Z]:[\\/]/.test(path)) {
    throw new Error('Proposal path must be relative')
  }
  const segments = path.replaceAll('\\', '/').split('/')
  if (segments.includes('..')) throw new Error('Proposal path traversal is not allowed')
  const normalized = segments.filter((segment) => segment && segment !== '.').join('/')
  if (!normalized) throw new Error('Invalid proposal path')
  return normalized
}

function isTextPath(path: string): boolean {
  const lower = path.toLowerCase()
  return lower === 'tectonic.toml' || TEXT_EXTENSIONS.has(extname(lower))
}

function assertValidUtf8String(text: string, path: string): void {
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index)
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = text.charCodeAt(index + 1)
      if (index + 1 >= text.length || next < 0xdc00 || next > 0xdfff) {
        throw new Error(`Proposal text is not valid UTF-8: ${path}`)
      }
      index += 1
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      throw new Error(`Proposal text is not valid UTF-8: ${path}`)
    }
  }
}

function digest(text: string): string {
  return createHash('sha256').update(text, 'utf8').digest('hex')
}

function validateProposalId(id: string): void {
  if (!/^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/.test(id)) throw new Error('Invalid proposal ID')
}

function validateProjectId(id: string): void {
  if (
    !/^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/.test(id) ||
    ['__proto__', 'constructor', 'prototype'].includes(id)
  ) {
    throw new Error('Invalid project ID')
  }
}

function validateSnapshotId(id: string): void {
  if (!/^[a-f0-9]{32}$/.test(id)) throw new Error('Invalid snapshot ID')
}

async function undoMarkerExists(path: string): Promise<boolean> {
  try {
    await access(path)
    return true
  } catch (error) {
    if (['ENOENT', 'ENOTDIR'].includes((error as NodeJS.ErrnoException).code ?? '')) return false
    throw error
  }
}
async function exists(path: string): Promise<boolean> {
  try {
    await access(path)
    return true
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return false
    throw error
  }
}

function isProjectChangedAfterUndoError(error: unknown): boolean {
  return error instanceof Error && error.message.startsWith('Project changed after undo:')
}

function isRollbackRecord(value: unknown): value is RollbackRecord {
  if (!value || typeof value !== 'object') return false
  const record = value as Partial<RollbackRecord>
  if (
    record.schemaVersion !== 1 ||
    typeof record.snapshotId !== 'string' ||
    !/^[a-f0-9]{32}$/.test(record.snapshotId) ||
    typeof record.projectId !== 'string' ||
    !/^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/.test(record.projectId) ||
    ['__proto__', 'constructor', 'prototype'].includes(record.projectId) ||
    !isHashRecord(record.expectedAbsentHashes) ||
    !isHashRecord(record.expectedCurrentHashes)
  ) {
    return false
  }
  return Object.entries(record.expectedAbsentHashes).every(
    ([path, hash]) => record.expectedCurrentHashes?.[path] === hash,
  )
}

function isHashRecord(value: unknown): value is Record<string, string> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  return Object.entries(value).every(([path, hash]) => {
    try {
      return (
        normalizePortablePath(path) === path &&
        isTextPath(path) &&
        typeof hash === 'string' &&
        /^[a-f0-9]{64}$/.test(hash)
      )
    } catch {
      return false
    }
  })
}

async function assertPreparedBaselines(
  project: LatexProject,
  policy: ProjectPathPolicy,
  files: readonly Pick<ProposalFile, 'path' | 'beforeSha256'>[],
): Promise<void> {
  for (const file of files) {
    const before = await readOptionalText(project, policy, file.path)
    const currentHash = before === undefined ? null : digest(before)
    if (currentHash !== file.beforeSha256) {
      throw new Error(`Proposal baseline conflict after snapshot: ${file.path}`)
    }
  }
}

async function assertExpectedCurrentHashes(
  project: LatexProject,
  policy: ProjectPathPolicy,
  expected: ReadonlyMap<string, string>,
): Promise<void> {
  for (const [path, hash] of expected) {
    const current = await readOptionalText(project, policy, path)
    if (current === undefined || digest(current) !== hash) {
      throw new Error(`Project changed after proposal apply: ${path}`)
    }
  }
}
