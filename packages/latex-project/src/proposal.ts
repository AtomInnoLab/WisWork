import { createHash } from 'node:crypto'
import { access, mkdir, readFile, rename, unlink } from 'node:fs/promises'
import { extname, join } from 'node:path'
import { atomicWriteFile } from './atomic-write.js'
import { ProjectPathPolicy } from './path-policy.js'
import type { LatexProject } from './project.js'
import { SnapshotStore } from './snapshot.js'
import { DEFAULT_MAX_TEXT_BYTES } from './types.js'
import { ProjectTransactionState, type ProjectTransactionJournal } from './transaction-state.js'

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
  projectRevision: number
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
  writeText?: (
    project: LatexProject,
    path: string,
    text: string,
    expectedSha256: string | null,
  ) => Promise<unknown>
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
    expectedSha256: string | null,
  ) => Promise<unknown>
  private readonly transactionState: ProjectTransactionState

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
    this.writeText =
      options.writeText ??
      ((project, path, text, expectedSha256) => project.saveText(path, text, { expectedSha256 }))
    this.transactionState = new ProjectTransactionState(cacheRoot)
    if (!Number.isSafeInteger(this.maxFileBytes) || this.maxFileBytes <= 0) {
      throw new Error('Proposal maxFileBytes must be a positive safe integer')
    }
    if (!Number.isSafeInteger(this.maxFiles) || this.maxFiles <= 0) {
      throw new Error('Proposal maxFiles must be a positive safe integer')
    }
  }

  async create(proposal: EditProposal): Promise<void> {
    const normalized = validateProposal(proposal, this.maxFileBytes, this.maxFiles, this.now(), 0)
    await this.transactionState.withProjectLock(proposal.projectId, async () => {
      const stored: StoredProposal = {
        ...normalized,
        projectRevision: await this.transactionState.readRevision(proposal.projectId),
      }
      await this.ensureDirectories()
      if (await exists(this.consumedPath(stored.id))) {
        throw new Error('Proposal was already consumed')
      }
      try {
        await atomicWriteFile(
          this.pendingPath(stored.id),
          Buffer.from(`${JSON.stringify(stored, null, 2)}\n`),
          {
            validateBeforeRename: async () => {
              if (await exists(this.pendingPath(stored.id))) {
                throw new Error('Proposal already exists')
              }
              if (await exists(this.consumedPath(stored.id))) {
                throw new Error('Proposal was already consumed')
              }
            },
          },
        )
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'EEXIST') {
          throw new Error('Proposal already exists', { cause: error })
        }
        throw error
      }
    })
  }

  async apply(
    proposalId: string,
    projectId: string,
    project: LatexProject,
  ): Promise<AppliedProposal> {
    validateProposalId(proposalId)
    validateProjectId(projectId)
    return this.transactionState.withProjectLock(projectId, async () => {
      await this.recoverLocked(projectId, project)
      return this.applyLocked(proposalId, projectId, project)
    })
  }

  async discard(proposalId: string, projectId: string): Promise<boolean> {
    validateProposalId(proposalId)
    validateProjectId(projectId)
    return this.transactionState.withProjectLock(projectId, async () => {
      await this.ensureDirectories()
      const pending = this.pendingPath(proposalId)
      if (!(await exists(pending))) return false
      const proposal = await readStoredProposal(pending, this.maxFileBytes, this.maxFiles)
      if (proposal.projectId !== projectId) throw new Error('Proposal belongs to another project')
      try {
        await unlink(pending)
        return true
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') return false
        throw error
      }
    })
  }

  async discardPreparedSnapshot(
    proposalId: string,
    projectId: string,
    snapshotId: string,
  ): Promise<boolean> {
    validateProposalId(proposalId)
    validateProjectId(projectId)
    validateSnapshotId(snapshotId)
    return this.transactionState.withProjectLock(projectId, async () => {
      const referenced = (await this.transactionState.listJournals(projectId)).some(
        (journal) => journal.id === proposalId && journal.snapshotId === snapshotId,
      )
      if (referenced) {
        throw new Error('Prepared snapshot is retained for transaction recovery')
      }
      return this.snapshots.discard(projectId, snapshotId)
    })
  }

  async applyPrepared(
    proposalId: string,
    projectId: string,
    snapshotId: string,
    project: LatexProject,
    validateBeforeCommit?: () => void | Promise<void>,
  ): Promise<AppliedProposal> {
    validateProposalId(proposalId)
    validateProjectId(projectId)
    validateSnapshotId(snapshotId)
    return this.transactionState.withProjectLock(projectId, async () => {
      await this.recoverLocked(projectId, project)
      return this.applyLocked(proposalId, projectId, project, snapshotId, validateBeforeCommit)
    })
  }

  private async applyLocked(
    proposalId: string,
    projectId: string,
    project: LatexProject,
    preparedSnapshotId?: string,
    validateBeforeCommit?: () => void | Promise<void>,
  ): Promise<AppliedProposal> {
    const proposal = await this.claim(proposalId)
    if (proposal.projectId !== projectId) throw new Error('Proposal belongs to another project')
    if (proposal.expiresAt <= this.now()) throw new Error('Proposal has expired')
    if ((await this.transactionState.readRevision(projectId)) !== proposal.projectRevision) {
      throw new Error('Proposal baseline revision changed')
    }

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
    const snapshot = preparedSnapshotId
      ? { id: preparedSnapshotId }
      : await this.snapshots.create(
          projectId,
          project,
          prepared.map((file) => file.path),
        )
    if (preparedSnapshotId) {
      const hashes = await this.snapshots.getFileHashes(projectId, preparedSnapshotId)
      if (
        hashes.size !== prepared.length ||
        prepared.some(
          (file) => !hashes.has(file.path) || hashes.get(file.path) !== file.beforeSha256,
        )
      ) {
        throw new Error('Prepared snapshot does not match proposal baseline')
      }
    }
    await assertPreparedBaselines(project, policy, prepared)
    await validateBeforeCommit?.()
    const journal: ProjectTransactionJournal = {
      schemaVersion: 1,
      id: proposal.id,
      operation: 'apply',
      phase: 'prepared',
      projectId,
      projectRevision: proposal.projectRevision,
      nextProjectRevision: proposal.projectRevision + 1,
      snapshotId: snapshot.id,
      previousRollback: previousRollback ?? null,
      files: prepared.map((file) => ({
        path: file.path,
        beforeSha256: file.beforeSha256,
        afterSha256: file.afterSha256,
      })),
    }
    await this.transactionState.writeJournal(journal)

    try {
      // Cancellation/revision validation must be adjacent to the first domain write. The journal
      // is recoverable if this second check rejects; no unguarded await follows before the loop.
      await validateBeforeCommit?.()
      for (const file of prepared) {
        await this.writeText(project, file.path, file.afterText, file.beforeSha256)
      }
      await assertExpectedCurrentHashes(
        project,
        policy,
        new Map(prepared.map((file) => [file.path, file.afterSha256])),
      )
      await this.transactionState.writeJournal({ ...journal, phase: 'committed' })
    } catch (transactionError) {
      await this.rollbackOrAggregate(journal, project, transactionError)
      throw transactionError
    }

    const committed = { ...journal, phase: 'committed' as const }
    await this.finalizeApplyJournal(committed, project)
    return { proposalId, snapshotId: snapshot.id }
  }

  async recover(projectId: string, project: LatexProject): Promise<{ recovered: number }> {
    validateProjectId(projectId)
    return this.transactionState.withProjectLock(projectId, async () => ({
      recovered: await this.recoverLocked(projectId, project),
    }))
  }

  private async recoverLocked(projectId: string, project: LatexProject): Promise<number> {
    const journals = await this.transactionState.listJournals(projectId)
    for (const journal of journals) {
      const revision = await this.transactionState.readRevision(projectId)
      if (revision !== journal.projectRevision && revision !== journal.nextProjectRevision) {
        throw new Error('Project transaction journal revision is inconsistent')
      }
      if (journal.operation === 'apply' && journal.phase === 'prepared') {
        if (revision !== journal.projectRevision) {
          throw new Error('Prepared transaction advanced its project revision')
        }
        await this.rollbackApplyJournal(journal, project)
      } else if (journal.operation === 'apply' && journal.phase === 'committed') {
        await this.finalizeApplyJournal(journal, project)
      } else if (journal.operation === 'undo' && journal.phase === 'restoring') {
        await this.completeUndoJournal(journal, project)
      } else {
        throw new Error('Unsupported project transaction journal state')
      }
    }
    return journals.length
  }

  async undo(
    projectId: string,
    snapshotId: string,
    project: LatexProject,
  ): Promise<{ snapshotId: string; restored: boolean }> {
    validateProjectId(projectId)
    validateSnapshotId(snapshotId)
    return this.transactionState.withProjectLock(projectId, async () => {
      await this.recoverLocked(projectId, project)
      return this.undoLocked(projectId, snapshotId, project)
    })
  }

  private async undoLocked(
    projectId: string,
    snapshotId: string,
    project: LatexProject,
  ): Promise<{ snapshotId: string; restored: boolean }> {
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
      await mkdir(this.undoRoot, { recursive: true })
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
    const snapshotHashes = await this.snapshots.getFileHashes(projectId, snapshotId)
    const expectedCurrentHashes = new Map(Object.entries(record.expectedCurrentHashes))
    if (
      snapshotHashes.size !== expectedCurrentHashes.size ||
      [...snapshotHashes.keys()].some((path) => !expectedCurrentHashes.has(path))
    ) {
      throw new Error('Invalid rollback record paths')
    }
    const policy = await ProjectPathPolicy.open(project.rootPath)
    await assertExpectedCurrentHashes(project, policy, expectedCurrentHashes)
    const revision = await this.transactionState.readRevision(projectId)
    const journal: ProjectTransactionJournal = {
      schemaVersion: 1,
      id: `undo-${snapshotId}`,
      operation: 'undo',
      phase: 'restoring',
      projectId,
      projectRevision: revision,
      nextProjectRevision: revision + 1,
      snapshotId,
      previousRollback: snapshotId,
      files: [...snapshotHashes].map(([path, beforeSha256]) => ({
        path,
        beforeSha256,
        afterSha256: expectedCurrentHashes.get(path)!,
      })),
    }
    await this.transactionState.writeJournal(journal)
    await this.completeUndoJournal(journal, project)
    return { snapshotId, restored: true }
  }

  private async finalizeApplyJournal(
    journal: ProjectTransactionJournal,
    project: LatexProject,
  ): Promise<void> {
    await this.validateJournalSnapshot(journal)
    const policy = await ProjectPathPolicy.open(project.rootPath)
    const expectedCurrentHashes = new Map(
      journal.files.map((file) => [file.path, file.afterSha256]),
    )
    await assertExpectedCurrentHashes(project, policy, expectedCurrentHashes)
    const record = rollbackRecordFromJournal(journal)
    await mkdir(this.rollbackRoot, { recursive: true })
    try {
      await atomicWriteFile(
        this.rollbackPath(journal.snapshotId),
        Buffer.from(`${JSON.stringify(record, null, 2)}\n`),
      )
    } catch (error) {
      const persisted = await this.readRollback(journal.snapshotId).catch(() => undefined)
      if (!persisted || JSON.stringify(persisted) !== JSON.stringify(record)) throw error
    }
    try {
      await this.snapshots.setCurrentRollback(journal.projectId, journal.snapshotId)
    } catch (error) {
      if ((await this.snapshots.getCurrentRollback(journal.projectId)) !== journal.snapshotId) {
        throw error
      }
    }
    await this.transactionState.advanceRevision(
      journal.projectId,
      journal.projectRevision,
      journal.nextProjectRevision,
    )
    await this.transactionState.deleteJournal(journal.id)
  }

  private async rollbackApplyJournal(
    journal: ProjectTransactionJournal,
    project: LatexProject,
  ): Promise<void> {
    await this.validateJournalSnapshot(journal)
    const policy = await ProjectPathPolicy.open(project.rootPath)
    const paths = new Set<string>()
    const conflicts: string[] = []
    for (const file of journal.files) {
      const current = await readOptionalText(project, policy, file.path)
      const currentSha256 = current === undefined ? null : digest(current)
      if (currentSha256 === file.afterSha256) paths.add(file.path)
      else if (currentSha256 !== file.beforeSha256) conflicts.push(file.path)
    }
    const expectedCurrentHashes = new Map(
      journal.files
        .filter((file) => paths.has(file.path))
        .map((file) => [file.path, file.afterSha256]),
    )
    await this.snapshots.restore(journal.projectId, journal.snapshotId, project, {
      expectedAbsentHashes: new Map(
        journal.files
          .filter((file) => paths.has(file.path) && file.beforeSha256 === null)
          .map((file) => [file.path, file.afterSha256]),
      ),
      expectedCurrentHashes,
      paths,
    })
    if (conflicts.length > 0) {
      throw new Error('Project changed while proposal rollback was in progress')
    }
    await this.snapshots.setCurrentRollback(journal.projectId, journal.previousRollback)
    await unlink(this.rollbackPath(journal.snapshotId)).catch((error) => {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
    })
    await this.transactionState.deleteJournal(journal.id)
  }

  private async rollbackOrAggregate(
    journal: ProjectTransactionJournal,
    project: LatexProject,
    transactionError: unknown,
  ): Promise<void> {
    try {
      await this.rollbackApplyJournal(journal, project)
    } catch (rollbackError) {
      throw new AggregateError(
        [transactionError, rollbackError],
        'Proposal transaction failed and persistent rollback is incomplete',
        { cause: rollbackError },
      )
    }
  }

  private async completeUndoJournal(
    journal: ProjectTransactionJournal,
    project: LatexProject,
  ): Promise<void> {
    await this.validateJournalSnapshot(journal)
    await this.snapshots.restore(journal.projectId, journal.snapshotId, project, {
      expectedAbsentHashes: new Map(
        journal.files
          .filter((file) => file.beforeSha256 === null)
          .map((file) => [file.path, file.afterSha256]),
      ),
      expectedCurrentHashes: new Map(journal.files.map((file) => [file.path, file.afterSha256])),
    })
    await mkdir(this.undoRoot, { recursive: true })
    await atomicWriteFile(
      this.undoPath(journal.snapshotId),
      Buffer.from(
        `${JSON.stringify({
          schemaVersion: 1,
          projectId: journal.projectId,
          snapshotId: journal.snapshotId,
        })}\n`,
      ),
    )
    await this.snapshots.setCurrentRollback(journal.projectId, null)
    await this.transactionState.advanceRevision(
      journal.projectId,
      journal.projectRevision,
      journal.nextProjectRevision,
    )
    await this.transactionState.deleteJournal(journal.id)
  }

  private async validateJournalSnapshot(journal: ProjectTransactionJournal): Promise<void> {
    const snapshotHashes = await this.snapshots.getFileHashes(journal.projectId, journal.snapshotId)
    if (
      snapshotHashes.size !== journal.files.length ||
      journal.files.some(
        (file) =>
          !snapshotHashes.has(file.path) || snapshotHashes.get(file.path) !== file.beforeSha256,
      )
    ) {
      throw new Error('Project transaction journal does not match its snapshot')
    }
  }

  private async claim(proposalId: string): Promise<StoredProposal> {
    await this.ensureDirectories()
    const pending = this.pendingPath(proposalId)
    const consumed = this.consumedPath(proposalId)
    try {
      await rename(pending, consumed)
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT' && (await exists(consumed))) {
        throw new Error('Proposal was already consumed', { cause: error })
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
  projectRevision: number,
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
    projectRevision,
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
  if (
    typeof item.projectRevision !== 'number' ||
    !Number.isSafeInteger(item.projectRevision) ||
    item.projectRevision < 0
  ) {
    throw new Error('Invalid stored proposal revision')
  }
  return validateProposal(
    value as EditProposal,
    maxFileBytes,
    maxFiles,
    Number.MIN_SAFE_INTEGER,
    item.projectRevision,
  )
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

function rollbackRecordFromJournal(journal: ProjectTransactionJournal): RollbackRecord {
  return {
    schemaVersion: 1,
    snapshotId: journal.snapshotId,
    projectId: journal.projectId,
    expectedAbsentHashes: Object.fromEntries(
      journal.files
        .filter((file) => file.beforeSha256 === null)
        .map((file) => [file.path, file.afterSha256]),
    ),
    expectedCurrentHashes: Object.fromEntries(
      journal.files.map((file) => [file.path, file.afterSha256]),
    ),
  }
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
