import { createHash, randomBytes } from 'node:crypto'
import { constants, watch as watchFs } from 'node:fs'
import { lstat, mkdir, open, rm } from 'node:fs/promises'
import { join } from 'node:path'
import {
  commitCompileGeneration,
  CompileQueue,
  compileIsolated,
  parseSyncTeX,
  parseTectonicDiagnostics,
  type CompileIsolatedResult,
  type StagedCompileResult,
  type SyncTeXIndex,
} from '@wiswork/latex-compiler'
import {
  openLatexProject,
  ProposalStore,
  SnapshotStore,
  type EditProposal,
  type LatexProject,
} from '@wiswork/latex-project'
import type { CompileResultDto, LatexBufferDto, LatexSaveDto } from '../shared/ipc.js'

export class MainFileRenameError extends Error {}

export class UnsavedBuffersError extends Error {
  constructor() {
    super('Project has unsaved LaTeX changes')
    this.name = 'UnsavedBuffersError'
  }
}

export interface WatcherLike {
  close(): void
}

export interface ProjectSessionRegistryOptions {
  watch?: (root: string, onChange: (relativePath: string) => void) => WatcherLike
  compiler?: typeof compileIsolated
  commitGeneration?: typeof commitCompileGeneration
  maxCompileResults?: number
  cleanupStaging?: (path: string) => Promise<void>
  compilerRuntime?: {
    tectonicPath: string
    userDataPath: string
  }
  onExternalChange?: (webContentsId: number, buffer: LatexBufferDto) => void
}

interface BufferState {
  path: string
  text: string
  baselineText: string
  baselineSha256: string
  dirty: boolean
  conflict: { diskText: string | null; diskSha256: string | null } | null
  version: number
  lastSaveRevision: number
  pendingSaveSha256?: string
}

interface ActiveCompile {
  revision: number
  token: string
  phase: 'pending' | 'running' | 'publishing'
  promise: Promise<CompileResultDto>
}

export class ProjectSession {
  readonly projectId: string
  readonly webContentsId: number
  readonly project: LatexProject
  private readonly watcher: WatcherLike
  private readonly buffers = new Map<string, BufferState>()
  private readonly compileCleanups = new Set<() => void>()
  private readonly downloadCleanups = new Set<() => void>()
  private readonly compiler: typeof compileIsolated
  private readonly commitGeneration: typeof commitCompileGeneration
  private readonly compilerRuntime?: ProjectSessionRegistryOptions['compilerRuntime']
  private readonly onExternalChange?: ProjectSessionRegistryOptions['onExternalChange']
  private readonly maxCompileResults: number
  private readonly cleanupStaging: (path: string) => Promise<void>
  private readonly proposalStore?: ProposalStore
  private readonly proposals = new Map<string, EditProposal>()
  private activeCompile: ActiveCompile | undefined
  private readonly cancelledCompileTokens = new Set<string>()
  private readonly compileQueue = new CompileQueue<StagedCompileResult>()
  private readonly compileResults = new Map<
    number,
    CompileResultDto & {
      pdfPath: string | null
      synctexPath: string | null
      syncTex?: SyncTeXIndex
    }
  >()
  private disposed = false

  constructor(
    webContentsId: number,
    project: LatexProject,
    watcherFactory: NonNullable<ProjectSessionRegistryOptions['watch']>,
    options: ProjectSessionRegistryOptions,
  ) {
    this.projectId = randomBytes(16).toString('hex')
    this.webContentsId = webContentsId
    this.project = project
    this.compiler = options.compiler ?? compileIsolated
    this.commitGeneration = options.commitGeneration ?? commitCompileGeneration
    this.compilerRuntime = options.compilerRuntime
    this.onExternalChange = options.onExternalChange
    this.maxCompileResults = options.maxCompileResults ?? 3
    this.cleanupStaging =
      options.cleanupStaging ?? ((path) => rm(path, { recursive: true, force: true }))
    if (
      !Number.isSafeInteger(this.maxCompileResults) ||
      this.maxCompileResults < 1 ||
      this.maxCompileResults > 20
    ) {
      throw new Error('maxCompileResults must be between 1 and 20')
    }
    if (this.compilerRuntime) {
      const cacheRoot = join(this.compilerRuntime.userDataPath, 'latex', 'project-state')
      this.proposalStore = new ProposalStore(cacheRoot, new SnapshotStore(cacheRoot))
    }
    this.watcher = watcherFactory(project.rootPath, (path) => {
      void this.handleExternalChange(path).catch(() => undefined)
    })
  }

  get mainFile(): string | undefined {
    return this.project.mainFile
  }

  async listTextFiles(): Promise<string[]> {
    this.assertActive()
    return this.project.listTextFiles()
  }

  async readText(path: string): Promise<LatexBufferDto> {
    this.assertActive()
    const existing = this.buffers.get(path)
    if (existing) return dto(existing)
    const text = await this.project.readText(path)
    const state: BufferState = {
      path,
      text,
      baselineText: text,
      baselineSha256: digest(text),
      dirty: false,
      conflict: null,
      version: 0,
      lastSaveRevision: -1,
    }
    this.buffers.set(path, state)
    return dto(state)
  }

  getBuffer(path: string): LatexBufferDto | undefined {
    const state = this.buffers.get(path)
    return state ? dto(state) : undefined
  }

  updateBuffer(path: string, text: string): LatexBufferDto {
    this.assertActive()
    const state = this.buffers.get(path)
    if (!state) throw new Error(`File must be read before editing: ${path}`)
    state.text = text
    state.version += 1
    state.dirty = text !== state.baselineText
    return dto(state)
  }

  async saveText(
    path: string,
    requestedText?: string,
    editRevision?: number,
  ): Promise<LatexSaveDto> {
    this.assertActive()
    const state = this.buffers.get(path)
    if (!state) throw new Error(`File must be read before saving: ${path}`)
    if (editRevision !== undefined) {
      if (!Number.isSafeInteger(editRevision) || editRevision < state.lastSaveRevision) {
        throw new Error('Stale renderer save revision')
      }
      state.lastSaveRevision = editRevision
    }
    if (requestedText !== undefined) this.updateBuffer(path, requestedText)
    if (state.conflict) throw new Error(`Project file changed externally: ${path}`)
    if (!state.dirty) {
      return { savedText: state.text, diskSha256: state.baselineSha256, buffer: dto(state) }
    }
    const snapshot = {
      text: state.text,
      version: state.version,
      baseline: state.baselineSha256,
    }
    const pendingHash = digest(snapshot.text)
    state.pendingSaveSha256 = pendingHash
    let saved
    try {
      saved = await this.project.saveText(path, snapshot.text, {
        expectedSha256: snapshot.baseline,
      })
    } catch (error) {
      if (state.pendingSaveSha256 === pendingHash) delete state.pendingSaveSha256
      throw error
    }
    state.baselineText = snapshot.text
    state.baselineSha256 = saved.sha256
    state.dirty =
      state.conflict !== null || state.version !== snapshot.version || state.text !== snapshot.text
    delete state.pendingSaveSha256
    return { savedText: snapshot.text, diskSha256: saved.sha256, buffer: dto(state) }
  }

  async createText(path: string, text: string): Promise<LatexBufferDto> {
    this.assertActive()
    const saved = await this.project.saveText(path, text, { expectedSha256: null })
    const state: BufferState = {
      path: saved.path,
      text,
      baselineText: text,
      baselineSha256: saved.sha256,
      dirty: false,
      conflict: null,
      version: 0,
      lastSaveRevision: -1,
    }
    this.buffers.set(saved.path, state)
    return dto(state)
  }

  async renameText(from: string, to: string): Promise<void> {
    this.assertActive()
    if (from === this.mainFile)
      throw new MainFileRenameError('Cannot rename the configured main file')
    const text = await this.project.readText(from)
    const hash = digest(text)
    await this.project.saveText(to, text, { expectedSha256: null })
    try {
      await this.project.deleteText(from, {
        expectedSha256: hash,
        transactionId: randomBytes(12).toString('hex'),
      })
    } catch (error) {
      await this.project
        .deleteText(to, {
          expectedSha256: hash,
          transactionId: randomBytes(12).toString('hex'),
        })
        .catch(() => undefined)
      throw error
    }
    const state = this.buffers.get(from)
    this.buffers.delete(from)
    if (state) this.buffers.set(to, { ...state, path: to })
  }

  async handleExternalChange(path: string): Promise<void> {
    if (this.disposed || !path || path.startsWith('.wiswork-delete-')) return
    const state = this.buffers.get(path)
    if (!state) return
    let diskText: string | null
    try {
      diskText = await this.project.readText(path)
    } catch {
      diskText = null
    }
    const diskHash = diskText === null ? null : digest(diskText)
    if (diskHash === state.baselineSha256) return
    if (diskText !== null && diskHash !== null && diskHash === state.pendingSaveSha256) {
      state.baselineText = diskText
      state.baselineSha256 = diskHash
      state.dirty = state.text !== diskText
      state.conflict = null
      this.onExternalChange?.(this.webContentsId, {
        ...dto(state),
        text: diskText,
        diskText,
        dirty: false,
        conflict: null,
      })
      return
    }
    if (state.dirty) {
      state.conflict = { diskText, diskSha256: diskHash }
    } else if (diskText === null) {
      state.conflict = { diskText: null, diskSha256: null }
    } else {
      state.text = diskText
      state.baselineText = diskText
      state.baselineSha256 = diskHash!
      state.dirty = false
      state.conflict = null
    }
    this.onExternalChange?.(this.webContentsId, dto(state))
  }

  isDirty(): boolean {
    return [...this.buffers.values()].some((state) => state.dirty)
  }

  async saveAll(): Promise<void> {
    for (const state of this.buffers.values()) if (state.dirty) await this.saveText(state.path)
  }

  async discardAll(): Promise<void> {
    for (const [path, state] of this.buffers) {
      let diskText: string
      try {
        diskText = await this.project.readText(path)
      } catch {
        this.buffers.delete(path)
        continue
      }
      state.text = diskText
      state.baselineText = diskText
      state.baselineSha256 = digest(diskText)
      state.dirty = false
      state.conflict = null
      state.version += 1
      delete state.pendingSaveSha256
    }
  }

  trackCompile(cancel: () => void): () => void {
    this.compileCleanups.add(cancel)
    return () => this.compileCleanups.delete(cancel)
  }

  trackDownload(cancel: () => void): () => void {
    this.downloadCleanups.add(cancel)
    return () => this.downloadCleanups.delete(cancel)
  }

  private assertAllBuffersPersisted(): void {
    if ([...this.buffers.values()].some((buffer) => buffer.dirty || buffer.conflict)) {
      throw new UnsavedBuffersError()
    }
  }

  async compile(revision: number, mainFile: string): Promise<CompileResultDto> {
    this.assertActive()
    this.assertAllBuffersPersisted()
    if (!this.compilerRuntime) throw new Error('LaTeX compiler runtime is not configured')
    if (this.activeCompile?.revision === revision) return this.activeCompile.promise
    const token = randomBytes(16).toString('hex')
    const cacheDirectory = join(
      this.compilerRuntime.userDataPath,
      'latex',
      'compile-cache',
      this.projectId,
    )
    const temporaryRoot = join(this.compilerRuntime.userDataPath, 'latex', 'compile-temp')
    let staged: StagedCompileResult | undefined
    let committed: CompileIsolatedResult | undefined
    const promise = this.compileQueue
      .request({
        projectId: this.projectId,
        revision: token,
        run: async ({ signal }) => {
          if (this.cancelledCompileTokens.has(token)) throw new Error('Compile cancelled')
          this.assertAllBuffersPersisted()
          if (this.activeCompile?.token === token) this.activeCompile.phase = 'running'
          await Promise.all([
            mkdir(cacheDirectory, { recursive: true }),
            mkdir(temporaryRoot, { recursive: true }),
          ])
          if (this.cancelledCompileTokens.has(token)) throw new Error('Compile cancelled')
          this.assertAllBuffersPersisted()
          staged = await this.compiler({
            projectDirectory: this.project.rootPath,
            temporaryRoot,
            cacheDirectory,
            executable: this.compilerRuntime!.tectonicPath,
            bundlePath: join(this.compilerRuntime!.userDataPath, 'latex', 'tectonic-bundle'),
            mainFile,
            signal,
          })
          return staged
        },
        publish: async (value) => {
          if (this.cancelledCompileTokens.has(token)) throw new Error('Compile cancelled')
          if (this.activeCompile?.token !== token) throw new Error('Stale compile result')
          this.activeCompile.phase = 'publishing'
          committed = await this.commitGeneration(value, cacheDirectory, {
            maxGenerations: this.maxCompileResults,
          })
        },
      })
      .then(async () => {
        if (!committed || this.activeCompile?.token !== token)
          throw new Error('Stale compile result')
        const result = committed
        let syncTex: SyncTeXIndex | undefined
        if (result.synctexPath) {
          try {
            syncTex = parseSyncTeX(
              await readBoundedArtifact(result.synctexPath, 32 * 1024 * 1024),
              result.synctexInputRoot,
            )
          } catch {
            syncTex = undefined
          }
        }
        const value = {
          revision,
          pdfUrl: result.pdfPath ? `wiswork-latex-pdf://${this.projectId}/${revision}` : null,
          pdfPath: result.pdfPath,
          synctexPath: result.synctexPath,
          syncTex,
          diagnostics: parseTectonicDiagnostics(result.log),
          log: result.log,
        }
        this.compileResults.delete(revision)
        this.compileResults.set(revision, value)
        while (this.compileResults.size > this.maxCompileResults) {
          const oldest = this.compileResults.keys().next().value as number | undefined
          if (oldest === undefined) break
          this.compileResults.delete(oldest)
        }
        return value
      })
      .finally(async () => {
        try {
          if (staged) await this.cleanupStaging(staged.stagingDirectory).catch(() => undefined)
        } finally {
          this.cancelledCompileTokens.delete(token)
          if (this.activeCompile?.token === token) this.activeCompile = undefined
        }
      })
    this.activeCompile = { revision, token, phase: 'pending', promise }
    return promise
  }

  cancelCompile(): boolean {
    if (!this.activeCompile) return false
    if (this.activeCompile.phase === 'publishing') return false
    const token = this.activeCompile.token
    this.cancelledCompileTokens.add(token)
    this.compileQueue.cancel(this.projectId)
    this.activeCompile = undefined
    return true
  }

  pdfPath(revision: number): string | undefined {
    return this.compileResults.get(revision)?.pdfPath ?? undefined
  }

  syncTexForward(revision: number, path: string, line: number) {
    return this.compileResults.get(revision)?.syncTex?.forward(path, line) ?? null
  }

  syncTexReverse(revision: number, page: number, x: number, y: number) {
    return this.compileResults.get(revision)?.syncTex?.inverse(page, x, y) ?? null
  }

  async registerProposal(proposal: Omit<EditProposal, 'projectId'>): Promise<void> {
    if (!this.proposalStore) throw new Error('Proposal store is not configured')
    const owned = { ...proposal, projectId: this.projectId }
    await this.proposalStore.create(owned)
    this.proposals.set(owned.id, structuredClone(owned))
  }

  getProposal(id: string): EditProposal | undefined {
    const proposal = this.proposals.get(id)
    return proposal ? structuredClone(proposal) : undefined
  }

  async applyProposal(id: string) {
    if (!this.proposalStore || !this.proposals.has(id)) throw new Error('Proposal not found')
    return this.proposalStore.apply(id, this.projectId, this.project)
  }

  async undoProposal(snapshotId: string) {
    if (!this.proposalStore) throw new Error('Proposal store is not configured')
    return this.proposalStore.undo(this.projectId, snapshotId, this.project)
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.watcher.close()
    this.cancelCompile()
    for (const cancel of [...this.compileCleanups, ...this.downloadCleanups]) cancel()
    this.compileCleanups.clear()
    this.downloadCleanups.clear()
    this.buffers.clear()
  }

  private assertActive(): void {
    if (this.disposed) throw new Error('Project session is closed')
  }
}

export class ProjectSessionRegistry {
  private readonly byWebContents = new Map<number, ProjectSession>()
  private readonly byProjectId = new Map<string, ProjectSession>()
  private readonly options: ProjectSessionRegistryOptions
  private readonly watcherFactory: NonNullable<ProjectSessionRegistryOptions['watch']>

  constructor(options: ProjectSessionRegistryOptions = {}) {
    this.options = options
    this.watcherFactory = options.watch ?? defaultWatch
  }

  async attach(webContentsId: number, projectRoot: string): Promise<ProjectSession> {
    if (!Number.isSafeInteger(webContentsId) || webContentsId <= 0) {
      throw new Error('Invalid WebContents ID')
    }
    if (this.byWebContents.has(webContentsId)) {
      throw new Error('WebContents already owns a LaTeX project session')
    }
    const project = await openLatexProject(projectRoot)
    const session = new ProjectSession(webContentsId, project, this.watcherFactory, this.options)
    this.byWebContents.set(webContentsId, session)
    this.byProjectId.set(session.projectId, session)
    return session
  }

  getByWebContents(webContentsId: number): ProjectSession | undefined {
    return this.byWebContents.get(webContentsId)
  }

  getOwned(webContentsId: number, projectId: string): ProjectSession | undefined {
    const session = this.byProjectId.get(projectId)
    return session?.webContentsId === webContentsId ? session : undefined
  }

  destroy(webContentsId: number): void {
    const session = this.byWebContents.get(webContentsId)
    if (!session) return
    this.byWebContents.delete(webContentsId)
    this.byProjectId.delete(session.projectId)
    session.dispose()
  }

  resolvePdf(projectId: string, revision: number): string | undefined {
    return this.byProjectId.get(projectId)?.pdfPath(revision)
  }

  disposeAll(): void {
    for (const id of [...this.byWebContents.keys()]) this.destroy(id)
  }
}

function defaultWatch(root: string, onChange: (relativePath: string) => void): WatcherLike {
  const watcher = watchFs(root, { recursive: true }, (_event, filename) => {
    if (!filename) return
    onChange(filename.toString().replaceAll('\\', '/'))
  })
  return watcher
}

function digest(text: string): string {
  return createHash('sha256').update(text, 'utf8').digest('hex')
}

async function readBoundedArtifact(path: string, maxBytes: number): Promise<Uint8Array> {
  const before = await lstat(path)
  if (!before.isFile() || before.isSymbolicLink() || before.size > maxBytes) {
    throw new Error('SyncTeX artifact exceeds host size limit')
  }
  const noFollow = 'O_NOFOLLOW' in constants ? constants.O_NOFOLLOW : 0
  const handle = await open(path, constants.O_RDONLY | noFollow)
  try {
    const opened = await handle.stat()
    if (opened.dev !== before.dev || opened.ino !== before.ino || opened.size !== before.size) {
      throw new Error('SyncTeX artifact changed before reading')
    }
    const bytes = Buffer.allocUnsafe(opened.size)
    let offset = 0
    while (offset < bytes.length) {
      const { bytesRead } = await handle.read(bytes, offset, bytes.length - offset, offset)
      if (bytesRead === 0) break
      offset += bytesRead
    }
    if (offset !== bytes.length || (await handle.stat()).size !== opened.size) {
      throw new Error('SyncTeX artifact changed while reading')
    }
    return bytes
  } finally {
    await handle.close()
  }
}

function dto(state: BufferState): LatexBufferDto {
  return {
    path: state.path,
    text: state.text,
    diskText: state.baselineText,
    diskSha256: state.baselineSha256,
    dirty: state.dirty,
    conflict: state.conflict ? { ...state.conflict } : null,
  }
}
