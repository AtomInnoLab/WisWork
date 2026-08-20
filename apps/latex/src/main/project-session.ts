import { createHash, randomBytes } from 'node:crypto'
import { constants, watch as watchFs } from 'node:fs'
import { lstat, mkdir, open, rm } from 'node:fs/promises'
import { join } from 'node:path'
import {
  commitCompileGeneration,
  CompileQueue,
  compileIsolated,
  isRemoteIndexedBundleUrl,
  LatexCompilerError,
  loadCurrentCompileGeneration,
  parseSyncTeX,
  parseTectonicDiagnostics,
  TectonicRunError,
  type CompileIsolatedResult,
  type PublishedCompileGeneration,
  type StagedCompileResult,
  type SyncTeXIndex,
  type TectonicBundleAsset,
} from '@wiswork/latex-compiler'
import {
  openLatexProject,
  ProjectPathPolicy,
  ProposalStore,
  SnapshotStore,
  type EditProposal,
  type LatexProject,
} from '@wiswork/latex-project'
import { ProjectStore } from '@wiswork/project-store'
import type {
  CompileResultDto,
  LatexBufferDto,
  LatexProposalDto,
  LatexSaveDto,
  ProposalVerificationDto,
} from '../shared/ipc.js'
import { isAiSensitivePath } from '../shared/ai-path-policy.js'
import {
  MAX_FORMAL_COMPILE_DIAGNOSTICS,
  normalizeProposalDiagnostics,
} from '../shared/proposal-verification.js'

export class MainFileRenameError extends Error {}

export class UnsavedBuffersError extends Error {
  constructor() {
    super('Project has unsaved LaTeX changes')
    this.name = 'UnsavedBuffersError'
  }
}

type ProposalVerificationRejectionReason =
  'not-found' | 'expired' | 'baseline' | 'configuration' | 'safety'

export class ProposalVerificationRejectedError extends Error {
  constructor(
    readonly reason: ProposalVerificationRejectionReason,
    message: string,
  ) {
    super(message)
    this.name = 'ProposalVerificationRejectedError'
  }
}

const MAX_PROPOSAL_VERIFICATION_LOG_BYTES = 16_000
const CONFIRMED_INTERNAL_COMPILE = Symbol('confirmed-internal-compile')

export interface WatcherLike {
  close(): void
}

export interface ProjectSessionRegistryOptions {
  watch?: (root: string, onChange: (relativePath: string) => void) => WatcherLike
  compiler?: typeof compileIsolated
  commitGeneration?: typeof commitCompileGeneration
  loadCurrentGeneration?: typeof loadCurrentCompileGeneration
  maxCompileResults?: number
  cleanupStaging?: (path: string) => Promise<void>
  compilerRuntime?: {
    tectonicPath: string
    userDataPath: string
    bundleAsset?: TectonicBundleAsset
  }
  onExternalChange?: (webContentsId: number, buffer: LatexBufferDto) => void
  acquireRendererFreeze?: (webContentsId: number) => Promise<() => void>
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
  kind: 'compile' | 'verification'
  revision: number | null
  token: string
  phase: 'pending' | 'running' | 'publishing'
  promise: Promise<unknown>
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
  private readonly loadCurrentGeneration: typeof loadCurrentCompileGeneration
  private readonly acquireRendererFreeze: () => Promise<() => void>
  private readonly proposalStore?: ProposalStore
  private readonly projectStore?: ProjectStore
  private readonly remoteBundleUrl?: string
  private readonly fallbackBundlePath?: string
  private readonly proposals = new Map<string, EditProposal>()
  private readonly proposalReviews = new Map<string, LatexProposalDto>()
  private readonly saveQueues = new Map<string, Promise<unknown>>()
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
  private confirmedEditRevision = 0
  private confirmedMutationInProgress = false
  private activeRendererMutations = 0
  private rendererMutationsSettled: Promise<void> = Promise.resolve()
  private resolveRendererMutationsSettled: (() => void) | undefined

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
    this.loadCurrentGeneration = options.loadCurrentGeneration ?? loadCurrentCompileGeneration
    this.acquireRendererFreeze = options.acquireRendererFreeze
      ? () => options.acquireRendererFreeze!(this.webContentsId)
      : async () => {
          throw new Error('LaTeX renderer freeze is not configured')
        }
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
      this.projectStore = new ProjectStore(this.compilerRuntime.userDataPath)
      if (
        this.compilerRuntime.bundleAsset &&
        isRemoteIndexedBundleUrl(this.compilerRuntime.bundleAsset.url)
      ) {
        this.remoteBundleUrl = this.compilerRuntime.bundleAsset.url
      } else if (!this.compilerRuntime.bundleAsset) {
        this.fallbackBundlePath = join(
          this.compilerRuntime.userDataPath,
          'latex',
          'tectonic-bundle',
        )
      }
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
    const release = this.acquireRendererMutation()
    try {
      return this.updateBufferState(path, text)
    } finally {
      release()
    }
  }

  private updateBufferState(path: string, text: string): LatexBufferDto {
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
    const release = this.acquireRendererMutation()
    try {
      const state = this.buffers.get(path)
      if (!state) throw new Error(`File must be read before saving: ${path}`)
      if (editRevision !== undefined) {
        if (!Number.isSafeInteger(editRevision) || editRevision < state.lastSaveRevision) {
          throw new Error('Stale renderer save revision')
        }
        state.lastSaveRevision = editRevision
      }
      if (requestedText !== undefined) this.updateBufferState(path, requestedText)
      return await this.enqueueSave(path, () => this.saveCurrentText(path))
    } finally {
      release()
    }
  }

  private enqueueSave<T>(path: string, operation: () => Promise<T>): Promise<T> {
    const previous = this.saveQueues.get(path) ?? Promise.resolve()
    const next = previous.catch(() => undefined).then(operation)
    this.saveQueues.set(path, next)
    const cleanup = () => {
      if (this.saveQueues.get(path) === next) this.saveQueues.delete(path)
    }
    void next.then(cleanup, cleanup)
    return next
  }

  async settleSaves(): Promise<void> {
    while (this.saveQueues.size > 0) {
      await Promise.allSettled([...this.saveQueues.values()])
    }
  }

  private async saveCurrentText(path: string): Promise<LatexSaveDto> {
    const state = this.buffers.get(path)
    if (!state) throw new Error(`File must be read before saving: ${path}`)
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
    const release = this.acquireRendererMutation()
    try {
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
    } finally {
      release()
    }
  }

  async renameText(from: string, to: string): Promise<void> {
    this.assertActive()
    const release = this.acquireRendererMutation()
    try {
      if (from === this.mainFile)
        throw new MainFileRenameError('Cannot rename the configured main file')
      const state = this.buffers.get(from)
      if (state?.dirty || state?.conflict) throw new UnsavedBuffersError()
      const text = await this.project.readText(from)
      const hash = digest(text)
      const saved = await this.project.saveText(to, text, { expectedSha256: null })
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
      this.buffers.delete(from)
      if (state) {
        this.buffers.set(to, {
          ...state,
          path: to,
          text,
          baselineText: text,
          baselineSha256: saved.sha256,
          dirty: false,
          conflict: null,
        })
      }
    } finally {
      release()
    }
  }

  async deleteText(path: string): Promise<void> {
    this.assertActive()
    const release = this.acquireRendererMutation()
    try {
      if (path === this.mainFile)
        throw new MainFileRenameError('Cannot delete the configured main file')
      const state = this.buffers.get(path)
      if (state?.dirty || state?.conflict) throw new UnsavedBuffersError()
      const text = await this.project.readText(path)
      await this.project.deleteText(path, {
        expectedSha256: digest(text),
        transactionId: randomBytes(12).toString('hex'),
      })
      this.buffers.delete(path)
    } finally {
      release()
    }
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
    if (this.disposed || this.buffers.get(path) !== state) return
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
    await this.settleSaves()
  }

  async discardAll(): Promise<void> {
    await this.settleSaves()
    for (const [path, state] of this.buffers) {
      const version = state.version
      let diskText: string
      try {
        diskText = await this.project.readText(path)
      } catch {
        if (this.buffers.get(path) === state && state.version === version) {
          this.buffers.delete(path)
        }
        continue
      }
      if (this.buffers.get(path) !== state || state.version !== version) continue
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

  private compileCacheDirectory(): string {
    if (!this.compilerRuntime) throw new Error('LaTeX compiler runtime is not configured')
    const projectKey = createHash('sha256').update(this.project.rootPath, 'utf8').digest('hex')
    return join(this.compilerRuntime.userDataPath, 'latex', 'compile-cache', projectKey)
  }

  async restoreLatestCompile(): Promise<void> {
    if (!this.compilerRuntime) return
    let restored: PublishedCompileGeneration | null
    try {
      restored = await this.loadCurrentGeneration(this.compileCacheDirectory())
    } catch {
      return
    }
    if (!restored?.pdfPath) return
    const value: CompileResultDto = {
      revision: 0,
      pdfUrl: `wiswork-latex-pdf://${this.projectId}/0`,
      diagnostics: normalizeProposalDiagnostics(
        parseTectonicDiagnostics(restored.log),
        MAX_FORMAL_COMPILE_DIAGNOSTICS,
      ),
      log: restored.log,
    }
    this.compileResults.set(0, {
      ...value,
      pdfPath: restored.pdfPath,
      synctexPath: restored.synctexPath,
    })
  }

  latestCompile(): CompileResultDto | null {
    const latest = [...this.compileResults.values()].at(-1)
    if (!latest) return null
    return {
      revision: latest.revision,
      pdfUrl: latest.pdfUrl,
      diagnostics: structuredClone(latest.diagnostics),
      log: latest.log,
    }
  }

  async compile(
    revision: number,
    mainFile: string,
    authorization?: typeof CONFIRMED_INTERNAL_COMPILE,
  ): Promise<CompileResultDto> {
    this.assertActive()
    if (this.confirmedMutationInProgress && authorization !== CONFIRMED_INTERNAL_COMPILE) {
      throw new Error('Confirmed edit transaction is in progress')
    }
    this.assertAllBuffersPersisted()
    if (!this.compilerRuntime) throw new Error('LaTeX compiler runtime is not configured')
    if (this.activeCompile?.kind === 'compile' && this.activeCompile.revision === revision) {
      return this.activeCompile.promise as Promise<CompileResultDto>
    }
    const token = randomBytes(16).toString('hex')
    const cacheDirectory = this.compileCacheDirectory()
    const temporaryRoot = join(this.compilerRuntime.userDataPath, 'latex', 'compile-temp')
    const tectonicCacheDirectory = join(
      this.compilerRuntime.userDataPath,
      'latex',
      'tectonic-cache',
    )
    let staged: StagedCompileResult | undefined
    let committed: CompileIsolatedResult | undefined
    const promise = this.compileQueue
      .request({
        projectId: this.projectId,
        revision: token,
        run: async ({ signal }) => {
          if (this.cancelledCompileTokens.has(token)) throw new Error('Compile cancelled')
          this.assertAllBuffersPersisted()
          const bundlePath = await this.ensureBundle()
          if (this.cancelledCompileTokens.has(token)) throw new Error('Compile cancelled')
          if (this.activeCompile?.token === token) this.activeCompile.phase = 'running'
          await Promise.all([
            mkdir(cacheDirectory, { recursive: true }),
            mkdir(temporaryRoot, { recursive: true }),
            mkdir(tectonicCacheDirectory, { recursive: true }),
          ])
          if (this.cancelledCompileTokens.has(token)) throw new Error('Compile cancelled')
          this.assertAllBuffersPersisted()
          staged = await this.compiler({
            projectDirectory: this.project.rootPath,
            temporaryRoot,
            cacheDirectory,
            executable: this.compilerRuntime!.tectonicPath,
            bundlePath,
            tectonicCacheDirectory,
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
        const value: CompileResultDto = {
          revision,
          pdfUrl: result.pdfPath ? `wiswork-latex-pdf://${this.projectId}/${revision}` : null,
          diagnostics: normalizeProposalDiagnostics(
            parseTectonicDiagnostics(result.log),
            MAX_FORMAL_COMPILE_DIAGNOSTICS,
          ),
          log: result.log,
        }
        this.compileResults.delete(revision)
        this.compileResults.set(revision, {
          ...value,
          pdfPath: result.pdfPath,
          synctexPath: result.synctexPath,
          syncTex,
        })
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
    this.activeCompile = { kind: 'compile', revision, token, phase: 'pending', promise }
    return promise
  }

  cancelCompile(): boolean {
    if (this.confirmedMutationInProgress) return false
    return this.cancelCompileInternal()
  }

  private cancelCompileInternal(): boolean {
    if (!this.activeCompile) return false
    if (this.activeCompile.phase === 'publishing') return false
    const token = this.activeCompile.token
    this.cancelledCompileTokens.add(token)
    this.compileQueue.cancel(this.projectId)
    this.activeCompile = undefined
    return true
  }

  getBundleStatus() {
    if (this.remoteBundleUrl) return { state: 'remote' as const }
    return { state: 'error' as const, code: 'BUNDLE_NOT_CONFIGURED' }
  }

  private async ensureBundle(): Promise<string> {
    if (this.remoteBundleUrl) return this.remoteBundleUrl
    if (this.fallbackBundlePath) return this.fallbackBundlePath
    throw new Error('TeX bundle is not configured')
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

  resolveDirectoryChat() {
    if (!this.projectStore) throw new Error('Project store is not configured')
    return this.projectStore.resolveChatForDirectory(this.project.rootPath)
  }

  appendDirectoryChat(
    storeProjectId: string,
    chatId: string,
    role: 'user' | 'assistant',
    text: string,
  ) {
    if (!this.projectStore) throw new Error('Project store is not configured')
    const owned = this.resolveDirectoryChat()
    if (owned.projectId !== storeProjectId || owned.chatId !== chatId)
      throw new Error('Chat does not belong to project session')
    this.projectStore.appendChatMessage(storeProjectId, chatId, { role, text })
  }

  loadDirectoryChat(storeProjectId: string, chatId: string, limit: number) {
    if (!this.projectStore) throw new Error('Project store is not configured')
    const owned = this.resolveDirectoryChat()
    if (owned.projectId !== storeProjectId || owned.chatId !== chatId)
      throw new Error('Chat does not belong to project session')
    return this.projectStore.loadChat(storeProjectId, chatId, limit)
  }

  async listProjectFilesForAi() {
    const files = await this.listTextFiles()
    const allowed = files.filter((path) => !isAiSensitivePath(path))
    return { files: allowed.slice(0, 200), truncated: allowed.length > 200 }
  }

  async readProjectTextForAi(path: string, offset: number, maxChars: number) {
    if (!Number.isSafeInteger(offset) || offset < 0 || offset > 10_000_000)
      throw new Error('Invalid read offset')
    if (!Number.isSafeInteger(maxChars) || maxChars < 1 || maxChars > 24_000)
      throw new Error('Invalid read size')
    if (isAiSensitivePath(path)) throw new Error('Sensitive project files are not AI-readable')
    const text = await this.project.readText(path)
    if (text.includes('\0')) throw new Error('Binary project files are not AI-readable')
    if (Buffer.byteLength(text, 'utf8') > 256 * 1024)
      throw new Error('AI read file exceeds size limit')
    return {
      path,
      offset,
      totalChars: text.length,
      text: text.slice(offset, offset + maxChars),
      truncated: offset + maxChars < text.length,
    }
  }

  async searchProjectTextForAi(query: string, maxResults: number) {
    if (!query || query.length > 256) throw new Error('Invalid search query')
    if (!Number.isSafeInteger(maxResults) || maxResults < 1 || maxResults > 50)
      throw new Error('Invalid search result limit')
    const files = (await this.listTextFiles())
      .filter((path) => !isAiSensitivePath(path))
      .slice(0, 200)
    const matches: Array<{ path: string; line: number; text: string }> = []
    let outputChars = 0
    for (const path of files) {
      let text: string
      try {
        text = await this.project.readText(path)
      } catch {
        continue
      }
      if (text.includes('\0') || Buffer.byteLength(text, 'utf8') > 256 * 1024) continue
      for (const [index, line] of text.split(/\r?\n/).entries()) {
        if (!line.includes(query)) continue
        const excerpt = line.slice(0, 400)
        if (outputChars + path.length + excerpt.length > 32_000) return { matches, truncated: true }
        matches.push({ path, line: index + 1, text: excerpt })
        outputChars += path.length + excerpt.length
        if (matches.length >= maxResults) return { matches, truncated: true }
      }
    }
    return { matches, truncated: false }
  }

  getCompileDiagnosticsForAi() {
    const latest = [...this.compileResults.values()].at(-1)
    return {
      revision: latest?.revision ?? null,
      diagnostics: (latest?.diagnostics ?? []).slice(0, 100),
    }
  }

  async compileForAi() {
    if (!this.mainFile) throw new Error('Main file is not configured')
    this.confirmedEditRevision = Math.max(this.confirmedEditRevision + 1, Date.now())
    const result = await this.compile(this.confirmedEditRevision, this.mainFile)
    return {
      revision: result.revision,
      pdfUrl: result.pdfUrl,
      diagnostics: result.diagnostics.slice(0, 100),
    }
  }

  async createEditProposal(
    files: Array<{ path: string; afterText: string }>,
  ): Promise<LatexProposalDto> {
    this.assertActive()
    if (!this.proposalStore) throw new Error('Proposal store is not configured')
    if (!Array.isArray(files) || files.length < 1 || files.length > 20)
      throw new Error('Proposal file count is invalid')
    const policy = await ProjectPathPolicy.open(this.project.rootPath)
    const seen = new Set<string>()
    let reviewBytes = 0
    for (const file of files) {
      if (typeof file.afterText !== 'string') throw new Error('Proposal text is invalid')
      const bytes = Buffer.byteLength(file.afterText, 'utf8')
      if (bytes > 2 * 1024 * 1024) throw new Error('Proposal text exceeds file size limit')
      reviewBytes += bytes
      if (reviewBytes > 4 * 1024 * 1024) throw new Error('Proposal review exceeds total size limit')
    }
    const reviewed: LatexProposalDto['files'] = []
    for (const file of files) {
      const path = policy.normalize(file.path)
      if (isAiSensitivePath(path))
        throw new Error('Sensitive project files cannot be proposed by AI')
      if (seen.has(path)) throw new Error('Proposal paths must be unique')
      seen.add(path)
      let beforeText: string | null
      try {
        beforeText = await this.project.readText(path)
      } catch (error) {
        if (error instanceof Error && /does not exist/.test(error.message)) beforeText = null
        else throw error
      }
      const beforeBytes = beforeText === null ? 0 : Buffer.byteLength(beforeText, 'utf8')
      if (beforeBytes > 256 * 1024) throw new Error('Proposal baseline exceeds file size limit')
      reviewBytes += beforeBytes
      if (reviewBytes > 4 * 1024 * 1024) throw new Error('Proposal review exceeds total size limit')
      if (beforeText?.includes('\0') || file.afterText.includes('\0'))
        throw new Error('Binary proposal targets are not supported')
      reviewed.push({
        path,
        beforeText,
        beforeSha256: beforeText === null ? null : digest(beforeText),
        afterText: file.afterText,
      })
    }
    const proposal: EditProposal = {
      id: `ai-${randomBytes(16).toString('hex')}`,
      projectId: this.projectId,
      expiresAt: Date.now() + 5 * 60_000,
      files: reviewed.map(({ path, beforeSha256, afterText }) => ({
        path,
        beforeSha256,
        afterText,
      })),
    }
    await this.proposalStore.create(proposal)
    this.proposals.set(proposal.id, structuredClone(proposal))
    const dto = { ...proposal, files: reviewed }
    this.proposalReviews.set(proposal.id, structuredClone(dto))
    return dto
  }

  async registerProposal(proposal: Omit<EditProposal, 'projectId'>): Promise<void> {
    if (!this.proposalStore) throw new Error('Proposal store is not configured')
    const owned = { ...proposal, projectId: this.projectId }
    await this.proposalStore.create(owned)
    this.proposals.set(owned.id, structuredClone(owned))
  }

  getProposal(id: string): EditProposal | LatexProposalDto | undefined {
    const review = this.proposalReviews.get(id)
    if (review) return structuredClone(review)
    const proposal = this.proposals.get(id)
    return proposal ? structuredClone(proposal) : undefined
  }

  async verifyProposal(id: string): Promise<ProposalVerificationDto> {
    this.assertActive()
    return this.withConfirmedMutation(async () => {
      this.assertAllBuffersPersisted()
      const proposal = this.proposals.get(id)
      if (!proposal) {
        throw new ProposalVerificationRejectedError('not-found', 'Proposal not found')
      }
      if (proposal.expiresAt <= Date.now()) {
        throw new ProposalVerificationRejectedError('expired', 'Proposal has expired')
      }
      for (const file of proposal.files) {
        if (file.beforeSha256 === null) continue
        let text: string
        try {
          text = await this.project.readText(file.path)
        } catch {
          throw new ProposalVerificationRejectedError(
            'baseline',
            'Proposal baseline changed on disk',
          )
        }
        if (digest(text) !== file.beforeSha256) {
          throw new ProposalVerificationRejectedError(
            'baseline',
            'Proposal baseline changed on disk',
          )
        }
      }
      if (proposal.files.some((file) => file.beforeSha256 === null)) {
        return {
          proposalId: proposal.id,
          state: 'unverifiable',
          reason: 'Proposals that create new files cannot be verified in isolation',
          diagnostics: [],
          logSummary: '',
          verifiedAt: Date.now(),
        }
      }
      if (!this.compilerRuntime || !this.mainFile) {
        throw new ProposalVerificationRejectedError(
          'configuration',
          'Proposal verification is not configured',
        )
      }
      let bundlePath: string
      try {
        bundlePath = await this.ensureBundle()
      } catch {
        throw new ProposalVerificationRejectedError(
          'configuration',
          'Proposal verification is not configured',
        )
      }
      const cacheDirectory = join(
        this.compilerRuntime.userDataPath,
        'latex',
        'proposal-verification-cache',
        this.projectId,
      )
      const temporaryRoot = join(this.compilerRuntime.userDataPath, 'latex', 'compile-temp')
      const tectonicCacheDirectory = join(
        this.compilerRuntime.userDataPath,
        'latex',
        'tectonic-cache',
      )
      await Promise.all([
        mkdir(cacheDirectory, { recursive: true }),
        mkdir(temporaryRoot, { recursive: true }),
        mkdir(tectonicCacheDirectory, { recursive: true }),
      ])
      return this.runIsolatedVerification(proposal, {
        cacheDirectory,
        temporaryRoot,
        tectonicCacheDirectory,
        bundlePath,
      })
    })
  }

  private runIsolatedVerification(
    proposal: EditProposal,
    runtime: {
      cacheDirectory: string
      temporaryRoot: string
      tectonicCacheDirectory: string
      bundlePath: string
    },
  ): Promise<ProposalVerificationDto> {
    const token = randomBytes(16).toString('hex')
    let staged: StagedCompileResult | undefined
    const promise = this.compileQueue
      .request({
        projectId: this.projectId,
        revision: `proposal-verification:${proposal.id}:${token}`,
        run: async ({ signal }) => {
          if (this.cancelledCompileTokens.has(token) || this.disposed) {
            throw new Error('Proposal verification cancelled')
          }
          if (this.activeCompile?.token === token) this.activeCompile.phase = 'running'
          staged = await this.compiler({
            projectDirectory: this.project.rootPath,
            temporaryRoot: runtime.temporaryRoot,
            cacheDirectory: runtime.cacheDirectory,
            executable: this.compilerRuntime!.tectonicPath,
            bundlePath: runtime.bundlePath,
            tectonicCacheDirectory: runtime.tectonicCacheDirectory,
            mainFile: this.mainFile!,
            overlay: proposal.files.map((file) => ({ path: file.path, text: file.afterText })),
            expectedSourceHashes: Object.fromEntries(
              proposal.files.map((file) => [file.path, file.beforeSha256!]),
            ),
            signal,
          })
          return staged
        },
        publish: () => {
          if (this.cancelledCompileTokens.has(token) || this.activeCompile?.token !== token) {
            throw new Error('Proposal verification cancelled')
          }
          this.activeCompile.phase = 'publishing'
        },
      })
      .then((result): ProposalVerificationDto => ({
        proposalId: proposal.id,
        state: 'verified',
        diagnostics: normalizeProposalDiagnostics(
          parseTectonicDiagnostics(result.log, result.synctexInputRoot),
        ),
        logSummary: boundedUtf8(result.log, MAX_PROPOSAL_VERIFICATION_LOG_BYTES),
        verifiedAt: Date.now(),
      }))
      .catch((error: unknown): ProposalVerificationDto => {
        if (
          error instanceof TectonicRunError &&
          error.code === 'TECTONIC_EXIT_NONZERO' &&
          error.terminationConfirmed &&
          error.exitCode !== null
        ) {
          return {
            proposalId: proposal.id,
            state: 'failed',
            reason: 'Tectonic could not compile the isolated proposal',
            diagnostics: normalizeProposalDiagnostics(parseTectonicDiagnostics(error.log)),
            logSummary: boundedUtf8(error.log, MAX_PROPOSAL_VERIFICATION_LOG_BYTES),
            verifiedAt: Date.now(),
          }
        }
        if (error instanceof LatexCompilerError) {
          throw new ProposalVerificationRejectedError(
            'safety',
            'Proposal verification was rejected by the compiler safety policy',
          )
        }
        throw error
      })
      .finally(async () => {
        try {
          if (staged) await this.cleanupStaging(staged.stagingDirectory)
        } finally {
          this.cancelledCompileTokens.delete(token)
          if (this.activeCompile?.token === token) this.activeCompile = undefined
        }
      })
    this.activeCompile = {
      kind: 'verification',
      revision: null,
      token,
      phase: 'pending',
      promise,
    }
    return promise
  }

  async applyProposal(id: string) {
    if (!this.proposalStore || !this.proposals.has(id)) throw new Error('Proposal not found')
    return this.proposalStore.apply(id, this.projectId, this.project)
  }

  async undoProposal(snapshotId: string) {
    if (!this.proposalStore) throw new Error('Proposal store is not configured')
    return this.proposalStore.undo(this.projectId, snapshotId, this.project)
  }

  async applyConfirmedProposal(id: string) {
    this.assertActive()
    return this.withConfirmedMutation(async () => {
      this.assertAllBuffersPersisted()
      const proposal = this.proposals.get(id)
      if (!this.proposalStore || !proposal) throw new Error('Proposal not found')
      const applied = await this.proposalStore.apply(id, this.projectId, this.project)
      this.proposals.delete(id)
      this.proposalReviews.delete(id)
      await this.refreshBuffers(proposal.files.map((file) => file.path))
      return { ...applied, compile: await this.compileAfterConfirmedEdit() }
    })
  }

  async undoConfirmedProposal(snapshotId: string) {
    this.assertActive()
    return this.withConfirmedMutation(async () => {
      this.assertAllBuffersPersisted()
      if (!this.proposalStore) throw new Error('Proposal store is not configured')
      const undone = await this.proposalStore.undo(this.projectId, snapshotId, this.project)
      if (!undone.restored) return { ...undone, compile: null }
      await this.refreshBuffers([...this.buffers.keys()])
      return { ...undone, compile: await this.compileAfterConfirmedEdit() }
    })
  }

  private async refreshBuffers(paths: readonly string[]): Promise<void> {
    for (const path of paths) {
      const state = this.buffers.get(path)
      if (!state) continue
      try {
        const text = await this.project.readText(path)
        state.text = text
        state.baselineText = text
        state.baselineSha256 = digest(text)
        state.dirty = false
        state.conflict = null
        state.version += 1
        delete state.pendingSaveSha256
        this.onExternalChange?.(this.webContentsId, dto(state))
      } catch {
        this.buffers.delete(path)
      }
    }
  }

  private async compileAfterConfirmedEdit() {
    if (!this.mainFile) return { ok: false as const, error: 'Main file is not configured' }
    this.confirmedEditRevision = Math.max(this.confirmedEditRevision + 1, Date.now())
    try {
      const result = await this.compile(
        this.confirmedEditRevision,
        this.mainFile,
        CONFIRMED_INTERNAL_COMPILE,
      )
      return { ok: true as const, result }
    } catch (error) {
      return { ok: false as const, error: error instanceof Error ? error.message : String(error) }
    }
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.watcher.close()
    this.cancelCompileInternal()
    for (const cancel of [...this.compileCleanups, ...this.downloadCleanups]) cancel()
    this.compileCleanups.clear()
    this.downloadCleanups.clear()
    void this.settleSaves().finally(() => this.buffers.clear())
  }

  private assertActive(): void {
    if (this.disposed) throw new Error('Project session is closed')
  }

  private assertRendererMutationAllowed(): void {
    if (this.confirmedMutationInProgress)
      throw new Error('Confirmed edit transaction is in progress')
  }

  private acquireRendererMutation(): () => void {
    this.assertRendererMutationAllowed()
    if (this.activeRendererMutations === 0) {
      this.rendererMutationsSettled = new Promise<void>((resolve) => {
        this.resolveRendererMutationsSettled = resolve
      })
    }
    this.activeRendererMutations += 1
    let released = false
    return () => {
      if (released) return
      released = true
      this.activeRendererMutations -= 1
      if (this.activeRendererMutations === 0) {
        this.resolveRendererMutationsSettled?.()
        this.resolveRendererMutationsSettled = undefined
      }
    }
  }

  private async withConfirmedMutation<T>(operation: () => Promise<T>): Promise<T> {
    this.assertActive()
    if (this.confirmedMutationInProgress)
      throw new Error('Confirmed edit transaction is already in progress')
    const releaseFreeze = await this.acquireRendererFreeze()
    let ownsMutation = false
    try {
      this.assertActive()
      if (this.confirmedMutationInProgress)
        throw new Error('Confirmed edit transaction is already in progress')
      this.confirmedMutationInProgress = true
      ownsMutation = true
      await this.rendererMutationsSettled
      this.assertActive()
      return await operation()
    } finally {
      if (ownsMutation) this.confirmedMutationInProgress = false
      releaseFreeze()
    }
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
    await session.restoreLatestCompile()
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

function boundedUtf8(text: string, maxBytes: number): string {
  if (Buffer.byteLength(text, 'utf8') <= maxBytes) return text
  let bytes = 0
  let result = ''
  for (const character of text) {
    const characterBytes = Buffer.byteLength(character, 'utf8')
    if (bytes + characterBytes > maxBytes) break
    result += character
    bytes += characterBytes
  }
  return result
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
