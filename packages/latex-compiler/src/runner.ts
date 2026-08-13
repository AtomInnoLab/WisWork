import { createHash, randomUUID } from 'node:crypto'
import { spawn as nodeSpawn, type SpawnOptions } from 'node:child_process'
import { constants } from 'node:fs'
import { lstat, mkdir, open, readFile, readdir, realpath, rename, rm } from 'node:fs/promises'
import { basename, extname, isAbsolute, join, resolve, sep, win32 } from 'node:path'
import { LatexCompilerError, type LatexCompilerErrorCode } from './errors.js'
import { isRemoteIndexedBundleUrl } from './manifest.js'
import {
  createCompileWorkspace,
  type CompileWorkspace,
  type CompileWorkspaceLimits,
} from './workspace.js'

interface ProcessStream {
  on(event: 'data', listener: (chunk: Buffer | string) => void): this
}

export interface TectonicProcess {
  readonly pid?: number
  readonly stdout: ProcessStream | null
  readonly stderr: ProcessStream | null
  on(event: 'error', listener: (error: Error) => void): this
  on(event: 'close', listener: (code: number | null, signal: NodeJS.Signals | null) => void): this
  kill(signal?: NodeJS.Signals): boolean
}

export type SpawnTectonic = (
  executable: string,
  args: readonly string[],
  options: SpawnOptions,
) => TectonicProcess

export interface RunTectonicRequest {
  readonly executable: string
  readonly bundlePath: string
  readonly tectonicCacheDirectory?: string
  readonly mainFile: string
  readonly workspace: CompileWorkspace
  readonly signal?: AbortSignal
  readonly totalTimeoutMs?: number
  readonly idleTimeoutMs?: number
  readonly maxOutputBytes?: number
  readonly killGraceMs?: number
  readonly killAttemptTimeoutMs?: number
  readonly maxOutputArtifactBytes?: number
  readonly maxOutputDirectoryBytes?: number
  readonly maxOutputEntries?: number
  readonly outputPollIntervalMs?: number
  readonly spawn?: SpawnTectonic
  readonly killTree?: (pid: number) => void | Promise<void>
}

export interface TectonicRunResult {
  readonly exitCode: 0
  readonly signal: null
  readonly log: string
}

export class TectonicRunError extends LatexCompilerError {
  readonly log: string
  terminationConfirmed = false
  quarantinedWorkspace: string | null = null
  readonly exitCode: number | null
  readonly signal: NodeJS.Signals | null

  constructor(
    code: LatexCompilerErrorCode,
    message: string,
    log: string,
    exitCode: number | null = null,
    signal: NodeJS.Signals | null = null,
    cause?: unknown,
  ) {
    super(code, message, cause)
    this.name = 'TectonicRunError'
    this.log = log
    this.exitCode = exitCode
    this.signal = signal
  }
}

export interface ProcessTreeKillOptions {
  readonly platform?: NodeJS.Platform
  readonly systemRoot?: string
  readonly spawn?: (
    executable: string,
    args: readonly string[],
    options: SpawnOptions,
  ) => {
    once(event: 'error', listener: (error: Error) => void): unknown
    once(event: 'close', listener: (code: number | null) => void): unknown
  }
}

export function killProcessTree(
  child: TectonicProcess,
  pid: number,
  options: ProcessTreeKillOptions = {},
): Promise<void> {
  if ((options.platform ?? process.platform) === 'win32') {
    const systemRoot = options.systemRoot ?? process.env.SystemRoot
    if (!systemRoot || !win32.isAbsolute(systemRoot) || systemRoot.includes('\0')) {
      return Promise.reject(new Error('A trusted absolute SystemRoot is required for taskkill'))
    }
    const taskkillPath = win32.join(systemRoot, 'System32', 'taskkill.exe')
    return new Promise((resolveKill, rejectKill) => {
      const spawnKiller =
        options.spawn ?? (nodeSpawn as unknown as NonNullable<ProcessTreeKillOptions['spawn']>)
      let killer: ReturnType<NonNullable<ProcessTreeKillOptions['spawn']>>
      try {
        killer = spawnKiller(taskkillPath, ['/pid', String(pid), '/t', '/f'], {
          shell: false,
          windowsHide: true,
          stdio: 'ignore',
        })
      } catch (error) {
        rejectKill(error)
        return
      }
      let settled = false
      const rejectOnce = (error: Error) => {
        if (settled) return
        settled = true
        rejectKill(error)
      }
      killer.once('error', rejectOnce)
      killer.once('close', (code) => {
        if (settled) return
        settled = true
        if (code === 0) resolveKill()
        else rejectKill(new Error(`taskkill exited with code ${code ?? 'null'}`))
      })
    })
  }
  try {
    process.kill(-pid, 'SIGKILL')
  } catch {
    child.kill('SIGKILL')
  }
  return Promise.resolve()
}

interface OutputDirectoryLimits {
  readonly maxFileBytes: number
  readonly maxTotalBytes: number
  readonly maxEntries: number
}

async function validateOutputDirectory(
  directory: string,
  limits: OutputDirectoryLimits,
): Promise<void> {
  let entries = 0
  let totalBytes = 0
  const visit = async (path: string): Promise<void> => {
    for (const entry of await readdir(path, { withFileTypes: true })) {
      entries += 1
      if (entries > limits.maxEntries) {
        throw new LatexCompilerError('TECTONIC_OUTPUT_LIMIT', 'Output entry limit exceeded')
      }
      const target = join(path, entry.name)
      const stats = await lstat(target)
      if (stats.isSymbolicLink()) {
        throw new LatexCompilerError('TECTONIC_OUTPUT_LIMIT', 'Output links are not allowed')
      }
      if (stats.isDirectory()) {
        await visit(target)
      } else if (stats.isFile()) {
        if (stats.size > limits.maxFileBytes) {
          throw new LatexCompilerError('TECTONIC_OUTPUT_LIMIT', 'Output file limit exceeded')
        }
        totalBytes += stats.size
        if (totalBytes > limits.maxTotalBytes) {
          throw new LatexCompilerError('TECTONIC_OUTPUT_LIMIT', 'Output total limit exceeded')
        }
      } else {
        throw new LatexCompilerError('TECTONIC_OUTPUT_LIMIT', 'Invalid output entry')
      }
    }
  }
  await visit(directory)
}

export function runTectonic(request: RunTectonicRequest): Promise<TectonicRunResult> {
  const remoteIndexedBundle = isRemoteIndexedBundleUrl(request.bundlePath)
  const tectonicCacheDirectory =
    request.tectonicCacheDirectory ?? join(request.workspace.root, '.tectonic-cache')
  const unsupportedLocalTar =
    isAbsolute(request.bundlePath) && request.bundlePath.toLowerCase().endsWith('.tar')
  if (
    request.mainFile !== request.workspace.mainFile ||
    !isAbsolute(request.executable) ||
    !isAbsolute(tectonicCacheDirectory) ||
    (!isAbsolute(request.bundlePath) && !remoteIndexedBundle) ||
    unsupportedLocalTar
  ) {
    throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Untrusted compiler path')
  }
  const spawn = request.spawn ?? (nodeSpawn as unknown as SpawnTectonic)
  const args = [request.mainFile, '--untrusted']
  if (!remoteIndexedBundle) args.push('--only-cached')
  args.push(
    '--synctex',
    '--bundle',
    request.bundlePath,
    '--outdir',
    request.workspace.outputDirectory,
  )
  const child = spawn(request.executable, args, {
    cwd: request.workspace.inputDirectory,
    shell: false,
    detached: true,
    windowsHide: true,
    stdio: ['ignore', 'pipe', 'pipe'],
    env: {
      LANG: 'C.UTF-8',
      TECTONIC_CACHE_DIR: tectonicCacheDirectory,
      TECTONIC_UNTRUSTED_MODE: '1',
    },
  })

  return new Promise((resolveRun, rejectRun) => {
    const totalTimeoutMs = request.totalTimeoutMs ?? 120_000
    const idleTimeoutMs = request.idleTimeoutMs ?? 30_000
    const maxOutputBytes = request.maxOutputBytes ?? 2 * 1024 * 1024
    const killGraceMs = request.killGraceMs ?? 1_000
    const killAttemptTimeoutMs = request.killAttemptTimeoutMs ?? 1_000
    const outputPollIntervalMs = request.outputPollIntervalMs ?? 250
    const outputLimits = {
      maxFileBytes: request.maxOutputArtifactBytes ?? 256 * 1024 * 1024,
      maxTotalBytes: request.maxOutputDirectoryBytes ?? 512 * 1024 * 1024,
      maxEntries: request.maxOutputEntries ?? 1_000,
    }
    const chunks: Buffer[] = []
    const closeWaiters = new Set<(closed: boolean) => void>()
    let outputBytes = 0
    let settled = false
    let childClosed = false
    let terminating: TectonicRunError | null = null
    let idleTimer: NodeJS.Timeout
    let outputTimer: NodeJS.Timeout | undefined

    const log = () => Buffer.concat(chunks).toString('utf8')
    const cleanup = () => {
      clearTimeout(totalTimer)
      clearTimeout(idleTimer)
      if (outputTimer) clearTimeout(outputTimer)
      request.signal?.removeEventListener('abort', cancel)
      for (const waiter of closeWaiters) waiter(false)
      closeWaiters.clear()
    }
    const rejectFailure = (failure: TectonicRunError, confirmed: boolean) => {
      if (settled) return
      failure.terminationConfirmed = confirmed
      settled = true
      cleanup()
      rejectRun(failure)
    }
    const waitForClose = (timeoutMs: number): Promise<boolean> => {
      if (childClosed) return Promise.resolve(true)
      return new Promise((resolveClose) => {
        const finish = (closed: boolean) => {
          clearTimeout(timer)
          closeWaiters.delete(finish)
          resolveClose(closed)
        }
        const timer = setTimeout(() => finish(false), timeoutMs)
        closeWaiters.add(finish)
      })
    }
    const terminateTree = async () => {
      if (!child.pid) throw new Error('Tectonic child PID is unavailable')
      if (request.killTree) await request.killTree(child.pid)
      else await killProcessTree(child, child.pid)
    }
    const attemptKill = async (): Promise<void> => {
      let timer: NodeJS.Timeout | undefined
      try {
        await Promise.race([
          terminateTree(),
          new Promise<never>((_resolve, reject) => {
            timer = setTimeout(
              () => reject(new Error('Process-tree kill attempt timed out')),
              killAttemptTimeoutMs,
            )
          }),
        ])
      } finally {
        if (timer) clearTimeout(timer)
      }
    }
    const finishTermination = async (failure: TectonicRunError) => {
      let treeKillConfirmed = false
      try {
        await attemptKill()
        treeKillConfirmed = true
      } catch {
        // A failed primary attempt is followed by a bounded second attempt.
      }
      if (treeKillConfirmed && (childClosed || (await waitForClose(killGraceMs)))) {
        rejectFailure(failure, true)
        return
      }
      try {
        await attemptKill()
        treeKillConfirmed = true
      } catch {
        // Failure is reported through terminationConfirmed=false below.
      }
      const rootClosed = childClosed || (await waitForClose(killGraceMs))
      const confirmed = treeKillConfirmed && rootClosed
      rejectFailure(failure, confirmed)
    }
    const beginTermination = (failure: TectonicRunError) => {
      if (settled || terminating) return
      terminating = failure
      clearTimeout(totalTimer)
      clearTimeout(idleTimer)
      if (outputTimer) clearTimeout(outputTimer)
      void finishTermination(failure)
    }
    const cancel = () =>
      beginTermination(
        new TectonicRunError('TECTONIC_CANCELLED', 'Tectonic compile cancelled', log()),
      )
    const resetIdle = () => {
      clearTimeout(idleTimer)
      idleTimer = setTimeout(
        () =>
          beginTermination(
            new TectonicRunError('TECTONIC_IDLE_TIMEOUT', 'Tectonic idle timeout', log()),
          ),
        idleTimeoutMs,
      )
    }
    const pollOutput = async () => {
      if (settled || terminating) return
      try {
        await validateOutputDirectory(request.workspace.outputDirectory, outputLimits)
      } catch (error) {
        beginTermination(
          new TectonicRunError(
            'TECTONIC_OUTPUT_LIMIT',
            'Tectonic output directory limit exceeded',
            log(),
            null,
            null,
            error,
          ),
        )
        return
      }
      outputTimer = setTimeout(() => void pollOutput(), outputPollIntervalMs)
    }
    const totalTimer = setTimeout(
      () =>
        beginTermination(
          new TectonicRunError('TECTONIC_TOTAL_TIMEOUT', 'Tectonic total timeout', log()),
        ),
      totalTimeoutMs,
    )
    const output = (chunk: Buffer | string) => {
      if (settled || terminating) return
      const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
      outputBytes += buffer.byteLength
      if (outputBytes > maxOutputBytes) {
        beginTermination(
          new TectonicRunError('TECTONIC_OUTPUT_LIMIT', 'Tectonic output limit exceeded', log()),
        )
        return
      }
      chunks.push(buffer)
      resetIdle()
    }

    child.stdout?.on('data', output)
    child.stderr?.on('data', output)
    child.on('error', (error) =>
      beginTermination(
        new TectonicRunError(
          'TECTONIC_EXIT_NONZERO',
          'Tectonic process failed',
          log(),
          null,
          null,
          error,
        ),
      ),
    )
    child.on('close', (code, signal) => {
      if (settled) return
      childClosed = true
      for (const waiter of closeWaiters) waiter(true)
      closeWaiters.clear()
      if (terminating) return
      if (code !== 0) {
        beginTermination(
          new TectonicRunError(
            'TECTONIC_EXIT_NONZERO',
            `Tectonic exited with code ${code ?? 'null'}`,
            log(),
            code,
            signal,
          ),
        )
        return
      }
      settled = true
      cleanup()
      resolveRun({ exitCode: 0, signal: null, log: log() })
    })
    request.signal?.addEventListener('abort', cancel, { once: true })
    resetIdle()
    outputTimer = setTimeout(() => void pollOutput(), outputPollIntervalMs)
    if (request.signal?.aborted) cancel()
  })
}

export interface CompileIsolatedRequest extends CompileWorkspaceLimits {
  readonly projectDirectory: string
  readonly temporaryRoot: string
  readonly cacheDirectory: string
  readonly executable: string
  readonly bundlePath: string
  readonly tectonicCacheDirectory?: string
  readonly mainFile: string
  readonly maxLogBytes?: number
  readonly maxArtifactBytes?: number
  readonly maxPublishedBytes?: number
  readonly run?: (request: RunTectonicRequest) => Promise<TectonicRunResult>
  readonly signal?: AbortSignal
  readonly totalTimeoutMs?: number
  readonly idleTimeoutMs?: number
  readonly maxOutputBytes?: number
  readonly killGraceMs?: number
  readonly killAttemptTimeoutMs?: number
  readonly maxOutputEntries?: number
  readonly outputPollIntervalMs?: number
  readonly spawn?: SpawnTectonic
  readonly killTree?: (pid: number) => void | Promise<void>
}

interface StagedFile {
  readonly name: string
  readonly bytes: number
  readonly sha256: string
}

export interface StagedCompileResult {
  readonly generationId: string
  readonly stagingDirectory: string
  readonly files: readonly StagedFile[]
  readonly log: string
  /** Trusted compile-workspace root used only to map absolute SyncTeX Input records. */
  readonly synctexInputRoot: string
  readonly workspaceCleaned: true
}

export interface CompileIsolatedResult {
  readonly generationId: string
  readonly pdfPath: string | null
  readonly synctexPath: string | null
  readonly logPath: string
  readonly log: string
  readonly synctexInputRoot: string
  readonly published: readonly string[]
  readonly workspaceCleaned: true
}

interface GenerationManifest {
  readonly schemaVersion: 1
  readonly generationId: string
  readonly files: readonly StagedFile[]
}

async function syncDirectory(path: string): Promise<void> {
  let directory
  try {
    directory = await open(path, 'r')
    await directory.sync()
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code
    if (
      process.platform !== 'win32' ||
      !['EINVAL', 'ENOTSUP', 'EPERM', 'EISDIR'].includes(code ?? '')
    ) {
      throw error
    }
  } finally {
    await directory?.close()
  }
}

async function copyBoundedFile(
  source: string,
  destination: string,
  maxFileBytes: number,
  remainingBytes: number,
): Promise<StagedFile | null> {
  let sourceStats
  try {
    sourceStats = await lstat(source)
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null
    throw error
  }
  if (!sourceStats.isFile() || sourceStats.isSymbolicLink()) {
    throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Generated artifact is not regular')
  }
  const sourceFile = await open(source, constants.O_RDONLY | constants.O_NOFOLLOW)
  const destinationFile = await open(
    destination,
    constants.O_CREAT | constants.O_EXCL | constants.O_WRONLY,
    0o600,
  )
  const hash = createHash('sha256')
  let bytes = 0
  try {
    const opened = await sourceFile.stat()
    if (opened.dev !== sourceStats.dev || opened.ino !== sourceStats.ino) {
      throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Artifact identity changed')
    }
    const buffer = Buffer.allocUnsafe(64 * 1024)
    while (true) {
      const { bytesRead } = await sourceFile.read(buffer, 0, buffer.length, null)
      if (bytesRead === 0) break
      bytes += bytesRead
      if (bytes > maxFileBytes || bytes > remainingBytes) {
        throw new LatexCompilerError(
          'TECTONIC_WORKSPACE_INVALID',
          'Generated artifact output limit exceeded',
        )
      }
      const chunk = buffer.subarray(0, bytesRead)
      hash.update(chunk)
      await destinationFile.write(chunk)
    }
    const after = await sourceFile.stat()
    if (
      after.dev !== opened.dev ||
      after.ino !== opened.ino ||
      after.size !== opened.size ||
      after.mtimeMs !== opened.mtimeMs ||
      after.ctimeMs !== opened.ctimeMs ||
      bytes !== after.size
    ) {
      throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Artifact changed while reading')
    }
    await destinationFile.sync()
    return { name: basename(destination), bytes, sha256: hash.digest('hex') }
  } finally {
    await Promise.allSettled([sourceFile.close(), destinationFile.close()])
  }
}

export async function compileIsolated(
  request: CompileIsolatedRequest,
): Promise<StagedCompileResult> {
  const workspace = await createCompileWorkspace(request.projectDirectory, request.temporaryRoot, {
    mainFile: request.mainFile,
    maxEntries: request.maxEntries,
    maxFileBytes: request.maxFileBytes,
    maxTotalBytes: request.maxTotalBytes,
    overlay: request.overlay,
    maxOverlayFiles: request.maxOverlayFiles,
    maxOverlayFileBytes: request.maxOverlayFileBytes,
    maxOverlayTotalBytes: request.maxOverlayTotalBytes,
    expectedSourceHashes: request.expectedSourceHashes,
    hooks: request.hooks,
  })
  let stagingDirectory: string | null = null
  try {
    const result = await (request.run ?? runTectonic)({
      executable: request.executable,
      bundlePath: request.bundlePath,
      tectonicCacheDirectory: request.tectonicCacheDirectory,
      mainFile: workspace.mainFile,
      workspace,
      signal: request.signal,
      totalTimeoutMs: request.totalTimeoutMs,
      idleTimeoutMs: request.idleTimeoutMs,
      maxOutputBytes: request.maxOutputBytes,
      killGraceMs: request.killGraceMs,
      killAttemptTimeoutMs: request.killAttemptTimeoutMs,
      maxOutputArtifactBytes: request.maxArtifactBytes,
      maxOutputDirectoryBytes: request.maxPublishedBytes,
      maxOutputEntries: request.maxOutputEntries,
      outputPollIntervalMs: request.outputPollIntervalMs,
      spawn: request.spawn,
      killTree: request.killTree,
    })
    const generationId = randomUUID()
    const stagingRoot = join(request.cacheDirectory, '.staging')
    stagingDirectory = join(stagingRoot, generationId)
    await mkdir(stagingRoot, { recursive: true })
    await mkdir(stagingDirectory)
    const maxArtifactBytes = request.maxArtifactBytes ?? 256 * 1024 * 1024
    const maxPublishedBytes = request.maxPublishedBytes ?? 512 * 1024 * 1024
    const stem = basename(request.mainFile, extname(request.mainFile))
    const files: StagedFile[] = []
    let publishedBytes = 0
    for (const extension of ['.pdf', '.synctex.gz']) {
      const staged = await copyBoundedFile(
        join(workspace.outputDirectory, `${stem}${extension}`),
        join(stagingDirectory, `${stem}${extension}`),
        maxArtifactBytes,
        maxPublishedBytes - publishedBytes,
      )
      if (staged) {
        publishedBytes += staged.bytes
        files.push(staged)
      }
    }
    const logBytes = Buffer.from(result.log).subarray(0, request.maxLogBytes ?? 256 * 1024)
    if (publishedBytes + logBytes.byteLength > maxPublishedBytes) {
      throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Generated output limit exceeded')
    }
    const logName = `${stem}.log`
    const logFile = await open(
      join(stagingDirectory, logName),
      constants.O_CREAT | constants.O_EXCL | constants.O_WRONLY,
      0o600,
    )
    try {
      await logFile.writeFile(logBytes)
      await logFile.sync()
    } finally {
      await logFile.close()
    }
    files.push({
      name: logName,
      bytes: logBytes.byteLength,
      sha256: createHash('sha256').update(logBytes).digest('hex'),
    })
    const manifest: GenerationManifest = {
      schemaVersion: 1,
      generationId,
      files,
    }
    const manifestFile = await open(
      join(stagingDirectory, 'manifest.json'),
      constants.O_CREAT | constants.O_EXCL | constants.O_WRONLY,
      0o600,
    )
    try {
      await manifestFile.writeFile(`${JSON.stringify(manifest)}\n`)
      await manifestFile.sync()
    } finally {
      await manifestFile.close()
    }
    await syncDirectory(stagingDirectory)
    await workspace.cleanup()
    return {
      generationId,
      stagingDirectory,
      files,
      log: logBytes.toString('utf8'),
      synctexInputRoot: workspace.inputDirectory,
      workspaceCleaned: true,
    }
  } catch (error) {
    if (stagingDirectory) await rm(stagingDirectory, { recursive: true, force: true })
    if (error instanceof TectonicRunError && !error.terminationConfirmed) {
      error.quarantinedWorkspace = workspace.root
    } else {
      await workspace.cleanup()
    }
    throw error
  }
}

export interface CommitGenerationOptions {
  readonly beforePointerCommit?: () => void | Promise<void>
  readonly maxGenerations?: number
  readonly platform?: NodeJS.Platform
  readonly syncDirectory?: (path: string) => Promise<void>
  readonly removeGeneration?: (path: string) => Promise<void>
}

const committingGenerations = new Set<string>()

async function verifyGeneration(path: string, expectedId: string): Promise<GenerationManifest> {
  const raw = JSON.parse(await readFile(join(path, 'manifest.json'), 'utf8')) as GenerationManifest
  if (
    raw.schemaVersion !== 1 ||
    raw.generationId !== expectedId ||
    !Array.isArray(raw.files) ||
    raw.files.length === 0
  ) {
    throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Invalid generation manifest')
  }
  for (const file of raw.files) {
    if (!/^[A-Za-z0-9._-]+$/.test(file.name) || !/^[a-f0-9]{64}$/.test(file.sha256)) {
      throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Invalid generation file')
    }
    const filePath = join(path, file.name)
    const fileStats = await lstat(filePath)
    if (!fileStats.isFile() || fileStats.isSymbolicLink() || fileStats.size !== file.bytes) {
      throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Generation file mismatch')
    }
    const handle = await open(filePath, constants.O_RDONLY | constants.O_NOFOLLOW)
    const hash = createHash('sha256')
    let bytes = 0
    try {
      const buffer = Buffer.allocUnsafe(64 * 1024)
      while (true) {
        const { bytesRead } = await handle.read(buffer, 0, buffer.length, null)
        if (bytesRead === 0) break
        bytes += bytesRead
        if (bytes > file.bytes) {
          throw new LatexCompilerError(
            'TECTONIC_WORKSPACE_INVALID',
            'Generation grew while reading',
          )
        }
        hash.update(buffer.subarray(0, bytesRead))
      }
    } finally {
      await handle.close()
    }
    if (bytes !== file.bytes || hash.digest('hex') !== file.sha256) {
      throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Generation digest mismatch')
    }
  }
  return raw
}

async function readCurrentGeneration(cacheDirectory: string): Promise<string | null> {
  const generationsRoot = join(cacheDirectory, 'generations')
  let pointerGeneration: string | null = null
  try {
    const value = JSON.parse(await readFile(join(cacheDirectory, 'current.json'), 'utf8')) as {
      schemaVersion?: unknown
      generationId?: unknown
    }
    pointerGeneration =
      value.schemaVersion === 1 &&
      typeof value.generationId === 'string' &&
      /^[0-9a-f-]{36}$/.test(value.generationId)
        ? value.generationId
        : null
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT' && !(error instanceof SyntaxError)) {
      throw error
    }
  }
  if (pointerGeneration) {
    const pointed = await lstat(join(generationsRoot, pointerGeneration)).catch(() => null)
    if (pointed?.isDirectory() && !pointed.isSymbolicLink()) return pointerGeneration
  }
  const entries = await readdir(generationsRoot, { withFileTypes: true }).catch((error) => {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return []
    throw error
  })
  const recoverable: Array<{ id: string; mtimeMs: number }> = []
  for (const entry of entries) {
    if (!entry.isDirectory() || !/^[0-9a-f-]{36}$/.test(entry.name)) continue
    const path = join(generationsRoot, entry.name)
    if (committingGenerations.has(path)) continue
    const stats = await lstat(path)
    if (stats.isDirectory() && !stats.isSymbolicLink()) {
      recoverable.push({ id: entry.name, mtimeMs: stats.mtimeMs })
    }
  }
  recoverable.sort((left, right) => right.mtimeMs - left.mtimeMs || right.id.localeCompare(left.id))
  return recoverable[0]?.id ?? null
}

async function pruneGenerations(
  cacheDirectory: string,
  generationsRoot: string,
  maxGenerations: number,
  removeGeneration: (path: string) => Promise<void>,
): Promise<void> {
  const current = await readCurrentGeneration(cacheDirectory)
  const entries = await readdir(generationsRoot, { withFileTypes: true })
  const generations: Array<{ id: string; path: string; mtimeMs: number }> = []
  for (const entry of entries) {
    if (!entry.isDirectory() || !/^[0-9a-f-]{36}$/.test(entry.name)) continue
    const path = join(generationsRoot, entry.name)
    const stats = await lstat(path)
    if (!stats.isDirectory() || stats.isSymbolicLink()) continue
    generations.push({ id: entry.name, path, mtimeMs: stats.mtimeMs })
  }
  generations.sort((left, right) => right.mtimeMs - left.mtimeMs || right.id.localeCompare(left.id))
  const protectedIds = new Set<string>()
  if (current) protectedIds.add(current)
  for (const path of committingGenerations) {
    if (path.startsWith(`${generationsRoot}${sep}`)) protectedIds.add(basename(path))
  }
  const keep = new Set(protectedIds)
  for (const generation of generations) {
    if (keep.size >= Math.max(1, maxGenerations)) break
    keep.add(generation.id)
  }
  await Promise.allSettled(
    generations
      .filter((generation) => !keep.has(generation.id))
      .map((generation) => removeGeneration(generation.path)),
  )
}

export async function commitCompileGeneration(
  staged: StagedCompileResult,
  cacheDirectory: string,
  options: CommitGenerationOptions = {},
): Promise<CompileIsolatedResult> {
  const stagingRoot = resolve(cacheDirectory, '.staging')
  const generationsRoot = resolve(cacheDirectory, 'generations')
  await mkdir(stagingRoot, { recursive: true })
  await mkdir(generationsRoot, { recursive: true })
  const generationDirectory = join(generationsRoot, staged.generationId)
  const sync = options.syncDirectory ?? syncDirectory
  const platform = options.platform ?? process.platform
  const prune = () =>
    pruneGenerations(
      cacheDirectory,
      generationsRoot,
      options.maxGenerations ?? 10,
      options.removeGeneration ?? ((path) => rm(path, { recursive: true, force: true })),
    )
  committingGenerations.add(generationDirectory)
  try {
    let sourceDirectory = staged.stagingDirectory
    try {
      const resolvedStage = await realpath(sourceDirectory)
      if (!resolvedStage.startsWith(`${stagingRoot}${sep}`)) {
        throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Staging path escaped cache')
      }
      await verifyGeneration(resolvedStage, staged.generationId)
      await rename(resolvedStage, generationDirectory)
      await sync(generationsRoot)
      await sync(stagingRoot)
      sourceDirectory = generationDirectory
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
      sourceDirectory = generationDirectory
    }
    const manifest = await verifyGeneration(sourceDirectory, staged.generationId)
    // On macOS, prune before switching the pointer so both the previous
    // generation and the in-flight generation remain crash-safe rollback targets.
    if (platform === 'darwin') await prune()
    await options.beforePointerCommit?.()
    const pointerPath = join(cacheDirectory, 'current.json')
    const temporaryPointer = `${pointerPath}.part-${randomUUID()}`
    try {
      const pointer = await open(
        temporaryPointer,
        constants.O_CREAT | constants.O_EXCL | constants.O_WRONLY,
        0o600,
      )
      try {
        await pointer.writeFile(
          `${JSON.stringify({ schemaVersion: 1, generationId: staged.generationId })}\n`,
        )
        await pointer.sync()
      } finally {
        await pointer.close()
      }
      await rename(temporaryPointer, pointerPath)
      // The pointer file is already fsynced and atomically renamed. A directory
      // fsync at this final boundary was observed to remain blocked on macOS,
      // preventing a successfully published compile from returning to the UI.
      if (platform !== 'darwin') await sync(cacheDirectory)
    } catch (error) {
      await rm(temporaryPointer, { force: true })
      if ((await readCurrentGeneration(cacheDirectory)) !== staged.generationId) throw error
    }

    const pathFor = (suffix: string) =>
      manifest.files.find((file) => file.name.endsWith(suffix))?.name ?? null
    const pdfName = pathFor('.pdf')
    const synctexName = pathFor('.synctex.gz')
    const logName = pathFor('.log')
    if (!logName) {
      throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Generation log missing')
    }
    const published = manifest.files.map((file) => join(generationDirectory, file.name))
    const result: CompileIsolatedResult = {
      generationId: staged.generationId,
      pdfPath: pdfName ? join(generationDirectory, pdfName) : null,
      synctexPath: synctexName ? join(generationDirectory, synctexName) : null,
      logPath: join(generationDirectory, logName),
      log: staged.log,
      synctexInputRoot: staged.synctexInputRoot,
      published,
      workspaceCleaned: true,
    }
    if (platform !== 'darwin') await prune()
    return result
  } finally {
    committingGenerations.delete(generationDirectory)
  }
}
