import { constants, type Stats } from 'node:fs'
import { lstat, mkdir, mkdtemp, open, readdir, realpath, rm } from 'node:fs/promises'
import { dirname, extname, isAbsolute, join, normalize, relative, resolve, sep } from 'node:path'
import { LatexCompilerError } from './errors.js'

export interface CompileWorkspaceHooks {
  readonly afterDirectoryRead?: (path: string) => void | Promise<void>
  readonly afterFileOpen?: (path: string) => void | Promise<void>
  readonly beforeFinalValidation?: () => void | Promise<void>
  /** @internal Deterministic overlay race hooks for security tests only. */
  readonly beforeOverlayTargetOpen?: (path: string) => void | Promise<void>
  readonly afterOverlayTargetOpen?: (path: string) => void | Promise<void>
}

export interface CompileWorkspaceLimits {
  readonly maxEntries?: number
  readonly maxFileBytes?: number
  readonly maxTotalBytes?: number
  readonly overlay?: readonly CompileTextOverlay[]
  readonly maxOverlayFiles?: number
  readonly maxOverlayFileBytes?: number
  readonly maxOverlayTotalBytes?: number
  readonly mainFile?: string
  readonly hooks?: CompileWorkspaceHooks
}

export interface CompileTextOverlay {
  readonly path: string
  readonly text: string
}

export interface CompileWorkspace {
  readonly root: string
  readonly inputDirectory: string
  readonly outputDirectory: string
  readonly mainFile: string
  cleanup(): Promise<void>
}

const DEFAULT_LIMITS = Object.freeze({
  maxEntries: 10_000,
  maxFileBytes: 100 * 1024 * 1024,
  maxTotalBytes: 500 * 1024 * 1024,
  maxOverlayFiles: 20,
  maxOverlayFileBytes: 2 * 1024 * 1024,
  maxOverlayTotalBytes: 4 * 1024 * 1024,
})

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

interface SourceIdentity {
  readonly path: string
  readonly realPath: string
  readonly dev: number
  readonly ino: number
  readonly size: number
  readonly mtimeMs: number
  readonly ctimeMs: number
  readonly directory: boolean
}

function safeRelativePath(path: string): string {
  if (!path || path.includes('\0') || isAbsolute(path) || /^[a-zA-Z]:[\\/]/.test(path)) {
    throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Invalid compile path')
  }
  const normalized = normalize(path).split(sep).join('/')
  if (normalized === '..' || normalized.startsWith('../')) {
    throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Compile path traversal')
  }
  return normalized
}

function overlayError(message: string, cause?: unknown): LatexCompilerError {
  return new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', message, cause)
}

function assertUnicodeScalarString(value: string, label: string): void {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index)
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1)
      if (index + 1 >= value.length || next < 0xdc00 || next > 0xdfff) {
        throw overlayError(`${label} is not valid Unicode`)
      }
      index += 1
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      throw overlayError(`${label} is not valid Unicode`)
    }
  }
}

function isWindowsReservedComponent(component: string): boolean {
  return /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(component)
}

function normalizeOverlayPath(path: unknown): string {
  if (typeof path !== 'string' || !path || path.includes('\0')) {
    throw overlayError('Invalid overlay path')
  }
  assertUnicodeScalarString(path, 'Overlay path')
  if (path.startsWith('/') || path.startsWith('\\') || /^[a-zA-Z]:[\\/]/.test(path)) {
    throw overlayError('Overlay path must be relative')
  }
  const segments = path.normalize('NFC').replaceAll('\\', '/').split('/')
  if (segments.includes('..')) throw overlayError('Overlay path traversal is not allowed')
  const meaningfulSegments = segments.filter((segment) => segment && segment !== '.')
  for (const segment of meaningfulSegments) {
    if (
      segment.includes(':') ||
      /[. ]$/.test(segment) ||
      isWindowsReservedComponent(segment)
    ) {
      throw overlayError(`Overlay path is not portable: ${path}`)
    }
  }
  const normalized = meaningfulSegments.join('/')
  if (!normalized) throw overlayError('Invalid overlay path')
  const lower = normalized.toLowerCase()
  if (lower !== 'tectonic.toml' && !TEXT_EXTENSIONS.has(extname(lower))) {
    throw overlayError(`Unsupported overlay text file type: ${normalized}`)
  }
  return normalized
}

function assertValidOverlayText(text: unknown, path: string): asserts text is string {
  if (typeof text !== 'string' || text.includes('\0')) {
    throw overlayError(`Overlay contains binary text: ${path}`)
  }
  assertUnicodeScalarString(text, `Overlay text for ${path}`)
}

function validatePositiveLimit(value: number, name: string): void {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw overlayError(`${name} must be a positive safe integer`)
  }
}

function validateOverlay(
  overlay: readonly CompileTextOverlay[] | undefined,
  limits: Pick<
    Required<CompileWorkspaceLimits>,
    'maxOverlayFiles' | 'maxOverlayFileBytes' | 'maxOverlayTotalBytes'
  >,
): readonly CompileTextOverlay[] {
  validatePositiveLimit(limits.maxOverlayFiles, 'Overlay file count limit')
  validatePositiveLimit(limits.maxOverlayFileBytes, 'Overlay file size limit')
  validatePositiveLimit(limits.maxOverlayTotalBytes, 'Overlay total size limit')
  if (overlay === undefined) return []
  if (!Array.isArray(overlay) || overlay.length > limits.maxOverlayFiles) {
    throw overlayError('Overlay file count limit exceeded')
  }
  const seen = new Set<string>()
  let totalBytes = 0
  return overlay.map((file) => {
    if (!file || typeof file !== 'object') throw overlayError('Invalid overlay file')
    const path = normalizeOverlayPath(file.path)
    const duplicateKey = path.toLowerCase()
    if (seen.has(duplicateKey)) throw overlayError(`Duplicate overlay path: ${path}`)
    seen.add(duplicateKey)
    assertValidOverlayText(file.text, path)
    const bytes = Buffer.byteLength(file.text, 'utf8')
    if (bytes > limits.maxOverlayFileBytes) {
      throw overlayError(`Overlay file size limit exceeded: ${path}`)
    }
    totalBytes += bytes
    if (totalBytes > limits.maxOverlayTotalBytes) {
      throw overlayError('Overlay total size limit exceeded')
    }
    return { path, text: file.text }
  })
}

interface OverlayDirectoryIdentity {
  readonly path: string
  readonly realPath: string
  readonly dev: number
  readonly ino: number
}

async function captureOverlayDirectoryChain(
  inputDirectory: string,
  parent: string,
): Promise<readonly OverlayDirectoryIdentity[]> {
  const parentRelative = relative(inputDirectory, parent)
  if (parentRelative.startsWith('..') || isAbsolute(parentRelative)) {
    throw overlayError('Overlay parent escaped workspace')
  }
  const paths = [inputDirectory]
  let current = inputDirectory
  for (const segment of parentRelative.split(sep).filter(Boolean)) {
    current = join(current, segment)
    paths.push(current)
  }
  return Promise.all(
    paths.map(async (path) => {
      const [realPath, stats] = await Promise.all([realpath(path), lstat(path)])
      if (
        realPath !== path ||
        stats.isSymbolicLink() ||
        !stats.isDirectory() ||
        (realPath !== inputDirectory && !realPath.startsWith(`${inputDirectory}${sep}`))
      ) {
        throw overlayError(`Overlay parent is not a safe directory: ${parentRelative}`)
      }
      return { path, realPath, dev: stats.dev, ino: stats.ino }
    }),
  )
}

async function validateOverlayDirectoryChain(
  expected: readonly OverlayDirectoryIdentity[],
): Promise<void> {
  for (const directory of expected) {
    const [realPath, stats] = await Promise.all([realpath(directory.path), lstat(directory.path)])
    if (
      realPath !== directory.realPath ||
      stats.isSymbolicLink() ||
      !stats.isDirectory() ||
      stats.dev !== directory.dev ||
      stats.ino !== directory.ino
    ) {
      throw overlayError(`Overlay parent changed during validation: ${directory.path}`)
    }
  }
}

async function applyOverlayFile(
  inputDirectory: string,
  file: CompileTextOverlay,
  hooks: CompileWorkspaceHooks | undefined,
): Promise<void> {
  const target = resolve(inputDirectory, file.path)
  if (!target.startsWith(`${inputDirectory}${sep}`)) throw overlayError('Overlay escaped workspace')
  const parent = dirname(target)
  try {
    const directoryChain = await captureOverlayDirectoryChain(inputDirectory, parent)

    let targetStats: Stats | undefined
    try {
      targetStats = await lstat(target)
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
    }
    if (targetStats?.isSymbolicLink() || (targetStats && !targetStats.isFile())) {
      throw overlayError(`Overlay target is not a safe regular file: ${file.path}`)
    }
    if (targetStats && (await realpath(target)) !== target) {
      throw overlayError(`Overlay target escaped workspace: ${file.path}`)
    }

    await hooks?.beforeOverlayTargetOpen?.(target)
    // Node does not expose portable openat(). Revalidating every ancestor immediately before open
    // closes deterministic swaps; once the handle is open, a second validation protects writes.
    await validateOverlayDirectoryChain(directoryChain)
    const flags = targetStats
      ? constants.O_WRONLY | constants.O_NOFOLLOW
      : constants.O_CREAT | constants.O_EXCL | constants.O_WRONLY | constants.O_NOFOLLOW
    const handle = await open(target, flags, 0o600)
    try {
      const opened = await handle.stat()
      if (
        !opened.isFile() ||
        (targetStats && (opened.dev !== targetStats.dev || opened.ino !== targetStats.ino))
      ) {
        throw overlayError(`Overlay target changed during validation: ${file.path}`)
      }
      await hooks?.afterOverlayTargetOpen?.(target)
      await validateOverlayDirectoryChain(directoryChain)
      if (targetStats) await handle.truncate(0)
      await handle.writeFile(file.text, 'utf8')
    } finally {
      await handle.close()
    }
  } catch (error) {
    if (error instanceof LatexCompilerError) throw error
    throw overlayError(`Invalid overlay target: ${file.path}`, error)
  }
}

function identity(path: string, realPath: string, stats: Stats): SourceIdentity {
  return {
    path,
    realPath,
    dev: stats.dev,
    ino: stats.ino,
    size: stats.size,
    mtimeMs: stats.mtimeMs,
    ctimeMs: stats.ctimeMs,
    directory: stats.isDirectory(),
  }
}

function sameIdentity(expected: SourceIdentity, actual: Stats): boolean {
  return (
    expected.dev === actual.dev &&
    expected.ino === actual.ino &&
    expected.size === actual.size &&
    expected.mtimeMs === actual.mtimeMs &&
    expected.ctimeMs === actual.ctimeMs &&
    expected.directory === actual.isDirectory()
  )
}

export async function createCompileWorkspace(
  projectDirectory: string,
  temporaryRoot: string,
  options: CompileWorkspaceLimits = {},
): Promise<CompileWorkspace> {
  const limits = {
    maxEntries: options.maxEntries ?? DEFAULT_LIMITS.maxEntries,
    maxFileBytes: options.maxFileBytes ?? DEFAULT_LIMITS.maxFileBytes,
    maxTotalBytes: options.maxTotalBytes ?? DEFAULT_LIMITS.maxTotalBytes,
    maxOverlayFiles: options.maxOverlayFiles ?? DEFAULT_LIMITS.maxOverlayFiles,
    maxOverlayFileBytes: options.maxOverlayFileBytes ?? DEFAULT_LIMITS.maxOverlayFileBytes,
    maxOverlayTotalBytes: options.maxOverlayTotalBytes ?? DEFAULT_LIMITS.maxOverlayTotalBytes,
  }
  const overlay = validateOverlay(options.overlay, limits)
  const mainFile = safeRelativePath(options.mainFile ?? 'main.tex')
  const projectRoot = await realpath(projectDirectory)
  await mkdir(dirname(temporaryRoot), { recursive: true })
  const root = await mkdtemp(`${temporaryRoot}-`)
  const inputDirectory = join(root, 'input')
  const outputDirectory = join(root, 'output')
  const identities: SourceIdentity[] = []
  let entries = 0
  let totalBytes = 0
  try {
    await mkdir(inputDirectory)
    await mkdir(outputDirectory)
    const copyDirectory = async (
      sourceDirectory: string,
      targetDirectory: string,
    ): Promise<void> => {
      const directoryStats = await lstat(sourceDirectory)
      const directoryRealPath = await realpath(sourceDirectory)
      if (
        !directoryStats.isDirectory() ||
        directoryStats.isSymbolicLink() ||
        (directoryRealPath !== projectRoot && !directoryRealPath.startsWith(`${projectRoot}${sep}`))
      ) {
        throw new LatexCompilerError(
          'TECTONIC_WORKSPACE_INVALID',
          'Directory identity escaped project',
        )
      }
      const directoryIdentity = identity(sourceDirectory, directoryRealPath, directoryStats)
      const directoryEntries = await readdir(sourceDirectory, { withFileTypes: true })
      await options.hooks?.afterDirectoryRead?.(sourceDirectory)
      const afterRead = await lstat(sourceDirectory)
      if (
        !sameIdentity(directoryIdentity, afterRead) ||
        (await realpath(sourceDirectory)) !== directoryRealPath
      ) {
        throw new LatexCompilerError(
          'TECTONIC_WORKSPACE_INVALID',
          'Directory changed during isolation',
        )
      }
      identities.push(directoryIdentity)

      for (const entry of directoryEntries) {
        entries += 1
        if (entries > limits.maxEntries) {
          throw new LatexCompilerError(
            'TECTONIC_WORKSPACE_INVALID',
            'Workspace entry limit exceeded',
          )
        }
        const source = join(sourceDirectory, entry.name)
        const target = join(targetDirectory, entry.name)
        const sourceRelative = relative(projectRoot, source)
        if (sourceRelative.startsWith('..') || isAbsolute(sourceRelative)) {
          throw new LatexCompilerError(
            'TECTONIC_WORKSPACE_INVALID',
            'Workspace path escaped project',
          )
        }
        const stats = await lstat(source)
        if (stats.isSymbolicLink()) {
          throw new LatexCompilerError(
            'TECTONIC_WORKSPACE_INVALID',
            'Symbolic links are not allowed',
          )
        }
        if (stats.isDirectory()) {
          await mkdir(target)
          await copyDirectory(source, target)
          continue
        }
        if (!stats.isFile()) {
          throw new LatexCompilerError(
            'TECTONIC_WORKSPACE_INVALID',
            'Only regular files and directories are allowed',
          )
        }

        const sourceFile = await open(source, constants.O_RDONLY | constants.O_NOFOLLOW)
        const targetFile = await open(
          target,
          constants.O_CREAT | constants.O_EXCL | constants.O_WRONLY,
          0o600,
        )
        try {
          const opened = await sourceFile.stat()
          if (
            !opened.isFile() ||
            opened.dev !== stats.dev ||
            opened.ino !== stats.ino ||
            opened.size !== stats.size ||
            opened.mtimeMs !== stats.mtimeMs ||
            opened.ctimeMs !== stats.ctimeMs
          ) {
            throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Source changed before read')
          }
          const sourceIdentity = identity(source, await realpath(source), opened)
          await options.hooks?.afterFileOpen?.(source)
          const buffer = Buffer.allocUnsafe(64 * 1024)
          let fileBytes = 0
          while (true) {
            const { bytesRead } = await sourceFile.read(buffer, 0, buffer.length, null)
            if (bytesRead === 0) break
            fileBytes += bytesRead
            totalBytes += bytesRead
            if (fileBytes > limits.maxFileBytes || totalBytes > limits.maxTotalBytes) {
              throw new LatexCompilerError(
                'TECTONIC_WORKSPACE_INVALID',
                'Workspace byte limit exceeded while reading',
              )
            }
            await targetFile.write(buffer.subarray(0, bytesRead))
          }
          const afterFileRead = await sourceFile.stat()
          if (!sameIdentity(sourceIdentity, afterFileRead) || fileBytes !== afterFileRead.size) {
            throw new LatexCompilerError(
              'TECTONIC_WORKSPACE_INVALID',
              'Source changed while reading',
            )
          }
          identities.push(sourceIdentity)
        } finally {
          await Promise.allSettled([sourceFile.close(), targetFile.close()])
        }
      }
    }

    await copyDirectory(projectRoot, inputDirectory)
    await options.hooks?.beforeFinalValidation?.()
    for (const expected of identities) {
      const actual = await lstat(expected.path)
      if (
        !sameIdentity(expected, actual) ||
        (expected.directory && (await realpath(expected.path)) !== expected.realPath)
      ) {
        throw new LatexCompilerError(
          'TECTONIC_WORKSPACE_INVALID',
          'Project revision changed during isolation',
        )
      }
    }

    for (const file of overlay) await applyOverlayFile(inputDirectory, file, options.hooks)

    const mainPath = resolve(inputDirectory, mainFile)
    if (!mainPath.startsWith(`${inputDirectory}${sep}`)) {
      throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Main file escaped workspace')
    }
    const mainStats = await lstat(mainPath)
    if (!mainStats.isFile() || mainStats.isSymbolicLink()) {
      throw new LatexCompilerError('TECTONIC_WORKSPACE_INVALID', 'Main file is not regular')
    }
    return {
      root,
      inputDirectory,
      outputDirectory,
      mainFile,
      cleanup: () => rm(root, { recursive: true, force: true }),
    }
  } catch (error) {
    await rm(root, { recursive: true, force: true })
    throw error
  }
}
