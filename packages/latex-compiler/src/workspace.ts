import { constants, type Stats } from 'node:fs'
import { lstat, mkdir, mkdtemp, open, readdir, realpath, rm } from 'node:fs/promises'
import { dirname, isAbsolute, join, normalize, relative, resolve, sep } from 'node:path'
import { LatexCompilerError } from './errors.js'

export interface CompileWorkspaceHooks {
  readonly afterDirectoryRead?: (path: string) => void | Promise<void>
  readonly afterFileOpen?: (path: string) => void | Promise<void>
  readonly beforeFinalValidation?: () => void | Promise<void>
}

export interface CompileWorkspaceLimits {
  readonly maxEntries?: number
  readonly maxFileBytes?: number
  readonly maxTotalBytes?: number
  readonly mainFile?: string
  readonly hooks?: CompileWorkspaceHooks
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
})

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
  const limits = { ...DEFAULT_LIMITS, ...options }
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
