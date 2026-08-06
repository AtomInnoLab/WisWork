import { constants } from 'node:fs'
import { lstat, open, realpath, type FileHandle } from 'node:fs/promises'
import { dirname, isAbsolute, relative, resolve, sep } from 'node:path'

const WINDOWS_ABSOLUTE = /^[a-zA-Z]:[\\/]/

/**
 * Security boundary: this policy blocks static links and detects normal external conflicts.
 * A malicious process running as the same OS user can still mutate project files directly. Node's
 * cross-platform path APIs cannot fully eliminate parent swap-away/swap-back races; a future
 * native implementation should use directory handles with openat/renameat semantics.
 */
export type FileIdentity = Pick<Awaited<ReturnType<typeof lstat>>, 'dev' | 'ino'>

/** @internal Deterministic race hooks for filesystem security tests only. */
export interface PathPolicyTestHooks {
  afterReadOpen?: () => Promise<void>
  afterReadStat?: () => Promise<void>
}

export interface PreparedWriteTarget {
  normalizedPath: string
  path: string
  parentPath: string
  parentRealPath: string
  parentIdentity: FileIdentity
}

export interface OpenedProjectFile {
  handle: FileHandle
  normalizedPath: string
  stat: Awaited<ReturnType<FileHandle['stat']>>
}

export class ProjectPathPolicy {
  readonly realRootPath: string
  private readonly hooks: PathPolicyTestHooks
  private readonly rootIdentity: FileIdentity

  private constructor(
    realRootPath: string,
    hooks: PathPolicyTestHooks,
    rootIdentity: FileIdentity,
  ) {
    this.realRootPath = realRootPath
    this.hooks = hooks
    this.rootIdentity = rootIdentity
  }

  static async open(rootPath: string, hooks: PathPolicyTestHooks = {}): Promise<ProjectPathPolicy> {
    const root = await realpath(rootPath)
    const info = await lstat(root)
    if (!info.isDirectory()) throw new Error('LaTeX project root is not a directory')
    return new ProjectPathPolicy(root, hooks, identity(info))
  }

  normalize(relativePath: string): string {
    if (relativePath.includes('\0')) throw new Error('Project paths must not contain NUL bytes')
    if (
      relativePath.length === 0 ||
      relativePath.startsWith('\\') ||
      isAbsolute(relativePath) ||
      WINDOWS_ABSOLUTE.test(relativePath)
    ) {
      throw new Error('Project paths must be non-empty relative paths')
    }

    const slashPath = relativePath.replaceAll('\\', '/')
    const segments = slashPath.split('/')
    if (segments.includes('..')) throw new Error('Project path traversal is not allowed')

    const normalized = segments.filter((segment) => segment !== '' && segment !== '.').join('/')
    if (!normalized) throw new Error('Project paths must be non-empty relative paths')
    return normalized
  }

  async resolveExisting(relativePath: string, expected?: 'file' | 'directory'): Promise<string> {
    await this.assertRootUnchanged()
    const normalized = this.normalize(relativePath)
    const target = this.absolute(normalized)
    const info = await this.checkSegments(normalized, false)
    if (!info) throw new Error(`Project path does not exist: ${normalized}`)
    if (expected === 'file' && !info.isFile()) {
      throw new Error(`Project path is not a file: ${normalized}`)
    }
    if (expected === 'directory' && !info.isDirectory()) {
      throw new Error(`Project path is not a directory: ${normalized}`)
    }
    return target
  }

  async resolveForWrite(relativePath: string): Promise<string> {
    return (await this.prepareWrite(relativePath)).path
  }

  async openTextFile(relativePath: string): Promise<OpenedProjectFile> {
    await this.assertRootUnchanged()
    const normalizedPath = this.normalize(relativePath)
    const path = await this.resolveExisting(normalizedPath, 'file')
    const noFollow = 'O_NOFOLLOW' in constants ? constants.O_NOFOLLOW : 0
    const handle = await open(path, constants.O_RDONLY | noFollow)
    try {
      await this.hooks.afterReadOpen?.()
      await this.assertRootUnchanged()
      const [handleStat, pathStat, currentRealPath] = await Promise.all([
        handle.stat(),
        lstat(path),
        realpath(path),
      ])
      this.assertInsideRoot(currentRealPath)
      if (pathStat.isSymbolicLink() || !sameIdentity(handleStat, pathStat)) {
        throw new Error(`Project file changed during validation: ${normalizedPath}`)
      }
      await this.hooks.afterReadStat?.()
      return { handle, normalizedPath, stat: handleStat }
    } catch (error) {
      await handle.close()
      throw error
    }
  }

  async prepareWrite(relativePath: string): Promise<PreparedWriteTarget> {
    await this.assertRootUnchanged()
    const normalizedPath = this.normalize(relativePath)
    const path = this.absolute(normalizedPath)
    const info = await this.checkSegments(normalizedPath, true)
    if (info && !info.isFile()) throw new Error(`Write target is not a file: ${normalizedPath}`)

    const parentPath = dirname(path)
    const parentRealPath = await realpath(parentPath)
    this.assertInsideRoot(parentRealPath, true)
    const parentStat = await lstat(parentPath)
    if (!parentStat.isDirectory() || parentStat.isSymbolicLink()) {
      throw new Error(`Write parent is not a safe directory: ${normalizedPath}`)
    }
    return {
      normalizedPath,
      path,
      parentPath,
      parentRealPath,
      parentIdentity: identity(parentStat),
    }
  }

  async validateWriteTarget(
    target: PreparedWriteTarget,
    stage: 'before-rename' | 'after-rename',
    expectedIdentity?: FileIdentity,
  ): Promise<void> {
    await this.assertRootUnchanged()
    const [parentRealPath, parentStat] = await Promise.all([
      realpath(target.parentPath),
      lstat(target.parentPath),
    ])
    this.assertInsideRoot(parentRealPath, true)
    if (
      parentRealPath !== target.parentRealPath ||
      parentStat.isSymbolicLink() ||
      !parentStat.isDirectory() ||
      !sameIdentity(parentStat, target.parentIdentity)
    ) {
      throw new Error(`Write parent directory changed during validation: ${target.normalizedPath}`)
    }

    try {
      const targetStat = await lstat(target.path)
      if (
        targetStat.isSymbolicLink() ||
        !targetStat.isFile() ||
        (expectedIdentity !== undefined && !sameIdentity(targetStat, expectedIdentity))
      ) {
        throw new Error(`Write target is not a safe regular file: ${target.normalizedPath}`)
      }
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT' && stage === 'before-rename') return
      throw error
    }
  }

  async assertRootUnchanged(): Promise<void> {
    try {
      const [currentRealPath, currentStat] = await Promise.all([
        realpath(this.realRootPath),
        lstat(this.realRootPath),
      ])
      if (
        currentRealPath !== this.realRootPath ||
        currentStat.isSymbolicLink() ||
        !currentStat.isDirectory() ||
        !sameIdentity(currentStat, this.rootIdentity)
      ) {
        throw new Error('Project root changed after it was opened')
      }
    } catch (error) {
      if (error instanceof Error && /Project root changed/.test(error.message)) throw error
      throw new Error('Project root changed after it was opened', { cause: error })
    }
  }

  private absolute(normalized: string): string {
    const target = resolve(this.realRootPath, ...normalized.split('/'))
    const fromRoot = relative(this.realRootPath, target)
    if (fromRoot === '..' || fromRoot.startsWith(`..${sep}`) || isAbsolute(fromRoot)) {
      throw new Error('Project path traversal is not allowed')
    }
    return target
  }

  private assertInsideRoot(path: string, allowRoot = false): void {
    const fromRoot = relative(this.realRootPath, path)
    if (
      (!allowRoot && fromRoot === '') ||
      fromRoot === '..' ||
      fromRoot.startsWith(`..${sep}`) ||
      isAbsolute(fromRoot)
    ) {
      throw new Error('Resolved project path escaped the real project root')
    }
  }

  private async checkSegments(normalized: string, allowMissingLeaf: boolean) {
    const segments = normalized.split('/')
    let current = this.realRootPath

    for (let index = 0; index < segments.length; index += 1) {
      current = resolve(current, segments[index]!)
      let info
      try {
        info = await lstat(current)
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
          if (allowMissingLeaf && index === segments.length - 1) return undefined
          throw new Error(`Project path does not exist: ${normalized}`, { cause: error })
        }
        throw error
      }

      if (info.isSymbolicLink()) {
        throw new Error(`Project path contains a symbolic link: ${normalized}`)
      }
      if (index < segments.length - 1 && !info.isDirectory()) {
        throw new Error(`Project path segment is not a directory: ${normalized}`)
      }
      if (index === segments.length - 1) return info
    }

    throw new Error(`Invalid project path: ${normalized}`)
  }
}

function identity(value: FileIdentity): FileIdentity {
  return { dev: value.dev, ino: value.ino }
}

function sameIdentity(left: FileIdentity, right: FileIdentity): boolean {
  return left.dev === right.dev && left.ino === right.ino
}
