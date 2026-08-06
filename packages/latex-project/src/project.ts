import { createHash } from 'node:crypto'
import { lstat, mkdir, readdir, realpath, rename, unlink } from 'node:fs/promises'
import { dirname, extname, join } from 'node:path'
import { atomicWriteFile } from './atomic-write.js'
import { discoverMainFile } from './main-file.js'
import { ProjectPathPolicy, ProjectWriteConflictError } from './path-policy.js'
import {
  DEFAULT_MAX_TEXT_BYTES,
  type DeleteTextOptions,
  type LatexProjectOptions,
  type MainFileDiscovery,
  type SavedText,
  type SaveTextOptions,
} from './types.js'

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
const TEXT_FILENAMES = new Set(['tectonic.toml'])
const ARTICLE_TEMPLATE = `\\documentclass{article}
\\usepackage[utf8]{inputenc}

\\title{Untitled}
\\author{}
\\date{\\today}

\\begin{document}
\\maketitle

\\end{document}
`

/** Main-process domain object. Never serialize it directly across preload; define an explicit DTO. */
export class LatexProject {
  readonly rootPath: string
  readonly mainFile: string | undefined
  readonly mainFileDiscovery: MainFileDiscovery
  private readonly maxTextBytes: number
  private readonly policy: ProjectPathPolicy

  private constructor(
    rootPath: string,
    policy: ProjectPathPolicy,
    maxTextBytes: number,
    mainFileDiscovery: MainFileDiscovery,
  ) {
    this.rootPath = rootPath
    this.policy = policy
    this.maxTextBytes = maxTextBytes
    this.mainFileDiscovery = mainFileDiscovery
    this.mainFile = mainFileDiscovery.kind === 'found' ? mainFileDiscovery.path : undefined
  }

  static async open(rootPath: string, options: LatexProjectOptions = {}): Promise<LatexProject> {
    const maxTextBytes = options.maxTextBytes ?? DEFAULT_MAX_TEXT_BYTES
    if (!Number.isSafeInteger(maxTextBytes) || maxTextBytes <= 0) {
      throw new Error('maxTextBytes must be a positive safe integer')
    }
    const policy = await ProjectPathPolicy.open(rootPath, options.pathHooks)
    const canonicalRoot = await realpath(rootPath)
    const provisional = new LatexProject(canonicalRoot, policy, maxTextBytes, {
      kind: 'not-found',
      candidates: [],
    })
    const discovery = await discoverMainFile(
      policy,
      { savedMainFile: options.savedMainFile },
      provisional,
    )
    return new LatexProject(canonicalRoot, policy, maxTextBytes, discovery)
  }

  async listTextFiles(): Promise<string[]> {
    const results: string[] = []
    const walk = async (relativeDir: string): Promise<void> => {
      await this.policy.assertRootUnchanged()
      const absoluteDir = relativeDir
        ? await this.policy.resolveExisting(relativeDir, 'directory')
        : this.policy.realRootPath
      const entries = await readdir(absoluteDir, { withFileTypes: true })
      for (const entry of entries) {
        const path = relativeDir ? `${relativeDir}/${entry.name}` : entry.name
        if (entry.isSymbolicLink()) continue
        if (entry.isDirectory()) await walk(path)
        else if (entry.isFile() && isTextPath(path)) results.push(path)
      }
    }
    await walk('')
    return results.sort()
  }

  async readText(relativePath: string): Promise<string> {
    const normalized = this.policy.normalize(relativePath)
    if (!isTextPath(normalized)) throw new Error(`Unsupported text file type: ${normalized}`)
    const opened = await this.policy.openTextFile(normalized)
    try {
      if (opened.stat.size > this.maxTextBytes) {
        throw new Error(`Text file exceeds size limit of ${this.maxTextBytes} bytes: ${normalized}`)
      }
      const bytes = Buffer.allocUnsafe(this.maxTextBytes + 1)
      let total = 0
      while (total <= this.maxTextBytes) {
        const { bytesRead } = await opened.handle.read(
          bytes,
          total,
          this.maxTextBytes + 1 - total,
          null,
        )
        if (bytesRead === 0) break
        total += bytesRead
      }
      if (total > this.maxTextBytes) {
        throw new Error(`Text file exceeds size limit of ${this.maxTextBytes} bytes: ${normalized}`)
      }
      try {
        return new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(0, total))
      } catch (error) {
        throw new Error(`Text file is not valid UTF-8: ${normalized}`, { cause: error })
      }
    } finally {
      await opened.handle.close()
    }
  }

  async saveText(
    relativePath: string,
    content: string,
    options: SaveTextOptions = {},
  ): Promise<SavedText> {
    const normalized = this.policy.normalize(relativePath)
    if (!isTextPath(normalized)) throw new Error(`Unsupported text file type: ${normalized}`)
    const bytes = Buffer.from(content, 'utf8')
    if (bytes.byteLength > this.maxTextBytes) {
      throw new Error(
        `Text content exceeds size limit of ${this.maxTextBytes} bytes: ${normalized}`,
      )
    }
    const target = await this.policy.prepareWrite(normalized)
    await atomicWriteFile(target.path, bytes, {
      validateBeforeRename: async () => {
        await this.policy.validateWriteTarget(target, 'before-rename')
        if (options.expectedSha256 === undefined) return
        if (options.expectedSha256 === null) {
          if (target.expectedTarget.kind !== 'absent') {
            throw new ProjectWriteConflictError(normalized)
          }
          return
        }
        if (target.expectedTarget.kind !== 'present') {
          throw new ProjectWriteConflictError(normalized)
        }
        const current = await this.readText(normalized)
        await this.policy.validateWriteTarget(target, 'before-rename')
        if (createHash('sha256').update(current, 'utf8').digest('hex') !== options.expectedSha256) {
          throw new ProjectWriteConflictError(normalized)
        }
      },
      validateAfterRename: (tempIdentity) =>
        this.policy.validateWriteTarget(target, 'after-rename', tempIdentity),
    })
    return { path: normalized, sha256: createHash('sha256').update(bytes).digest('hex') }
  }

  async deleteText(relativePath: string, options: DeleteTextOptions): Promise<void> {
    const normalized = this.policy.normalize(relativePath)
    if (!isTextPath(normalized)) throw new Error(`Unsupported text file type: ${normalized}`)
    if (!/^[a-f0-9]{64}$/.test(options.expectedSha256)) {
      throw new Error('Conditional delete requires a SHA-256 hash')
    }
    if (!/^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$/.test(options.transactionId)) {
      throw new Error('Invalid conditional delete transaction ID')
    }

    const targetPath = join(this.rootPath, ...normalized.split('/'))
    const quarantineName = `.wiswork-delete-${options.transactionId}-${createHash('sha256')
      .update(normalized)
      .digest('hex')
      .slice(0, 16)}`
    const parent = normalized.includes('/') ? normalized.slice(0, normalized.lastIndexOf('/')) : ''
    const quarantineRelative = parent ? `${parent}/${quarantineName}` : quarantineName
    const quarantinePath = join(dirname(targetPath), quarantineName)

    if (await pathExists(quarantinePath)) {
      await this.resumeConditionalDelete(
        normalized,
        targetPath,
        quarantineRelative,
        quarantinePath,
        options.expectedSha256,
      )
      return
    }

    let opened
    try {
      opened = await this.policy.openTextFile(normalized)
    } catch (error) {
      if (error instanceof Error && /does not exist/.test(error.message)) return
      throw error
    }
    let closed = false
    try {
      if ((await hashOpenedFile(opened.handle, this.maxTextBytes)) !== options.expectedSha256) {
        throw new ProjectWriteConflictError(normalized)
      }
      await this.policy.beforeConditionalDelete(opened.path)
      await rename(opened.path, quarantinePath)
      await this.policy.syncConditionalDeleteDirectory(dirname(targetPath))
      const quarantined = await lstat(quarantinePath)
      if (
        quarantined.isSymbolicLink() ||
        !quarantined.isFile() ||
        quarantined.dev !== opened.stat.dev ||
        quarantined.ino !== opened.stat.ino ||
        (await hashOpenedFile(opened.handle, this.maxTextBytes)) !== options.expectedSha256
      ) {
        await opened.handle.close()
        closed = true
        if (await restoreQuarantinedFile(quarantinePath, targetPath)) {
          await this.policy.syncConditionalDeleteDirectory(dirname(targetPath))
        }
        throw new ProjectWriteConflictError(normalized)
      }
      await opened.handle.close()
      closed = true
      const beforeDelete = await lstat(quarantinePath)
      if (beforeDelete.dev !== quarantined.dev || beforeDelete.ino !== quarantined.ino) {
        throw new ProjectWriteConflictError(normalized)
      }
      await unlink(quarantinePath)
      await this.policy.syncConditionalDeleteDirectory(dirname(targetPath))
    } catch (error) {
      if (!closed) await opened.handle.close().catch(() => undefined)
      throw error
    }
  }

  private async resumeConditionalDelete(
    normalized: string,
    targetPath: string,
    quarantineRelative: string,
    quarantinePath: string,
    expectedSha256: string,
  ): Promise<void> {
    const opened = await this.policy.openTextFile(quarantineRelative)
    let quarantineHash: string
    try {
      quarantineHash = await hashOpenedFile(opened.handle, this.maxTextBytes)
    } finally {
      await opened.handle.close()
    }
    const targetExists = await pathExists(targetPath)
    if (quarantineHash !== expectedSha256) {
      if (!targetExists) {
        await rename(quarantinePath, targetPath)
        await this.policy.syncConditionalDeleteDirectory(dirname(targetPath))
      }
      throw new ProjectWriteConflictError(normalized)
    }
    await unlink(quarantinePath)
    await this.policy.syncConditionalDeleteDirectory(dirname(targetPath))
    if (targetExists) throw new ProjectWriteConflictError(normalized)
  }
}

export async function openLatexProject(
  rootPath: string,
  options: LatexProjectOptions = {},
): Promise<LatexProject> {
  return LatexProject.open(rootPath, options)
}

export async function createLatexProject(
  rootPath: string,
  options: LatexProjectOptions = {},
): Promise<LatexProject> {
  try {
    await mkdir(rootPath)
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'EEXIST') {
      throw new Error(`LaTeX project target already exists: ${rootPath}`, { cause: error })
    }
    throw error
  }
  const policy = await ProjectPathPolicy.open(rootPath, options.pathHooks)
  const target = await policy.prepareWrite('main.tex')
  await atomicWriteFile(target.path, Buffer.from(ARTICLE_TEMPLATE, 'utf8'), {
    validateBeforeRename: () => policy.validateWriteTarget(target, 'before-rename'),
    validateAfterRename: (tempIdentity) =>
      policy.validateWriteTarget(target, 'after-rename', tempIdentity),
  })
  return LatexProject.open(rootPath, { ...options, savedMainFile: 'main.tex' })
}

function isTextPath(path: string): boolean {
  const lower = path.toLowerCase()
  return TEXT_FILENAMES.has(lower) || TEXT_EXTENSIONS.has(extname(lower))
}

async function hashOpenedFile(
  handle: import('node:fs/promises').FileHandle,
  maxBytes: number,
): Promise<string> {
  const stat = await handle.stat()
  if (stat.size > maxBytes) throw new Error(`Text file exceeds size limit of ${maxBytes} bytes`)
  const bytes = Buffer.allocUnsafe(maxBytes + 1)
  let total = 0
  while (total <= maxBytes) {
    const { bytesRead } = await handle.read(bytes, total, maxBytes + 1 - total, total)
    if (bytesRead === 0) break
    total += bytesRead
  }
  if (total > maxBytes) throw new Error(`Text file exceeds size limit of ${maxBytes} bytes`)
  return createHash('sha256').update(bytes.subarray(0, total)).digest('hex')
}

async function restoreQuarantinedFile(
  quarantinePath: string,
  targetPath: string,
): Promise<boolean> {
  if (await pathExists(targetPath)) return false
  await rename(quarantinePath, targetPath)
  return true
}

async function pathExists(path: string): Promise<boolean> {
  try {
    await lstat(path)
    return true
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return false
    throw error
  }
}
