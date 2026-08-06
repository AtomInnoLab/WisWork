import { createHash } from 'node:crypto'
import { mkdir, readdir, realpath } from 'node:fs/promises'
import { extname } from 'node:path'
import { atomicWriteFile } from './atomic-write.js'
import { discoverMainFile } from './main-file.js'
import { ProjectPathPolicy } from './path-policy.js'
import {
  DEFAULT_MAX_TEXT_BYTES,
  type LatexProjectOptions,
  type MainFileDiscovery,
  type SavedText,
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

  async saveText(relativePath: string, content: string): Promise<SavedText> {
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
      validateBeforeRename: () => this.policy.validateWriteTarget(target, 'before-rename'),
      validateAfterRename: (tempIdentity) =>
        this.policy.validateWriteTarget(target, 'after-rename', tempIdentity),
    })
    return { path: normalized, sha256: createHash('sha256').update(bytes).digest('hex') }
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
