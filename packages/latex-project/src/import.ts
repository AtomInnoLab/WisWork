import { createWriteStream } from 'node:fs'
import { lstat, mkdir, mkdtemp, open, rename, rm, stat } from 'node:fs/promises'
import { basename, dirname, isAbsolute, join } from 'node:path'
import { Readable, Transform } from 'node:stream'
import { pipeline } from 'node:stream/promises'
import yauzl, { type Entry, type ZipFile } from 'yauzl'

const WINDOWS_ABSOLUTE = /^[a-zA-Z]:[\\/]/
const UNIX_FILE_TYPE = 0o170000
const UNIX_REGULAR = 0o100000
const UNIX_DIRECTORY = 0o040000

export interface ImportLimits {
  maxArchiveBytes: number
  maxEntries: number
  maxDirectoryDepth: number
  maxFileBytes: number
  maxTotalBytes: number
  maxCompressionRatio: number
}

export const DEFAULT_IMPORT_LIMITS: Readonly<ImportLimits> = Object.freeze({
  maxArchiveBytes: 64 * 1024 * 1024,
  maxEntries: 2_000,
  maxDirectoryDepth: 16,
  maxFileBytes: 16 * 1024 * 1024,
  maxTotalBytes: 128 * 1024 * 1024,
  maxCompressionRatio: 100,
})

export interface ImportResult {
  rootPath: string
  entryCount: number
  totalUncompressedBytes: number
}

/**
 * Extracts to a private sibling and publishes only after every entry validates. ZIP has no
 * standard hard-link metadata. Unix entries therefore permit only regular files/directories and
 * reject every other representable file type, including symbolic links and device nodes.
 */
export async function importLatexProject(
  archivePath: string,
  targetPath: string,
  overrides: Partial<ImportLimits> = {},
): Promise<ImportResult> {
  const limits = resolveLimits(overrides)
  const archiveStat = await stat(archivePath)
  if (!archiveStat.isFile()) throw new Error('ZIP archive is not a regular file')
  if (archiveStat.size > limits.maxArchiveBytes) throw new Error('ZIP archive exceeds size limit')
  await assertMissing(targetPath)

  const parent = dirname(targetPath)
  const leaf = basename(targetPath)
  if (!leaf || leaf === '.' || leaf === '..') throw new Error('Import target must have a file name')
  const parentStat = await lstat(parent)
  if (!parentStat.isDirectory() || parentStat.isSymbolicLink()) {
    throw new Error('Import parent is not a safe directory')
  }
  const temporary = await mkdtemp(join(parent, `.${leaf}.import-`))
  try {
    const result = await extractArchive(archivePath, temporary, limits)
    await assertMissing(targetPath)
    try {
      await rename(temporary, targetPath)
    } catch (error) {
      if (['EEXIST', 'ENOTEMPTY'].includes((error as NodeJS.ErrnoException).code ?? '')) {
        throw new Error(`LaTeX project target already exists: ${targetPath}`, { cause: error })
      }
      throw error
    }
    return { rootPath: targetPath, ...result }
  } catch (error) {
    await rm(temporary, { recursive: true, force: true })
    throw error
  }
}

function resolveLimits(overrides: Partial<ImportLimits>): ImportLimits {
  const limits = { ...DEFAULT_IMPORT_LIMITS, ...overrides }
  for (const [name, value] of Object.entries(limits)) {
    const validInteger = name === 'maxCompressionRatio' || Number.isSafeInteger(value)
    if (!Number.isFinite(value) || value <= 0 || !validInteger) {
      throw new Error(`${name} must be a positive limit`)
    }
  }
  return limits
}

async function extractArchive(archivePath: string, root: string, limits: ImportLimits) {
  const zip = await openZip(archivePath)
  const targets = new Map<string, 'file' | 'directory'>()
  let entryCount = 0
  let totalUncompressedBytes = 0
  try {
    while (true) {
      const entry = await nextEntry(zip)
      if (!entry) break
      entryCount += 1
      if (entryCount > limits.maxEntries) throw new Error('ZIP entry count exceeds limit')
      if ((entry.generalPurposeBitFlag & 1) !== 0) {
        throw new Error('Encrypted ZIP entries are not allowed')
      }

      const rawName = decodeEntryName(entry)
      const normalized = normalizeEntryName(rawName)
      const kind = entryKind(entry, rawName)
      assertUniqueTarget(targets, normalized, kind)
      const depth = normalized.split('/').length - (kind === 'file' ? 1 : 0)
      if (depth > limits.maxDirectoryDepth) throw new Error('ZIP directory depth exceeds limit')
      const ratio =
        entry.compressedSize === 0
          ? entry.uncompressedSize === 0
            ? 0
            : Infinity
          : entry.uncompressedSize / entry.compressedSize
      if (ratio > limits.maxCompressionRatio) {
        throw new Error('ZIP entry compression ratio exceeds limit')
      }
      if (kind === 'file' && entry.uncompressedSize > limits.maxFileBytes) {
        throw new Error('ZIP entry file size exceeds limit')
      }
      totalUncompressedBytes += entry.uncompressedSize
      if (
        !Number.isSafeInteger(totalUncompressedBytes) ||
        totalUncompressedBytes > limits.maxTotalBytes
      ) {
        throw new Error('ZIP total uncompressed size exceeds limit')
      }

      const destination = join(root, ...normalized.split('/'))
      if (kind === 'directory') {
        await mkdir(destination, { recursive: true })
      } else {
        await mkdir(dirname(destination), { recursive: true })
        await extractFile(zip, entry, destination, limits.maxFileBytes)
      }
    }
    return { entryCount, totalUncompressedBytes }
  } finally {
    zip.close()
  }
}

function openZip(path: string): Promise<ZipFile> {
  return new Promise((resolve, reject) => {
    yauzl.open(
      path,
      {
        lazyEntries: true,
        autoClose: false,
        decodeStrings: false,
        validateEntrySizes: true,
        strictFileNames: true,
      },
      (error, zip) => {
        if (error || !zip) reject(error ?? new Error('Unable to open ZIP'))
        else resolve(zip)
      },
    )
  })
}

function nextEntry(zip: ZipFile): Promise<Entry | undefined> {
  return new Promise((resolve, reject) => {
    const cleanup = () => {
      zip.off('entry', onEntry)
      zip.off('end', onEnd)
      zip.off('error', onError)
    }
    const onEntry = (entry: Entry) => {
      cleanup()
      resolve(entry)
    }
    const onEnd = () => {
      cleanup()
      resolve(undefined)
    }
    const onError = (error: Error) => {
      cleanup()
      reject(error)
    }
    zip.once('entry', onEntry)
    zip.once('end', onEnd)
    zip.once('error', onError)
    zip.readEntry()
  })
}

function decodeEntryName(entry: Entry): string {
  const raw = entry.fileName as unknown
  if (!Buffer.isBuffer(raw)) throw new Error('ZIP parser did not expose a raw entry name')
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(raw)
  } catch (error) {
    throw new Error('ZIP entry name is not valid UTF-8', { cause: error })
  }
}

function normalizeEntryName(name: string): string {
  if (name.includes('\0')) throw new Error('ZIP entry names must not contain NUL bytes')
  if (
    !name ||
    name.startsWith('/') ||
    name.startsWith('\\') ||
    isAbsolute(name) ||
    WINDOWS_ABSOLUTE.test(name)
  ) {
    throw new Error('ZIP entry names must be relative')
  }
  if (name.includes('\\')) throw new Error('ZIP entry names must use forward slashes')
  const segments = name.split('/').filter((segment) => segment !== '' && segment !== '.')
  if (segments.includes('..')) throw new Error('ZIP path traversal is not allowed')
  if (segments.length === 0) throw new Error('ZIP entry name is empty after normalization')
  return segments.join('/')
}

function entryKind(entry: Entry, name: string): 'file' | 'directory' {
  const creator = entry.versionMadeBy >>> 8
  const attributes = entry.externalFileAttributes >>> 0
  const slashDirectory = name.endsWith('/')
  if (creator === 3) {
    const type = (attributes >>> 16) & UNIX_FILE_TYPE
    if (type === UNIX_DIRECTORY && slashDirectory) return 'directory'
    if (type === UNIX_REGULAR && !slashDirectory) return 'file'
    throw new Error('ZIP Unix entry type is not a matching regular file or directory')
  }
  const dosDirectory = (attributes & 0x10) !== 0
  if (dosDirectory !== slashDirectory) {
    throw new Error('ZIP entry directory metadata is inconsistent')
  }
  return slashDirectory ? 'directory' : 'file'
}

function assertUniqueTarget(
  targets: Map<string, 'file' | 'directory'>,
  path: string,
  kind: 'file' | 'directory',
): void {
  if (targets.has(path)) throw new Error(`ZIP contains duplicate normalized target: ${path}`)
  const segments = path.split('/')
  for (let index = 1; index < segments.length; index += 1) {
    if (targets.get(segments.slice(0, index).join('/')) === 'file') {
      throw new Error(`ZIP target has a file parent: ${path}`)
    }
  }
  if (
    kind === 'file' &&
    [...targets.keys()].some((candidate) => candidate.startsWith(`${path}/`))
  ) {
    throw new Error(`ZIP file target collides with a directory: ${path}`)
  }
  targets.set(path, kind)
}

async function extractFile(
  zip: ZipFile,
  entry: Entry,
  destination: string,
  maxBytes: number,
): Promise<void> {
  const stream = await new Promise<NodeJS.ReadableStream>((resolve, reject) => {
    zip.openReadStream(entry, (error, value) => {
      if (error || !value) reject(error ?? new Error('Unable to read ZIP entry'))
      else resolve(value)
    })
  })
  let written = 0
  const limiter = new Transform({
    transform(chunk: Buffer, _encoding, callback) {
      written += chunk.length
      if (written > maxBytes) {
        callback(new Error('ZIP entry exceeded file size limit while extracting'))
      } else {
        callback(null, chunk)
      }
    },
  })
  const output = createWriteStream(destination, { flags: 'wx', mode: 0o600 })
  await pipeline(stream as Readable, limiter, output)
  if (written !== entry.uncompressedSize) {
    throw new Error('ZIP entry size changed during extraction')
  }

  const handle = await open(destination, 'r+')
  try {
    await handle.sync()
  } finally {
    await handle.close()
  }
}

async function assertMissing(path: string): Promise<void> {
  try {
    await lstat(path)
    throw new Error(`LaTeX project target already exists: ${path}`)
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
  }
}
