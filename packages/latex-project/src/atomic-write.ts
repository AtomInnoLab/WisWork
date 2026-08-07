import { randomBytes } from 'node:crypto'
import { constants } from 'node:fs'
import { open, rename, stat, unlink } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'

export interface AtomicFileIdentity {
  dev: number
  ino: number
}

export interface AtomicWriteOptions {
  validateBeforeRename?: () => Promise<void>
  validateAfterRename?: (tempIdentity: AtomicFileIdentity) => Promise<void>
  syncDirectory?: (directory: string) => Promise<void>
}

export class AtomicWriteCommittedError extends Error {
  readonly committed = true
  readonly code: string

  constructor(kind: 'validation' | 'durability', cause: unknown) {
    const detail = cause instanceof Error ? cause.message : String(cause)
    super(`Atomic write committed but post-rename ${kind} failed: ${detail}`, { cause })
    this.name = 'AtomicWriteCommittedError'
    this.code =
      kind === 'validation'
        ? 'LATEX_ATOMIC_WRITE_COMMITTED_VALIDATION_FAILED'
        : 'LATEX_ATOMIC_WRITE_COMMITTED_DURABILITY_FAILED'
  }
}

export async function atomicWriteFile(
  filePath: string,
  data: Uint8Array,
  options: AtomicWriteOptions = {},
): Promise<void> {
  const directory = dirname(filePath)
  const tempPath = join(
    directory,
    `.${basename(filePath)}.${process.pid}.${randomBytes(6).toString('hex')}.tmp`,
  )
  let mode: number | undefined
  try {
    mode = (await stat(filePath)).mode
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
  }

  const noFollow = 'O_NOFOLLOW' in constants ? constants.O_NOFOLLOW : 0
  const handle = await open(
    tempPath,
    constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | noFollow,
    mode ?? 0o600,
  )
  let renamed = false
  try {
    await handle.writeFile(data)
    await handle.sync()
    const tempStat = await handle.stat()
    const tempIdentity = { dev: tempStat.dev, ino: tempStat.ino }
    await handle.close()
    await options.validateBeforeRename?.()
    await rename(tempPath, filePath)
    renamed = true

    let validationError: unknown
    try {
      await options.validateAfterRename?.(tempIdentity)
    } catch (error) {
      validationError = error
    }
    let syncError: unknown
    try {
      await (options.syncDirectory ?? syncParentDirectory)(directory)
    } catch (error) {
      syncError = error
    }
    if (validationError) {
      const cause = syncError
        ? new AggregateError([validationError, syncError], 'Validation and directory sync failed')
        : validationError
      throw new AtomicWriteCommittedError('validation', cause)
    }
    if (syncError) throw new AtomicWriteCommittedError('durability', syncError)
  } catch (error) {
    await handle.close().catch(() => undefined)
    if (!renamed) await unlink(tempPath).catch(() => undefined)
    throw error
  }
}

async function syncParentDirectory(directory: string): Promise<void> {
  let handle
  try {
    handle = await open(directory, constants.O_RDONLY)
    await handle.sync()
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code ?? ''
    const unsupportedOnWindows = new Set(['EACCES', 'EINVAL', 'EISDIR', 'ENOTSUP', 'EPERM'])
    if (process.platform !== 'win32' || !unsupportedOnWindows.has(code)) throw error
  } finally {
    await handle?.close().catch(() => undefined)
  }
}
