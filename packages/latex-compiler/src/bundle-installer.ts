import { createHash, randomUUID } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { lstat, mkdir, open, readFile, rename, rm, stat } from 'node:fs/promises'
import { join } from 'node:path'
import { createHttpBundleDownload } from './download.js'
import { LatexCompilerError, type LatexCompilerErrorCode } from './errors.js'
import type { TectonicBundleAsset } from './manifest.js'

export type BundleInstallState =
  | { readonly state: 'missing' }
  | { readonly state: 'downloading'; readonly receivedBytes: number; readonly totalBytes: number }
  | { readonly state: 'ready'; readonly path: string; readonly bytes: number }
  | { readonly state: 'error'; readonly code: LatexCompilerErrorCode }

export interface BundleDownloadRequest {
  readonly url: string
  readonly destination: string
  readonly offset: number
  readonly expectedBytes: number
  readonly signal: AbortSignal
  readonly onBytes: (receivedBytes: number) => void
  readonly allowedRedirectHosts: readonly string[]
}

export type BundleDownload = (request: BundleDownloadRequest) => Promise<void>

export interface BundleInstallLog {
  readonly assetId: string
  readonly event: 'download-started' | 'download-ready' | 'download-cancelled' | 'download-error'
  readonly bytes?: number
  readonly code?: LatexCompilerErrorCode
}

export interface BundleInstallerOptions {
  readonly download?: BundleDownload
  readonly log?: (entry: BundleInstallLog) => void
  readonly syncDirectory?: (path: string) => Promise<void>
}

export interface InstalledBundle {
  readonly path: string
  readonly bytes: number
  readonly sha256: string
}

interface ActiveInstall {
  readonly promise: Promise<InstalledBundle>
  readonly controller: AbortController
  readonly owner: BundleInstaller
}

const activeInstalls = new Map<string, ActiveInstall>()

async function sha256(path: string): Promise<string> {
  const hash = createHash('sha256')
  for await (const chunk of createReadStream(path)) hash.update(chunk)
  return hash.digest('hex')
}

async function regularFileSize(path: string): Promise<number | null> {
  try {
    const stats = await lstat(path)
    return stats.isFile() && !stats.isSymbolicLink() ? stats.size : null
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null
    throw error
  }
}

export class BundleInstaller {
  readonly targetPath: string
  readonly temporaryPath: string
  readonly lockPath: string
  current: BundleInstallState = { state: 'missing' }

  private readonly download: BundleDownload
  private readonly log: (entry: BundleInstallLog) => void
  private readonly syncDirectory: (path: string) => Promise<void>
  private active: Promise<InstalledBundle> | null = null
  private controller: AbortController | null = null

  constructor(
    private readonly installDirectory: string,
    private readonly asset: TectonicBundleAsset,
    options: BundleInstallerOptions,
  ) {
    this.targetPath = join(installDirectory, `${asset.id}.tar`)
    this.temporaryPath = join(installDirectory, `.${asset.id}.tar.part`)
    this.lockPath = join(installDirectory, `.${asset.id}.install.lock`)
    this.download = options.download ?? createHttpBundleDownload()
    this.log = options.log ?? (() => undefined)
    this.syncDirectory = options.syncDirectory ?? syncDirectory
  }

  async status(): Promise<BundleInstallState> {
    const size = await regularFileSize(this.targetPath)
    if (size === null) {
      const shared = activeInstalls.get(this.targetPath)
      if (shared) return (this.current = shared.owner.current)
      return this.active ? this.current : { state: 'missing' }
    }
    if (size === this.asset.bytes && (await sha256(this.targetPath)) === this.asset.sha256) {
      return (this.current = { state: 'ready', path: this.targetPath, bytes: size })
    }
    return (this.current = { state: 'error', code: 'BUNDLE_INTEGRITY_FAILED' })
  }

  install(): Promise<InstalledBundle> {
    if (this.active) return this.active
    const shared = activeInstalls.get(this.targetPath)
    if (shared) {
      this.controller = shared.controller
      this.active = shared.promise
      this.current = shared.owner.current
      void shared.promise.then(
        () => {
          this.active = null
          this.controller = null
        },
        () => {
          this.active = null
          this.controller = null
        },
      )
      return shared.promise
    }
    const controller = new AbortController()
    this.controller = controller
    this.current = { state: 'downloading', receivedBytes: 0, totalBytes: this.asset.bytes }
    const operation = this.runInstall(controller.signal).finally(() => {
      if (activeInstalls.get(this.targetPath)?.promise === operation) {
        activeInstalls.delete(this.targetPath)
      }
      if (this.active === operation) {
        this.active = null
        this.controller = null
      }
    })
    this.active = operation
    activeInstalls.set(this.targetPath, { promise: operation, controller, owner: this })
    return operation
  }

  cancel(): void {
    this.controller?.abort(
      new LatexCompilerError('BUNDLE_DOWNLOAD_CANCELLED', 'Bundle download cancelled'),
    )
  }

  private async runInstall(signal: AbortSignal): Promise<InstalledBundle> {
    await mkdir(this.installDirectory, { recursive: true })
    const release = await this.acquireLock(signal)
    try {
      return await this.runInstallLocked(signal)
    } finally {
      await release()
    }
  }

  private async acquireLock(signal: AbortSignal): Promise<() => Promise<void>> {
    const token = randomUUID()
    const recoveryPath = `${this.lockPath}.recovery`
    const deadline = Date.now() + 300_000
    while (true) {
      signal.throwIfAborted()
      if (Date.now() >= deadline) {
        throw new LatexCompilerError(
          'BUNDLE_INSTALL_FAILED',
          'Timed out waiting for bundle installation lock',
        )
      }
      if (await pathExists(recoveryPath)) {
        const recoveryInfo = await stat(recoveryPath).catch(() => null)
        if (recoveryInfo && Date.now() - recoveryInfo.mtimeMs > 30_000) {
          await rm(recoveryPath, { force: true })
          continue
        }
        await waitForRetry(signal)
        continue
      }
      let handle
      try {
        handle = await open(this.lockPath, 'wx')
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error
        const lockInfo = await stat(this.lockPath).catch(() => null)
        if (!lockInfo) continue
        const owner = await readLockOwner(this.lockPath)
        const invalidOwnerIsStale = !owner && Date.now() - lockInfo.mtimeMs > 30_000
        if ((owner && !lockOwnerIsAlive(owner.pid)) || invalidOwnerIsStale) {
          await recoverStaleLock(this.lockPath, recoveryPath, owner?.token ?? null)
          continue
        }
        await waitForRetry(signal)
        continue
      }
      try {
        await handle.writeFile(JSON.stringify({ pid: process.pid, token }), 'utf8')
        await handle.sync()
        return async () => {
          const owner = await readLockOwner(this.lockPath)
          if (owner?.token === token) await rm(this.lockPath, { force: true })
        }
      } catch (error) {
        await handle.close()
        handle = undefined
        const owner = await readLockOwner(this.lockPath)
        if (owner?.token === token) await rm(this.lockPath, { force: true })
        throw error
      } finally {
        await handle?.close()
      }
    }
  }

  private async runInstallLocked(signal: AbortSignal): Promise<InstalledBundle> {
    const existing = await this.status()
    if (existing.state === 'ready') {
      return { path: existing.path, bytes: existing.bytes, sha256: this.asset.sha256 }
    }
    let offset = (await regularFileSize(this.temporaryPath)) ?? 0
    if (offset > this.asset.bytes) {
      await rm(this.temporaryPath, { force: true })
      offset = 0
    }
    if (offset === this.asset.bytes) {
      if ((await sha256(this.temporaryPath)) === this.asset.sha256) {
        return this.publishVerifiedTemporary(offset, this.asset.sha256)
      }
      await rm(this.temporaryPath, { force: true })
      offset = 0
    }
    this.current = { state: 'downloading', receivedBytes: offset, totalBytes: this.asset.bytes }
    this.log({ assetId: this.asset.id, event: 'download-started', bytes: offset })
    try {
      await this.download({
        url: this.asset.url,
        destination: this.temporaryPath,
        offset,
        expectedBytes: this.asset.bytes,
        signal,
        allowedRedirectHosts: ['data1.fullyjustified.net'],
        onBytes: (receivedBytes) => {
          this.current = {
            state: 'downloading',
            receivedBytes: Math.min(this.asset.bytes, offset + receivedBytes),
            totalBytes: this.asset.bytes,
          }
        },
      })
      signal.throwIfAborted()
    } catch (error) {
      if (signal.aborted) {
        const cancelled = new LatexCompilerError(
          'BUNDLE_DOWNLOAD_CANCELLED',
          'Bundle download cancelled',
        )
        this.current = { state: 'missing' }
        this.log({ assetId: this.asset.id, event: 'download-cancelled', code: cancelled.code })
        throw cancelled
      }
      const wrapped = new LatexCompilerError(
        'BUNDLE_DOWNLOAD_FAILED',
        'Bundle download failed',
        error,
      )
      this.current = { state: 'error', code: wrapped.code }
      this.log({ assetId: this.asset.id, event: 'download-error', code: wrapped.code })
      throw wrapped
    }
    const size = await regularFileSize(this.temporaryPath)
    const digest = size === this.asset.bytes ? await sha256(this.temporaryPath) : null
    if (size !== this.asset.bytes || digest !== this.asset.sha256) {
      await rm(this.temporaryPath, { force: true })
      const error = new LatexCompilerError(
        'BUNDLE_INTEGRITY_FAILED',
        'Bundle integrity verification failed',
      )
      this.current = { state: 'error', code: error.code }
      this.log({
        assetId: this.asset.id,
        event: 'download-error',
        bytes: size ?? 0,
        code: error.code,
      })
      throw error
    }
    return this.publishVerifiedTemporary(size, digest)
  }

  private async publishVerifiedTemporary(size: number, digest: string): Promise<InstalledBundle> {
    try {
      const handle = await open(this.temporaryPath, 'r')
      try {
        await handle.sync()
      } finally {
        await handle.close()
      }
      await rm(this.targetPath, { force: true })
      await rename(this.temporaryPath, this.targetPath)
      await this.syncDirectory(this.installDirectory)
    } catch (error) {
      const committedSize = await regularFileSize(this.targetPath).catch(() => null)
      if (
        committedSize === this.asset.bytes &&
        (await sha256(this.targetPath).catch(() => null)) === this.asset.sha256
      ) {
        this.current = { state: 'ready', path: this.targetPath, bytes: committedSize }
        this.log({ assetId: this.asset.id, event: 'download-ready', bytes: committedSize })
        return { path: this.targetPath, bytes: committedSize, sha256: this.asset.sha256 }
      }
      const wrapped = new LatexCompilerError(
        'BUNDLE_INSTALL_FAILED',
        'Bundle installation failed',
        error,
      )
      this.current = { state: 'error', code: wrapped.code }
      this.log({ assetId: this.asset.id, event: 'download-error', code: wrapped.code })
      throw wrapped
    }
    this.current = { state: 'ready', path: this.targetPath, bytes: size }
    this.log({ assetId: this.asset.id, event: 'download-ready', bytes: size })
    return { path: this.targetPath, bytes: size, sha256: digest }
  }
}

async function pathExists(path: string): Promise<boolean> {
  return lstat(path).then(
    () => true,
    (error: NodeJS.ErrnoException) => {
      if (error.code === 'ENOENT') return false
      throw error
    },
  )
}

async function waitForRetry(signal: AbortSignal): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const settled = () => signal.removeEventListener('abort', aborted)
    const timer = setTimeout(() => {
      settled()
      resolve()
    }, 50)
    const aborted = () => {
      clearTimeout(timer)
      settled()
      reject(signal.reason)
    }
    signal.addEventListener('abort', aborted, { once: true })
  })
}

async function recoverStaleLock(
  lockPath: string,
  recoveryPath: string,
  expectedToken: string | null,
): Promise<void> {
  let guard
  try {
    guard = await open(recoveryPath, 'wx')
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'EEXIST') return
    throw error
  }
  try {
    await guard.writeFile(JSON.stringify({ pid: process.pid }), 'utf8')
    await guard.sync()
    const current = await readLockOwner(lockPath)
    if ((current?.token ?? null) === expectedToken) await rm(lockPath, { force: true })
  } finally {
    await guard.close()
    await rm(recoveryPath, { force: true })
  }
}

async function syncDirectory(path: string): Promise<void> {
  const directory = await open(path, 'r')
  try {
    await directory.sync()
  } finally {
    await directory.close()
  }
}

interface LockOwner {
  readonly pid: number
  readonly token: string
}

async function readLockOwner(path: string): Promise<LockOwner | null> {
  try {
    const value = JSON.parse(await readFile(path, 'utf8')) as Partial<LockOwner>
    if (
      !Number.isSafeInteger(value.pid) ||
      (value.pid ?? 0) <= 0 ||
      typeof value.token !== 'string'
    ) {
      return null
    }
    return { pid: value.pid!, token: value.token }
  } catch {
    return null
  }
}

function lockOwnerIsAlive(owner: number): boolean {
  try {
    process.kill(owner, 0)
    return true
  } catch (error) {
    return (error as NodeJS.ErrnoException).code !== 'ESRCH'
  }
}
