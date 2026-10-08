#!/usr/bin/env node

import { createHash, randomBytes } from 'node:crypto'
import { execFile } from 'node:child_process'
import { createReadStream, createWriteStream } from 'node:fs'
import {
  chmod,
  copyFile,
  lstat,
  mkdir,
  open,
  readFile,
  readdir,
  rename,
  rm,
  stat,
} from 'node:fs/promises'
import { dirname, join, relative, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { promisify } from 'node:util'
import { pipeline } from 'node:stream/promises'
import { Transform } from 'node:stream'

const TOOL_ROOT = dirname(fileURLToPath(import.meta.url))
export const DEFAULT_MANIFEST_PATH = join(TOOL_ROOT, 'tectonic', 'manifest.json')
export const DEFAULT_CACHE_PATH = join(TOOL_ROOT, 'tectonic', '.cache')
export const INITIAL_HOSTS = Object.freeze(['github.com'])
export const REDIRECT_HOSTS = Object.freeze(['release-assets.githubusercontent.com'])
const execFileAsync = promisify(execFile)

class TectonicFetchError extends Error {
  constructor(stage, cause) {
    super(`Tectonic ${stage} failed`, { cause })
    this.stage = stage
  }
}

async function runAtStage(stage, task) {
  try {
    return await task()
  } catch (error) {
    throw new TectonicFetchError(stage, error)
  }
}

export function diagnosticForFailure(error) {
  const cause = error instanceof TectonicFetchError ? error.cause : error
  return {
    code: 'TECTONIC_FETCH_FAILED',
    stage: error instanceof TectonicFetchError ? error.stage : 'prepare',
    systemCode: systemCodeForFailure(cause),
  }
}

function systemCodeForFailure(error) {
  const visited = new Set()
  let current = error
  for (let depth = 0; depth < 6 && current && typeof current === 'object'; depth += 1) {
    if (visited.has(current)) break
    visited.add(current)
    const code = Object.getOwnPropertyDescriptor(current, 'code')?.value
    if (typeof code === 'string' && /^[A-Z0-9_]{1,32}$/.test(code)) return code
    current = Object.getOwnPropertyDescriptor(current, 'cause')?.value
  }
  return 'UNKNOWN'
}

export function parseArguments(argv) {
  let platform
  let manifestPath = DEFAULT_MANIFEST_PATH
  let cachePath = DEFAULT_CACHE_PATH
  let outputPath
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index]
    if (argument === '--platform') platform = argv[++index]
    else if (argument === '--manifest') manifestPath = resolve(argv[++index] ?? '')
    else if (argument === '--cache') cachePath = resolve(argv[++index] ?? '')
    else if (argument === '--output') outputPath = resolve(argv[++index] ?? '')
    else throw new Error(`unsupported argument: ${argument}`)
  }
  if (!platform || !/^[a-z0-9]+-[a-z0-9_]+$/.test(platform)) {
    throw new Error('--platform is required (for example, darwin-arm64)')
  }
  return Object.freeze({ platform, manifestPath, cachePath, outputPath })
}

export async function publishExecutable(sourcePath, outputPath) {
  const source = await lstat(sourcePath)
  if (!source.isFile() || source.isSymbolicLink()) throw new Error('verified executable is unsafe')
  await mkdir(dirname(outputPath), { recursive: true })
  const temporaryPath = `${outputPath}.${randomBytes(6).toString('hex')}.part`
  try {
    await copyFile(sourcePath, temporaryPath)
    await chmod(temporaryPath, 0o755)
    await rename(temporaryPath, outputPath)
  } catch (error) {
    await rm(temporaryPath, { force: true })
    throw error
  }
  return outputPath
}

export async function sha256File(path) {
  const hash = createHash('sha256')
  for await (const chunk of createReadStream(path)) hash.update(chunk)
  return hash.digest('hex')
}

export function selectPlatformAsset(manifest, platform) {
  if (!manifest || typeof manifest !== 'object' || Array.isArray(manifest)) {
    throw new Error('manifest must be an object')
  }
  if (manifest.schemaVersion !== 1 || !manifest.tectonic || typeof manifest.tectonic !== 'object') {
    throw new Error('manifest schema is invalid')
  }
  if (!/^\d+\.\d+\.\d+$/.test(manifest.tectonic.version ?? '')) {
    throw new Error('manifest version must be exact')
  }
  if (!Array.isArray(manifest.tectonic.assets)) throw new Error('manifest assets are invalid')
  const matches = manifest?.tectonic?.assets?.filter((asset) => asset.platform === platform) ?? []
  if (matches.length !== 1) throw new Error(`manifest must contain exactly one ${platform} asset`)
  const asset = matches[0]
  if (
    !/^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/.test(asset.id ?? '') ||
    !/^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/.test(asset.platform ?? '') ||
    !Number.isSafeInteger(asset.bytes) ||
    asset.bytes <= 0 ||
    !/^[a-f0-9]{64}$/.test(asset.sha256)
  ) {
    throw new Error('manifest asset integrity fields are invalid')
  }
  const url = new URL(asset.url)
  if (
    url.protocol !== 'https:' ||
    !INITIAL_HOSTS.includes(url.hostname) ||
    url.username ||
    url.password ||
    url.search ||
    url.hash ||
    url.pathname.toLowerCase().split('/').includes('latest')
  ) {
    throw new Error('manifest asset URL is not approved')
  }
  const supportedArchive =
    (asset.archive?.format === 'tar.gz' && asset.archive.executable === 'tectonic') ||
    (asset.archive?.format === 'zip' && asset.archive.executable === 'tectonic.exe')
  if (!supportedArchive) {
    throw new Error('manifest archive is invalid')
  }
  return Object.freeze({ ...asset, url: url.href })
}

async function fetchPinned(url, fetchImplementation, signal, redirects = 0) {
  if (redirects > 5) throw new Error('too many redirects')
  const response = await fetchImplementation(url, { redirect: 'manual', signal })
  if (response.status >= 300 && response.status < 400) {
    const location = response.headers.get('location')
    if (!location) throw new Error('redirect is missing Location')
    const next = new URL(location, url)
    if (next.protocol !== 'https:' || !REDIRECT_HOSTS.includes(next.hostname)) {
      throw new Error('redirect host is not approved')
    }
    return fetchPinned(next, fetchImplementation, signal, redirects + 1)
  }
  if (!response.ok || !response.body)
    throw new Error(`download failed with HTTP ${response.status}`)
  return response
}

export async function fetchVerifiedAsset(asset, targetPath, options = {}) {
  try {
    const existing = await lstat(targetPath)
    if (
      existing.isFile() &&
      !existing.isSymbolicLink() &&
      existing.size === asset.bytes &&
      (await sha256File(targetPath)) === asset.sha256
    ) {
      return Object.freeze({ path: targetPath, bytes: existing.size, sha256: asset.sha256 })
    }
  } catch (error) {
    if ((error.code ?? '') !== 'ENOENT') throw error
  }
  const maxAttempts = options.maxAttempts ?? 5
  const retryDelayMs = options.retryDelayMs ?? 1_000
  const sleepImplementation =
    options.sleepImplementation ??
    ((delay) => new Promise((resolveWait) => setTimeout(resolveWait, delay)))
  if (!Number.isSafeInteger(maxAttempts) || maxAttempts < 1 || maxAttempts > 5) {
    throw new Error('download attempt limit is invalid')
  }
  if (!Number.isSafeInteger(retryDelayMs) || retryDelayMs < 0 || retryDelayMs > 30_000) {
    throw new Error('download retry delay is invalid')
  }
  let lastError
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return await fetchVerifiedAssetOnce(asset, targetPath, options)
    } catch (error) {
      lastError = error
      if (attempt === maxAttempts) break
      const nextAttempt = attempt + 1
      options.onRetry?.(nextAttempt, systemCodeForFailure(error))
      const delay = Math.min(retryDelayMs * 2 ** (attempt - 1), 30_000)
      await sleepImplementation(delay)
    }
  }
  throw lastError
}

async function fetchVerifiedAssetOnce(asset, targetPath, options) {
  const fetchImplementation = options.fetchImplementation ?? globalThis.fetch
  const openImplementation = options.openImplementation ?? open
  const temporaryPath = `${targetPath}.${randomBytes(6).toString('hex')}.part`
  const controller = new AbortController()
  const timeout = setTimeout(
    () => controller.abort(new Error('download timed out')),
    options.timeoutMs ?? 300_000,
  )
  await mkdir(dirname(targetPath), { recursive: true })
  try {
    const response = await fetchPinned(asset.url, fetchImplementation, controller.signal)
    let received = 0
    const limiter = new Transform({
      transform(chunk, _encoding, callback) {
        received += Buffer.byteLength(chunk)
        if (received > asset.bytes) return callback(new Error('download exceeded expected size'))
        callback(null, chunk)
      },
    })
    await pipeline(response.body, limiter, createWriteStream(temporaryPath, { flags: 'wx' }), {
      signal: controller.signal,
    })
    const actual = await stat(temporaryPath)
    const digest = actual.size === asset.bytes ? await sha256File(temporaryPath) : null
    if (actual.size !== asset.bytes || digest !== asset.sha256) {
      throw new Error('downloaded asset failed integrity verification')
    }
    // Windows FlushFileBuffers requires GENERIC_WRITE access.
    const file = await openImplementation(temporaryPath, process.platform === 'win32' ? 'r+' : 'r')
    try {
      await file.sync()
    } finally {
      await file.close()
    }
    await rm(targetPath, { force: true })
    await rename(temporaryPath, targetPath)
    await syncDirectory(dirname(targetPath))
    return Object.freeze({ path: targetPath, bytes: actual.size, sha256: digest })
  } catch (error) {
    await rm(temporaryPath, { force: true })
    throw error
  } finally {
    clearTimeout(timeout)
  }
}

async function syncDirectory(path, openImplementation = open) {
  if (process.platform === 'win32') return
  const directory = await openImplementation(path, 'r')
  try {
    await directory.sync()
  } finally {
    await directory.close()
  }
}

export async function extractVerifiedTectonic(
  asset,
  archivePath,
  cachePath,
  version,
  options = {},
) {
  if (!/^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/.test(asset.id ?? '')) {
    throw new Error('asset ID is invalid')
  }
  await mkdir(cachePath, { recursive: true })
  const release = await acquireFilesystemLock(join(cachePath, `.${asset.id}.extract.lock`))
  try {
    return await extractVerifiedTectonicLocked(asset, archivePath, cachePath, version, options)
  } finally {
    await release()
  }
}

async function acquireFilesystemLock(path) {
  const deadline = Date.now() + 300_000
  const token = randomBytes(16).toString('hex')
  const recoveryPath = `${path}.recovery`
  while (true) {
    if (await pathExists(recoveryPath)) {
      const recoveryInfo = await stat(recoveryPath).catch(() => null)
      if (recoveryInfo && Date.now() - recoveryInfo.mtimeMs > 30_000) {
        await rm(recoveryPath, { force: true })
        continue
      }
      await new Promise((resolveWait) => setTimeout(resolveWait, 50))
      continue
    }
    let handle
    try {
      handle = await open(path, 'wx')
    } catch (error) {
      if ((error.code ?? '') !== 'EEXIST') throw error
      if (Date.now() >= deadline) {
        throw new Error('timed out waiting for extraction lock', { cause: error })
      }
      const info = await stat(path).catch(() => null)
      const owner = await readLockOwner(path)
      const invalidOwnerIsStale = info && !owner && Date.now() - info.mtimeMs > 30_000
      if ((owner && !lockOwnerIsAlive(owner.pid)) || invalidOwnerIsStale) {
        await recoverStaleLock(path, recoveryPath, owner?.token ?? null)
        continue
      }
      await new Promise((resolveWait) => setTimeout(resolveWait, 50))
      continue
    }
    try {
      await handle.writeFile(JSON.stringify({ pid: process.pid, token }), 'utf8')
      await handle.sync()
      return async () => {
        const owner = await readLockOwner(path)
        if (owner?.token === token) await rm(path, { force: true })
      }
    } catch (error) {
      await handle.close()
      handle = undefined
      const owner = await readLockOwner(path)
      if (owner?.token === token) await rm(path, { force: true })
      throw error
    } finally {
      await handle?.close()
    }
  }
}

async function pathExists(path) {
  return lstat(path).then(
    () => true,
    (error) => {
      if ((error.code ?? '') === 'ENOENT') return false
      throw error
    },
  )
}

async function readLockOwner(path) {
  try {
    const owner = JSON.parse(await readFile(path, 'utf8'))
    if (!Number.isSafeInteger(owner.pid) || owner.pid <= 0 || typeof owner.token !== 'string') {
      return null
    }
    return owner
  } catch {
    return null
  }
}

function lockOwnerIsAlive(owner) {
  try {
    process.kill(owner, 0)
    return true
  } catch (error) {
    return (error.code ?? '') !== 'ESRCH'
  }
}

async function recoverStaleLock(path, recoveryPath, expectedToken) {
  let guard
  try {
    guard = await open(recoveryPath, 'wx')
  } catch (error) {
    if ((error.code ?? '') === 'EEXIST') return
    throw error
  }
  try {
    await guard.writeFile(JSON.stringify({ pid: process.pid }), 'utf8')
    await guard.sync()
    const current = await readLockOwner(path)
    if ((current?.token ?? null) === expectedToken) await rm(path, { force: true })
  } finally {
    await guard.close()
    await rm(recoveryPath, { force: true })
  }
}

async function extractVerifiedTectonicLocked(asset, archivePath, cachePath, version, options = {}) {
  const execute = options.execFileImplementation ?? execFileAsync
  const renameImplementation = options.renameImplementation ?? rename
  const syncCacheDirectory = () => syncDirectory(cachePath, options.openImplementation)
  const targetDirectory = resolve(cachePath, asset.id)
  assertWithin(cachePath, targetDirectory)
  const targetExecutable = join(targetDirectory, asset.archive.executable)
  const staging = join(cachePath, `.${asset.id}.${randomBytes(6).toString('hex')}.extracting`)
  const backup = join(cachePath, `.${asset.id}.previous`)
  const targetExists = await lstat(targetDirectory).then(
    () => true,
    () => false,
  )
  const backupExists = await lstat(backup).then(
    () => true,
    () => false,
  )
  if (!targetExists && backupExists) {
    await rename(backup, targetDirectory)
    await syncCacheDirectory()
  } else if (targetExists && backupExists) {
    await rm(backup, { recursive: true, force: true })
    await syncCacheDirectory()
  }
  await mkdir(staging, { recursive: false })
  try {
    const archiveArgs =
      asset.archive.format === 'zip'
        ? { list: ['-tf', archivePath], extract: ['-xf', archivePath, '-C', staging] }
        : { list: ['-tzf', archivePath], extract: ['-xzf', archivePath, '-C', staging] }
    const listing = await execute('tar', archiveArgs.list)
    const entries = listing.stdout.split(/\r?\n/).filter(Boolean)
    if (entries.length !== 1 || entries[0] !== asset.archive.executable) {
      throw new Error('Tectonic archive layout is invalid')
    }
    await execute('tar', archiveArgs.extract)
    const extractedEntries = await readdir(staging)
    if (extractedEntries.length !== 1 || extractedEntries[0] !== asset.archive.executable) {
      throw new Error('Tectonic archive extracted unexpected files')
    }
    const executable = join(staging, asset.archive.executable)
    const info = await lstat(executable)
    if (!info.isFile() || info.isSymbolicLink()) throw new Error('Tectonic executable is unsafe')
    await chmod(executable, 0o755)
    const { stdout } = await execute(executable, ['--version'])
    if (parseVersion(stdout) !== version) throw new Error('Tectonic executable version mismatch')
    await rm(backup, { recursive: true, force: true })
    let hadTarget = false
    try {
      await lstat(targetDirectory)
      hadTarget = true
      await renameImplementation(targetDirectory, backup)
    } catch (error) {
      if ((error.code ?? '') !== 'ENOENT') throw error
    }
    try {
      await renameImplementation(staging, targetDirectory)
      await syncCacheDirectory()
    } catch (error) {
      if (hadTarget) {
        await rm(targetDirectory, { recursive: true, force: true })
        await rename(backup, targetDirectory)
        await syncCacheDirectory()
      }
      throw error
    }
    await rm(backup, { recursive: true, force: true })
    await syncCacheDirectory()
    return targetExecutable
  } catch (error) {
    await rm(staging, { recursive: true, force: true })
    throw error
  }
}

function assertWithin(root, candidate) {
  const relation = relative(resolve(root), candidate)
  if (relation.startsWith('..') || resolve(root, relation) !== candidate) {
    throw new Error('asset path escapes cache')
  }
}

function parseVersion(stdout) {
  const match = String(stdout)
    .trim()
    .match(/^tectonic\s+(\d+\.\d+\.\d+)$/i)
  return match?.[1]
}

export async function main(argv = process.argv.slice(2)) {
  const options = parseArguments(argv)
  const manifest = await runAtStage('manifest', async () =>
    JSON.parse(await readFile(options.manifestPath, 'utf8')),
  )
  const asset = await runAtStage('manifest', async () =>
    selectPlatformAsset(manifest, options.platform),
  )
  const archiveExtension = asset.archive.format === 'zip' ? '.zip' : '.tar.gz'
  const targetPath = resolve(options.cachePath, `${asset.id}${archiveExtension}`)
  assertWithin(options.cachePath, targetPath)
  const result = await runAtStage('download', () =>
    fetchVerifiedAsset(asset, targetPath, {
      onRetry: (attempt, systemCode) =>
        process.stderr.write(
          `${JSON.stringify({ code: 'TECTONIC_FETCH_RETRY', attempt, systemCode })}\n`,
        ),
    }),
  )
  const executablePath = await runAtStage('extract', () =>
    extractVerifiedTectonic(asset, result.path, options.cachePath, manifest.tectonic.version),
  )
  const publishedPath = options.outputPath
    ? await runAtStage('publish', () => publishExecutable(executablePath, options.outputPath))
    : executablePath
  process.stdout.write(
    `${JSON.stringify({ assetId: asset.id, bytes: result.bytes, executable: publishedPath })}\n`,
  )
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  main().catch((error) => {
    process.stderr.write(`${JSON.stringify(diagnosticForFailure(error))}\n`)
    process.exitCode = 1
  })
}
