import { execFile as nodeExecFile } from 'node:child_process'
import { createHash, randomBytes } from 'node:crypto'
import { createReadStream, createWriteStream } from 'node:fs'
import {
  chmod,
  lstat,
  mkdir,
  open,
  readFile,
  readdir,
  realpath,
  rename,
  rmdir,
  rm,
  statfs,
} from 'node:fs/promises'
import { isAbsolute, relative, resolve, sep } from 'node:path'
import { Readable, Transform } from 'node:stream'
import { pipeline } from 'node:stream/promises'
import { promisify } from 'node:util'
import { inflateRawSync } from 'node:zlib'
import { extract as extractTar, type ReadEntry } from 'tar'

const PINNED_VERSION = '0.147.0'
const PINNED_VERSION_OUTPUT = `codex-app-server ${PINNED_VERSION}`
const MAX_VERSION_OUTPUT = 65_536
const DEFAULT_DOWNLOAD_TIMEOUT_MS = 10 * 60_000
const DEFAULT_PROBE_TIMEOUT_MS = 5_000
const DEFAULT_LOCK_TIMEOUT_MS = 30_000
const CAPACITY_MARGIN_BYTES = 64 * 1024 * 1024
const PRIMARY_DOWNLOAD_HOST = 'downloads.wiswork.com'
const FALLBACK_DOWNLOAD_HOST = 'github.com'
const MANIFEST_SOURCE_PREFIX = 'https://github.com/openai/codex/releases/download/rust-v0.147.0/'
const MANIFEST_MIRROR_PREFIX = 'https://downloads.wiswork.com/components/codex/0.147.0/'
const LICENSE_SOURCE = 'https://github.com/openai/codex/blob/rust-v0.147.0/LICENSE'
const execFile = promisify(nodeExecFile)

export type CodexComponentFileMode = 'executable' | 'data'

export interface CodexComponentFileManifest {
  readonly path: string
  readonly bytes: number
  readonly sha256: string
  readonly mode: CodexComponentFileMode
  /** Whether this verified archive entry is retained in the installed component. */
  readonly install: boolean
}

export interface CodexComponentAssetManifest {
  readonly id: string
  readonly platform: 'darwin' | 'win32'
  readonly arch: 'arm64' | 'x64'
  readonly target: 'aarch64-apple-darwin' | 'x86_64-apple-darwin' | 'x86_64-pc-windows-msvc'
  readonly primaryUrl: string
  readonly fallbackUrl: string
  readonly bytes: number
  readonly sha256: string
  readonly archive: { readonly format: 'tar.gz' | 'zip'; readonly maxExtractedBytes: number }
  readonly trust:
    | {
        readonly policy: 'macos'
        readonly teamIdentifier: string
        readonly requireNotarization: true
      }
    | { readonly policy: 'windows'; readonly publisher: string }
  readonly layout: {
    readonly entrypoint: string
    readonly directories: readonly string[]
    readonly files: readonly CodexComponentFileManifest[]
  }
}

export interface CodexComponentManifest {
  readonly schemaVersion: 1
  readonly component: {
    readonly version: '0.147.0'
    readonly license: {
      readonly spdx: 'Apache-2.0'
      readonly sourceUrl: string
    }
    readonly assets: readonly CodexComponentAssetManifest[]
  }
}

export type EnhancedModeComponentState = 'unsupported' | 'missing' | 'ready' | 'invalid'

export interface EnhancedModeComponentStatus {
  readonly state: EnhancedModeComponentState
  readonly supported: boolean
  readonly version: '0.147.0'
}

export interface InstalledEnhancedModeComponent {
  readonly executablePath: string
  readonly version: '0.147.0'
  readonly platform: 'darwin' | 'win32'
  readonly arch: 'arm64' | 'x64'
}

export interface ComponentProbeOptions {
  readonly signal?: AbortSignal
  readonly timeoutMs?: number
}

export type ComponentVersionProbe = (
  executablePath: string,
  options: ComponentProbeOptions,
) => Promise<string>

export type ComponentPlatformTrustVerifier = (
  executablePaths: readonly string[],
  policy: CodexComponentAssetManifest['trust'],
  options: ComponentProbeOptions,
) => Promise<void>

export type ComponentCapacityCheck = (cacheRoot: string, requiredBytes: number) => Promise<void>
export type ComponentInstallPhase =
  'download' | 'digest' | 'signature' | 'promote' | 'update' | 'remove'
export type ComponentPhaseEvent = Readonly<{
  phase: ComponentInstallPhase
  outcome: 'started' | 'succeeded' | 'failed'
}>

export interface EnhancedModeComponentManagerOptions {
  readonly cacheRoot: string
  readonly manifest: CodexComponentManifest | unknown
  readonly platform?: NodeJS.Platform | string
  readonly arch?: NodeJS.Architecture | string
  readonly fetchImplementation?: typeof fetch
  readonly probeVersion?: ComponentVersionProbe
  readonly verifyPlatformTrust?: ComponentPlatformTrustVerifier
  readonly checkAvailableCapacity?: ComponentCapacityCheck
  readonly downloadTimeoutMs?: number
  readonly probeTimeoutMs?: number
  readonly lockTimeoutMs?: number
  readonly lockStaleMs?: number
  readonly onPhase?: (event: ComponentPhaseEvent) => void
}

export class EnhancedModeComponentError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'EnhancedModeComponentError'
  }
}

type UnknownRecord = Record<string, unknown>

function fail(code: string): never {
  throw new EnhancedModeComponentError(code)
}

const MACOS_TRUST_DIAGNOSTIC_CODES = new Set([
  'enhanced_mode_macos_codesign_failed',
  'enhanced_mode_macos_codesign_details_failed',
  'enhanced_mode_macos_team_identifier_mismatch',
  'enhanced_mode_macos_notarization_failed',
])

function failPlatformTrust(error: unknown, signal?: AbortSignal): never {
  if (error instanceof EnhancedModeComponentError && MACOS_TRUST_DIAGNOSTIC_CODES.has(error.code)) {
    throw error
  }
  fail(signal?.aborted ? 'enhanced_mode_cancelled' : 'enhanced_mode_platform_trust_failed')
}

function isRecord(value: unknown): value is UnknownRecord {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function hasOnlyKeys(value: UnknownRecord, allowed: readonly string[]): boolean {
  const set = new Set(allowed)
  return Object.keys(value).every((key) => set.has(key))
}

function exactPositiveInteger(value: unknown): value is number {
  return Number.isSafeInteger(value) && (value as number) > 0
}

function exactSha256(value: unknown): value is string {
  return typeof value === 'string' && /^[a-f0-9]{64}$/.test(value)
}

function safeRelativePath(value: unknown): value is string {
  if (typeof value !== 'string' || value === '' || value.length > 256) return false
  if (value.includes('\\') || value.includes('\0') || value.startsWith('/')) return false
  const segments = value.replace(/\/$/, '').split('/')
  return segments.every(
    (segment) => segment !== '' && segment !== '.' && segment !== '..' && segment.length <= 128,
  )
}

const WINDOWS_RESERVED_NAME = /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i

function canonicalTargetPath(
  path: string,
  platform: string,
  errorCode = 'enhanced_mode_manifest_invalid',
): string {
  if (!safeRelativePath(path)) fail(errorCode)
  if (platform !== 'win32') return path
  const canonical: string[] = []
  for (const segment of path.split('/')) {
    const trimmed = segment.replace(/[ .]+$/u, '')
    const deviceBase = trimmed.split('.')[0]!.replace(/[ .]+$/u, '')
    const hasControlCharacter = [...trimmed].some((character) => character.charCodeAt(0) <= 0x1f)
    if (
      trimmed !== segment ||
      trimmed === '' ||
      /[<>:"|?*]/u.test(trimmed) ||
      hasControlCharacter ||
      WINDOWS_RESERVED_NAME.test(deviceBase)
    ) {
      fail(errorCode)
    }
    canonical.push(trimmed.toLowerCase())
  }
  return canonical.join('/')
}

function parseManifestFile(value: unknown): CodexComponentFileManifest {
  if (
    !isRecord(value) ||
    !hasOnlyKeys(value, ['path', 'bytes', 'sha256', 'mode', 'install']) ||
    !safeRelativePath(value.path) ||
    !exactPositiveInteger(value.bytes) ||
    !exactSha256(value.sha256) ||
    (value.mode !== 'executable' && value.mode !== 'data') ||
    typeof value.install !== 'boolean'
  ) {
    fail('enhanced_mode_manifest_invalid')
  }
  return value as unknown as CodexComponentFileManifest
}

function parseManifestAsset(value: unknown): CodexComponentAssetManifest {
  if (
    !isRecord(value) ||
    !hasOnlyKeys(value, [
      'id',
      'platform',
      'arch',
      'target',
      'primaryUrl',
      'fallbackUrl',
      'bytes',
      'sha256',
      'archive',
      'trust',
      'layout',
    ]) ||
    typeof value.id !== 'string' ||
    !/^codex-app-server-0\.147\.0-(darwin-(?:arm64|x64)|win32-x64)$/.test(value.id) ||
    (value.platform !== 'darwin' && value.platform !== 'win32') ||
    (value.arch !== 'arm64' && value.arch !== 'x64') ||
    !['aarch64-apple-darwin', 'x86_64-apple-darwin', 'x86_64-pc-windows-msvc'].includes(
      value.target as string,
    ) ||
    !exactPositiveInteger(value.bytes) ||
    !exactSha256(value.sha256) ||
    !isRecord(value.archive) ||
    !hasOnlyKeys(value.archive, ['format', 'maxExtractedBytes']) ||
    (value.archive.format !== 'tar.gz' && value.archive.format !== 'zip') ||
    !exactPositiveInteger(value.archive.maxExtractedBytes) ||
    !isRecord(value.trust) ||
    !isRecord(value.layout) ||
    !hasOnlyKeys(value.layout, ['entrypoint', 'directories', 'files']) ||
    !safeRelativePath(value.layout.entrypoint) ||
    !Array.isArray(value.layout.directories) ||
    !Array.isArray(value.layout.files)
  ) {
    fail('enhanced_mode_manifest_invalid')
  }
  const expectedTarget =
    value.platform === 'darwin'
      ? value.arch === 'arm64'
        ? 'aarch64-apple-darwin'
        : 'x86_64-apple-darwin'
      : 'x86_64-pc-windows-msvc'
  if (value.target !== expectedTarget || (value.platform === 'win32' && value.arch !== 'x64')) {
    fail('enhanced_mode_manifest_invalid')
  }
  const officialBasename =
    value.target === 'aarch64-apple-darwin'
      ? 'codex-app-server-package-aarch64-apple-darwin.tar.gz'
      : value.target === 'x86_64-apple-darwin'
        ? 'codex-app-server-package-x86_64-apple-darwin.tar.gz'
        : 'codex-app-server-package-x86_64-pc-windows-msvc.tar.gz'
  const trust = value.trust as UnknownRecord
  if (value.platform === 'darwin') {
    if (
      !hasOnlyKeys(trust, ['policy', 'teamIdentifier', 'requireNotarization']) ||
      trust.policy !== 'macos' ||
      !/^[A-Z0-9]{10}$/.test(trust.teamIdentifier as string) ||
      trust.requireNotarization !== true
    )
      fail('enhanced_mode_manifest_invalid')
  } else if (
    !hasOnlyKeys(trust, ['policy', 'publisher']) ||
    trust.policy !== 'windows' ||
    typeof trust.publisher !== 'string' ||
    trust.publisher.length < 3 ||
    trust.publisher.length > 128
  ) {
    fail('enhanced_mode_manifest_invalid')
  }
  let primaryUrl: URL
  let fallbackUrl: URL
  try {
    primaryUrl = new URL(value.primaryUrl as string)
    fallbackUrl = new URL(value.fallbackUrl as string)
  } catch {
    fail('enhanced_mode_manifest_invalid')
  }
  if (
    primaryUrl.protocol !== 'https:' ||
    primaryUrl.hostname !== PRIMARY_DOWNLOAD_HOST ||
    primaryUrl.username !== '' ||
    primaryUrl.password !== '' ||
    primaryUrl.search !== '' ||
    primaryUrl.hash !== '' ||
    !primaryUrl.href.startsWith(MANIFEST_MIRROR_PREFIX) ||
    fallbackUrl.protocol !== 'https:' ||
    fallbackUrl.hostname !== FALLBACK_DOWNLOAD_HOST ||
    fallbackUrl.username !== '' ||
    fallbackUrl.password !== '' ||
    fallbackUrl.search !== '' ||
    fallbackUrl.hash !== '' ||
    !fallbackUrl.href.startsWith(MANIFEST_SOURCE_PREFIX) ||
    fallbackUrl.pathname.toLowerCase().split('/').includes('latest') ||
    primaryUrl.pathname.split('/').at(-1) !== officialBasename ||
    fallbackUrl.pathname.split('/').at(-1) !== officialBasename
  ) {
    fail('enhanced_mode_manifest_invalid')
  }
  const layout = value.layout as UnknownRecord
  const directories = layout.directories as unknown[]
  if (!directories.every(safeRelativePath) || new Set(directories).size !== directories.length) {
    fail('enhanced_mode_manifest_invalid')
  }
  const files = (layout.files as unknown[]).map(parseManifestFile)
  if (files.length === 0 || new Set(files.map((file) => file.path)).size !== files.length) {
    fail('enhanced_mode_manifest_invalid')
  }
  const maxExtractedBytes = (value.archive as UnknownRecord).maxExtractedBytes
  if (
    typeof maxExtractedBytes !== 'number' ||
    files.reduce((sum, file) => sum + file.bytes, 0) > maxExtractedBytes
  ) {
    fail('enhanced_mode_manifest_invalid')
  }
  const canonicalEntries = [...directories, ...files.map((file) => file.path)].map((path) =>
    canonicalTargetPath(path, value.platform as string),
  )
  if (new Set(canonicalEntries).size !== canonicalEntries.length) {
    fail('enhanced_mode_manifest_invalid')
  }
  if (
    !files.some(
      (file) =>
        file.path === layout.entrypoint && file.mode === 'executable' && file.install === true,
    )
  ) {
    fail('enhanced_mode_manifest_invalid')
  }
  const allowed = new Set([...directories, ...files.map((file) => file.path)])
  if (
    [...directories, ...files.map((file) => file.path)].some((path) => {
      const parent = path.split('/').slice(0, -1).join('/')
      return parent !== '' && !allowed.has(parent)
    })
  ) {
    fail('enhanced_mode_manifest_invalid')
  }
  return Object.freeze({
    ...(value as unknown as CodexComponentAssetManifest),
    primaryUrl: primaryUrl.href,
    fallbackUrl: fallbackUrl.href,
    layout: Object.freeze({
      entrypoint: layout.entrypoint as string,
      directories: Object.freeze([...directories] as string[]),
      files: Object.freeze(files.map((file) => Object.freeze({ ...file }))),
    }),
  })
}

export function parseCodexComponentManifest(value: unknown): CodexComponentManifest {
  if (
    !isRecord(value) ||
    !hasOnlyKeys(value, ['schemaVersion', 'component']) ||
    value.schemaVersion !== 1 ||
    !isRecord(value.component) ||
    !hasOnlyKeys(value.component, ['version', 'license', 'assets']) ||
    value.component.version !== PINNED_VERSION ||
    !isRecord(value.component.license) ||
    !hasOnlyKeys(value.component.license, ['spdx', 'sourceUrl']) ||
    value.component.license.spdx !== 'Apache-2.0' ||
    value.component.license.sourceUrl !== LICENSE_SOURCE ||
    !Array.isArray(value.component.assets)
  ) {
    fail('enhanced_mode_manifest_invalid')
  }
  const assets = value.component.assets.map(parseManifestAsset)
  if (
    assets.length !== 3 ||
    new Set(assets.map((asset) => `${asset.platform}-${asset.arch}`)).size !== 3
  )
    fail('enhanced_mode_manifest_invalid')
  for (const required of ['darwin-arm64', 'darwin-x64', 'win32-x64']) {
    if (!assets.some((asset) => `${asset.platform}-${asset.arch}` === required))
      fail('enhanced_mode_manifest_invalid')
  }
  return Object.freeze({
    schemaVersion: 1,
    component: Object.freeze({
      version: PINNED_VERSION,
      license: Object.freeze({ spdx: 'Apache-2.0', sourceUrl: LICENSE_SOURCE }),
      assets: Object.freeze(assets),
    }),
  })
}

function normalizeArchivePath(path: string): string {
  return path.endsWith('/') ? path.slice(0, -1) : path
}

export function validateCodexArchiveEntry(
  path: string,
  type: string,
  size: number,
  seen: Set<string>,
  platform = 'darwin',
): string {
  const normalized = normalizeArchivePath(path)
  let canonical: string
  try {
    canonical = canonicalTargetPath(normalized, platform, 'enhanced_mode_archive_unsafe')
  } catch {
    fail('enhanced_mode_archive_unsafe')
  }
  if (
    (type !== 'File' && type !== 'Directory') ||
    !Number.isSafeInteger(size) ||
    size < 0 ||
    seen.has(canonical)
  ) {
    fail('enhanced_mode_archive_unsafe')
  }
  seen.add(canonical)
  return normalized
}

function safeOwnedChild(root: string, ...parts: string[]): string {
  const candidate = resolve(root, ...parts)
  const rel = relative(root, candidate)
  if (rel === '' || rel === '..' || rel.startsWith(`..${sep}`) || isAbsolute(rel)) {
    fail('enhanced_mode_cache_unsafe')
  }
  return candidate
}

async function sha256File(path: string): Promise<string> {
  const hash = createHash('sha256')
  for await (const chunk of createReadStream(path)) hash.update(chunk)
  return hash.digest('hex')
}

async function defaultCheckAvailableCapacity(
  cacheRoot: string,
  requiredBytes: number,
): Promise<void> {
  try {
    const stats = await statfs(cacheRoot, { bigint: true })
    if (stats.bavail * stats.bsize < BigInt(requiredBytes)) fail('enhanced_mode_capacity_failed')
  } catch (error) {
    if (error instanceof EnhancedModeComponentError) throw error
    fail('enhanced_mode_capacity_failed')
  }
}

async function defaultProbeVersion(
  executablePath: string,
  options: ComponentProbeOptions,
): Promise<string> {
  try {
    const result = await execFile(executablePath, ['--version'], {
      timeout: options.timeoutMs ?? DEFAULT_PROBE_TIMEOUT_MS,
      maxBuffer: MAX_VERSION_OUTPUT,
      windowsHide: true,
      signal: options.signal,
      env: {
        NO_PROXY: '127.0.0.1,localhost',
        no_proxy: '127.0.0.1,localhost',
      },
    })
    return result.stdout.trim()
  } catch {
    fail(options.signal?.aborted ? 'enhanced_mode_cancelled' : 'enhanced_mode_version_probe_failed')
  }
}

async function defaultVerifyPlatformTrust(
  executablePaths: readonly string[],
  policy: CodexComponentAssetManifest['trust'],
  options: ComponentProbeOptions,
): Promise<void> {
  for (const executablePath of executablePaths) {
    if (policy.policy === 'macos') {
      if (process.platform !== 'darwin') fail('enhanced_mode_platform_trust_failed')
      try {
        await execFile(
          '/usr/bin/codesign',
          ['--verify', '--strict', '--verbose=4', executablePath],
          {
            timeout: options.timeoutMs ?? DEFAULT_PROBE_TIMEOUT_MS,
            maxBuffer: MAX_VERSION_OUTPUT,
            windowsHide: true,
            signal: options.signal,
            env: {},
          },
        )
      } catch {
        fail(
          options.signal?.aborted
            ? 'enhanced_mode_cancelled'
            : 'enhanced_mode_macos_codesign_failed',
        )
      }
      try {
        const details = await execFile(
          '/usr/bin/codesign',
          ['-dv', '--verbose=4', executablePath],
          {
            timeout: options.timeoutMs ?? DEFAULT_PROBE_TIMEOUT_MS,
            maxBuffer: MAX_VERSION_OUTPUT,
            signal: options.signal,
            env: {},
          },
        )
        if (
          !`${details.stdout}\n${details.stderr}`.includes(
            `TeamIdentifier=${policy.teamIdentifier}`,
          )
        )
          fail('enhanced_mode_macos_team_identifier_mismatch')
      } catch (error) {
        if (error instanceof EnhancedModeComponentError) throw error
        fail(
          options.signal?.aborted
            ? 'enhanced_mode_cancelled'
            : 'enhanced_mode_macos_codesign_details_failed',
        )
      }
      if (policy.requireNotarization) {
        try {
          await execFile(
            '/usr/sbin/spctl',
            ['--assess', '--type', 'execute', '--verbose=4', executablePath],
            {
              timeout: options.timeoutMs ?? DEFAULT_PROBE_TIMEOUT_MS,
              maxBuffer: MAX_VERSION_OUTPUT,
              windowsHide: true,
              signal: options.signal,
              env: {},
            },
          )
        } catch {
          fail(
            options.signal?.aborted
              ? 'enhanced_mode_cancelled'
              : 'enhanced_mode_macos_notarization_failed',
          )
        }
      }
    } else {
      try {
        if (process.platform !== 'win32') fail('enhanced_mode_platform_trust_failed')
        const escapedPath = executablePath.replaceAll("'", "''")
        const escapedPublisher = policy.publisher.replaceAll("'", "''")
        const script = `$s=Get-AuthenticodeSignature -LiteralPath '${escapedPath}'; if($s.Status -ne 'Valid' -or $s.SignerCertificate.Subject -ne '${escapedPublisher}') { exit 23 }`
        await execFile(
          'powershell.exe',
          ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', script],
          {
            timeout: options.timeoutMs ?? DEFAULT_PROBE_TIMEOUT_MS,
            maxBuffer: MAX_VERSION_OUTPUT,
            windowsHide: true,
            signal: options.signal,
            env: {},
          },
        )
      } catch {
        fail(
          options.signal?.aborted
            ? 'enhanced_mode_cancelled'
            : 'enhanced_mode_platform_trust_failed',
        )
      }
    }
  }
}

export function platformTrustCommands(
  platform: string,
  executablePath: string,
  policy: CodexComponentAssetManifest['trust'],
): readonly { readonly file: string; readonly args: readonly string[] }[] {
  if (policy.policy === 'macos') {
    if (platform !== 'darwin') fail('enhanced_mode_platform_trust_failed')
    return [
      { file: '/usr/bin/codesign', args: ['--verify', '--strict', '--verbose=4', executablePath] },
      { file: '/usr/bin/codesign', args: ['-dv', '--verbose=4', executablePath] },
      {
        file: '/usr/sbin/spctl',
        args: ['--assess', '--type', 'execute', '--verbose=4', executablePath],
      },
    ]
  }
  if (platform !== 'win32') fail('enhanced_mode_platform_trust_failed')
  const escapedPath = executablePath.replaceAll("'", "''")
  const escapedPublisher = policy.publisher.replaceAll("'", "''")
  return [
    {
      file: 'powershell.exe',
      args: [
        '-NoLogo',
        '-NoProfile',
        '-NonInteractive',
        '-Command',
        `$s=Get-AuthenticodeSignature -LiteralPath '${escapedPath}'; if($s.Status -ne 'Valid' -or $s.SignerCertificate.Subject -ne '${escapedPublisher}') { exit 23 }`,
      ],
    },
  ]
}

async function ensurePrivateCacheRoot(cacheRoot: string): Promise<void> {
  if (!isAbsolute(cacheRoot) || cacheRoot === resolve(sep)) fail('enhanced_mode_cache_unsafe')
  await mkdir(cacheRoot, { recursive: true, mode: 0o700 })
  const info = await lstat(cacheRoot)
  if (!info.isDirectory() || info.isSymbolicLink()) fail('enhanced_mode_cache_unsafe')
  if ((await realpath(cacheRoot)) !== resolve(cacheRoot)) fail('enhanced_mode_cache_unsafe')
  if (process.platform !== 'win32') {
    if (typeof process.getuid === 'function' && info.uid !== process.getuid()) {
      fail('enhanced_mode_cache_unsafe')
    }
    await chmod(cacheRoot, 0o700)
  }
}

async function ensureOwnedDirectory(path: string): Promise<void> {
  try {
    await mkdir(path, { recursive: false, mode: 0o700 })
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error
  }
  const info = await lstat(path)
  if (!info.isDirectory() || info.isSymbolicLink() || (await realpath(path)) !== path) {
    fail('enhanced_mode_cache_unsafe')
  }
  if (process.platform !== 'win32') {
    if (typeof process.getuid === 'function' && info.uid !== process.getuid()) {
      fail('enhanced_mode_cache_unsafe')
    }
    await chmod(path, 0o700)
  }
}

function supportedAsset(
  manifest: CodexComponentManifest,
  platform: string,
  arch: string,
): CodexComponentAssetManifest | undefined {
  return manifest.component.assets.find(
    (asset) => asset.platform === platform && asset.arch === arch,
  )
}

async function fetchPinned(
  url: string,
  fetchImplementation: typeof fetch,
  signal: AbortSignal,
  redirects = 0,
  redirectHosts?: ReadonlySet<string>,
): Promise<Response> {
  if (redirects > 5) fail('enhanced_mode_download_failed')
  const initialHost = new URL(url).hostname
  const allowedRedirectHosts =
    redirectHosts ??
    new Set(
      initialHost === PRIMARY_DOWNLOAD_HOST
        ? [PRIMARY_DOWNLOAD_HOST]
        : ['release-assets.githubusercontent.com'],
    )
  let response: Response
  try {
    response = await fetchImplementation(url, { redirect: 'manual', signal })
  } catch {
    fail(signal.aborted ? 'enhanced_mode_cancelled' : 'enhanced_mode_download_failed')
  }
  if (response.status >= 300 && response.status < 400) {
    const location = response.headers.get('location')
    if (!location) fail('enhanced_mode_download_failed')
    const next = new URL(location, url)
    if (
      next.protocol !== 'https:' ||
      !allowedRedirectHosts.has(next.hostname) ||
      next.username !== '' ||
      next.password !== ''
    ) {
      fail('enhanced_mode_download_failed')
    }
    return fetchPinned(next.href, fetchImplementation, signal, redirects + 1, allowedRedirectHosts)
  }
  if (!response.ok || !response.body) fail('enhanced_mode_download_failed')
  return response
}

function combinedAbortSignal(external: AbortSignal | undefined, timeoutMs: number) {
  const controller = new AbortController()
  const timeout = setTimeout(() => controller.abort(), timeoutMs)
  const onAbort = () => controller.abort()
  external?.addEventListener('abort', onAbort, { once: true })
  if (external?.aborted) controller.abort()
  return {
    signal: controller.signal,
    dispose: () => {
      clearTimeout(timeout)
      external?.removeEventListener('abort', onAbort)
    },
  }
}

async function downloadVerifiedArchive(
  asset: CodexComponentAssetManifest,
  destination: string,
  fetchImplementation: typeof fetch,
  timeoutMs: number,
  externalSignal?: AbortSignal,
  onPhase?: (event: ComponentPhaseEvent) => void,
): Promise<void> {
  const abort = combinedAbortSignal(externalSignal, timeoutMs)
  try {
    if (abort.signal.aborted) fail('enhanced_mode_cancelled')
    let lastError: unknown
    for (const url of [asset.primaryUrl, asset.fallbackUrl]) {
      await rm(destination, { force: true })
      let currentPhase: Extract<ComponentInstallPhase, 'download' | 'digest'> = 'download'
      try {
        onPhase?.({ phase: 'download', outcome: 'started' })
        const response = await fetchPinned(url, fetchImplementation, abort.signal)
        if (abort.signal.aborted) fail('enhanced_mode_cancelled')
        let received = 0
        const limiter = new Transform({
          transform(chunk, _encoding, callback) {
            received += Buffer.byteLength(chunk)
            if (received > asset.bytes) return callback(new Error('size limit'))
            callback(null, chunk)
          },
        })
        await pipeline(
          Readable.from(response.body as unknown as AsyncIterable<Uint8Array>),
          limiter,
          createWriteStream(destination, { flags: 'wx', mode: 0o600 }),
          { signal: abort.signal },
        )
        onPhase?.({ phase: 'download', outcome: 'succeeded' })
        currentPhase = 'digest'
        onPhase?.({ phase: 'digest', outcome: 'started' })
        const info = await lstat(destination)
        if (
          !info.isFile() ||
          info.isSymbolicLink() ||
          info.size !== asset.bytes ||
          (await sha256File(destination)) !== asset.sha256
        )
          fail('enhanced_mode_integrity_failed')
        onPhase?.({ phase: 'digest', outcome: 'succeeded' })
        return
      } catch (error) {
        onPhase?.({
          phase: currentPhase,
          outcome: 'failed',
        })
        if (abort.signal.aborted) fail('enhanced_mode_cancelled')
        lastError = error
      }
    }
    if (
      lastError instanceof EnhancedModeComponentError &&
      lastError.code === 'enhanced_mode_integrity_failed'
    )
      throw lastError
    fail('enhanced_mode_download_failed')
  } finally {
    abort.dispose()
  }
}

async function extractVerifiedArchive(
  archivePath: string,
  stagingPath: string,
  asset: CodexComponentAssetManifest,
): Promise<void> {
  const allowedDirectories = new Set(asset.layout.directories)
  const allowedFiles = new Map(asset.layout.files.map((file) => [file.path, file]))
  const seen = new Set<string>()
  let extractionError: EnhancedModeComponentError | undefined
  await mkdir(stagingPath, { recursive: false, mode: 0o700 })
  if (asset.archive.format === 'zip') {
    await extractVerifiedZip(
      archivePath,
      stagingPath,
      asset,
      allowedDirectories,
      allowedFiles,
      seen,
    )
  } else {
    try {
      await extractTar({
        cwd: stagingPath,
        file: archivePath,
        gzip: true,
        strict: true,
        preservePaths: false,
        filter(path, entry) {
          try {
            const readEntry = entry as ReadEntry
            const normalized = validateCodexArchiveEntry(
              path,
              readEntry.type,
              readEntry.size,
              seen,
              asset.platform,
            )
            if (readEntry.type === 'Directory') {
              if (!allowedDirectories.has(normalized) || readEntry.size !== 0)
                fail('enhanced_mode_archive_unsafe')
            } else {
              const expected = allowedFiles.get(normalized)
              if (!expected || readEntry.size !== expected.bytes)
                fail('enhanced_mode_archive_unsafe')
            }
            return true
          } catch (error) {
            extractionError =
              error instanceof EnhancedModeComponentError
                ? error
                : new EnhancedModeComponentError('enhanced_mode_archive_unsafe')
            return false
          }
        },
        onwarn() {
          extractionError = new EnhancedModeComponentError('enhanced_mode_archive_unsafe')
        },
      })
    } catch (error) {
      if (error instanceof EnhancedModeComponentError) throw error
      fail('enhanced_mode_archive_unsafe')
    }
    if (extractionError) throw extractionError
  }
  if (seen.size !== allowedDirectories.size + allowedFiles.size) {
    fail('enhanced_mode_archive_unsafe')
  }
  for (const directory of allowedDirectories) {
    const info = await lstat(safeOwnedChild(stagingPath, directory))
    if (!info.isDirectory() || info.isSymbolicLink()) fail('enhanced_mode_archive_unsafe')
    if (process.platform !== 'win32') await chmod(safeOwnedChild(stagingPath, directory), 0o700)
  }
  for (const file of allowedFiles.values()) {
    const path = safeOwnedChild(stagingPath, file.path)
    const info = await lstat(path)
    if (
      !info.isFile() ||
      info.isSymbolicLink() ||
      info.size !== file.bytes ||
      (await sha256File(path)) !== file.sha256
    ) {
      fail('enhanced_mode_integrity_failed')
    }
    if (process.platform !== 'win32') await chmod(path, file.mode === 'executable' ? 0o700 : 0o600)
  }
  const installedFiles = [...allowedFiles.values()].filter((file) => file.install)
  const installedDirectories = installedDirectorySet(installedFiles)
  for (const file of allowedFiles.values()) {
    if (!file.install) await rm(safeOwnedChild(stagingPath, file.path), { force: true })
  }
  for (const directory of [...allowedDirectories].sort(
    (left, right) => right.length - left.length,
  )) {
    if (!installedDirectories.has(directory)) await rmdir(safeOwnedChild(stagingPath, directory))
  }
}

async function extractVerifiedZip(
  archivePath: string,
  stagingPath: string,
  asset: CodexComponentAssetManifest,
  allowedDirectories: Set<string>,
  allowedFiles: Map<string, CodexComponentFileManifest>,
  seen: Set<string>,
): Promise<void> {
  const bytes = await readFile(archivePath)
  const endStart = Math.max(0, bytes.length - 65_557)
  let end = -1
  for (let offset = bytes.length - 22; offset >= endStart; offset -= 1) {
    if (bytes.readUInt32LE(offset) === 0x06054b50) {
      end = offset
      break
    }
  }
  if (
    end < 0 ||
    bytes.readUInt16LE(end + 4) !== 0 ||
    bytes.readUInt16LE(end + 6) !== 0 ||
    bytes.readUInt16LE(end + 8) !== bytes.readUInt16LE(end + 10) ||
    bytes.readUInt16LE(end + 20) !== 0
  ) {
    fail('enhanced_mode_archive_unsafe')
  }
  const count = bytes.readUInt16LE(end + 10)
  const centralSize = bytes.readUInt32LE(end + 12)
  const centralOffset = bytes.readUInt32LE(end + 16)
  if (count === 0 || count > 512 || centralOffset + centralSize !== end)
    fail('enhanced_mode_archive_unsafe')
  let offset = centralOffset
  let total = 0
  let expectedLocalOffset = 0
  for (let index = 0; index < count; index += 1) {
    if (offset + 46 > end || bytes.readUInt32LE(offset) !== 0x02014b50)
      fail('enhanced_mode_archive_unsafe')
    const flags = bytes.readUInt16LE(offset + 8)
    const method = bytes.readUInt16LE(offset + 10)
    const compressedSize = bytes.readUInt32LE(offset + 20)
    const size = bytes.readUInt32LE(offset + 24)
    const nameLength = bytes.readUInt16LE(offset + 28)
    const extraLength = bytes.readUInt16LE(offset + 30)
    const commentLength = bytes.readUInt16LE(offset + 32)
    const externalAttributes = bytes.readUInt32LE(offset + 38)
    const localOffset = bytes.readUInt32LE(offset + 42)
    if (
      (flags & ~0x800) !== 0 ||
      (method !== 0 && method !== 8) ||
      extraLength !== 0 ||
      commentLength !== 0
    )
      fail('enhanced_mode_archive_unsafe')
    const nameEnd = offset + 46 + nameLength
    if (nameEnd + extraLength + commentLength > end) fail('enhanced_mode_archive_unsafe')
    const name = bytes.subarray(offset + 46, nameEnd).toString('utf8')
    const isDirectory = name.endsWith('/')
    validateZipEntryMode(bytes.readUInt16LE(offset + 4), externalAttributes, isDirectory)
    const normalized = validateCodexArchiveEntry(
      name,
      isDirectory ? 'Directory' : 'File',
      size,
      seen,
      asset.platform,
    )
    if (
      localOffset !== expectedLocalOffset ||
      localOffset + 30 > centralOffset ||
      bytes.readUInt32LE(localOffset) !== 0x04034b50
    )
      fail('enhanced_mode_archive_unsafe')
    const localNameLength = bytes.readUInt16LE(localOffset + 26)
    const localExtraLength = bytes.readUInt16LE(localOffset + 28)
    const dataStart = localOffset + 30 + localNameLength + localExtraLength
    if (
      bytes.readUInt16LE(localOffset + 6) !== flags ||
      bytes.readUInt16LE(localOffset + 8) !== method ||
      bytes.readUInt32LE(localOffset + 18) !== compressedSize ||
      bytes.readUInt32LE(localOffset + 22) !== size ||
      localExtraLength !== 0 ||
      bytes.subarray(localOffset + 30, localOffset + 30 + localNameLength).toString('utf8') !==
        name ||
      dataStart + compressedSize > centralOffset
    )
      fail('enhanced_mode_archive_unsafe')
    expectedLocalOffset = dataStart + compressedSize
    if (isDirectory) {
      if (!allowedDirectories.has(normalized) || size !== 0 || compressedSize !== 0)
        fail('enhanced_mode_archive_unsafe')
      await mkdir(safeOwnedChild(stagingPath, normalized), { recursive: false, mode: 0o700 })
    } else {
      const expected = allowedFiles.get(normalized)
      if (!expected || expected.bytes !== size) fail('enhanced_mode_archive_unsafe')
      const compressed = bytes.subarray(dataStart, dataStart + compressedSize)
      let contents: Buffer
      try {
        contents =
          method === 0
            ? Buffer.from(compressed)
            : inflateRawSync(compressed, { maxOutputLength: size })
      } catch {
        fail('enhanced_mode_archive_unsafe')
      }
      if (contents.length !== size) fail('enhanced_mode_archive_unsafe')
      total += size
      if (total > asset.archive.maxExtractedBytes) fail('enhanced_mode_archive_unsafe')
      await writeFileSafe(safeOwnedChild(stagingPath, normalized), contents)
    }
    offset = nameEnd + extraLength + commentLength
  }
  if (offset !== end || expectedLocalOffset !== centralOffset) fail('enhanced_mode_archive_unsafe')
}

export function validateZipEntryMode(
  versionMadeBy: number,
  externalAttributes: number,
  isDirectory: boolean,
): void {
  const creator = versionMadeBy >>> 8
  const unixType = (externalAttributes >>> 16) & 0xf000
  if (creator !== 3 || unixType !== (isDirectory ? 0x4000 : 0x8000)) {
    fail('enhanced_mode_archive_unsafe')
  }
}

async function writeFileSafe(path: string, contents: Buffer): Promise<void> {
  const handle = await open(path, 'wx', 0o600).catch(() => fail('enhanced_mode_archive_unsafe'))
  try {
    await handle.writeFile(contents)
    await handle.sync()
  } finally {
    await handle.close()
  }
}

function installedDirectorySet(files: readonly CodexComponentFileManifest[]): Set<string> {
  const result = new Set<string>()
  for (const file of files) {
    const segments = file.path.split('/').slice(0, -1)
    for (let index = 1; index <= segments.length; index += 1) {
      result.add(segments.slice(0, index).join('/'))
    }
  }
  return result
}

interface LockOwner {
  readonly pid: number
  readonly token: string
}

async function readLockOwner(path: string): Promise<LockOwner | null> {
  try {
    const value: unknown = JSON.parse(await readFile(path, 'utf8'))
    if (
      !isRecord(value) ||
      !Number.isSafeInteger(value.pid) ||
      (value.pid as number) <= 0 ||
      typeof value.token !== 'string' ||
      !/^[a-f0-9]{32}$/.test(value.token)
    ) {
      return null
    }
    return value as unknown as LockOwner
  } catch {
    return null
  }
}

function processAlive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === 'EPERM'
  }
}

async function acquireLock(
  path: string,
  timeoutMs: number,
  staleMs: number,
  signal?: AbortSignal,
): Promise<() => Promise<void>> {
  const deadline = Date.now() + timeoutMs
  const token = randomBytes(16).toString('hex')
  while (true) {
    if (signal?.aborted) fail('enhanced_mode_cancelled')
    let handle
    try {
      handle = await open(path, 'wx', 0o600)
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'EEXIST') fail('enhanced_mode_cache_unsafe')
      const owner = await readLockOwner(path)
      const info = await lstat(path).catch(() => null)
      const age = info ? Date.now() - info.mtimeMs : 0
      const stale = owner
        ? !processAlive(owner.pid) || age > staleMs
        : age > Math.min(staleMs, 30_000)
      if (stale) {
        await rm(path, { force: true })
        continue
      }
      if (Date.now() >= deadline) fail('enhanced_mode_busy')
      await new Promise((resolveWait) => setTimeout(resolveWait, 50))
      continue
    }
    try {
      await handle.writeFile(JSON.stringify({ pid: process.pid, token }), 'utf8')
      await handle.sync()
    } finally {
      await handle.close()
    }
    return async () => {
      const current = await readLockOwner(path)
      if (current?.token === token) await rm(path, { force: true })
    }
  }
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

export class EnhancedModeComponentManager {
  readonly #cacheRoot: string
  readonly #manifest: CodexComponentManifest
  readonly #platform: string
  readonly #arch: string
  readonly #fetch: typeof fetch
  readonly #probeVersion: ComponentVersionProbe
  readonly #verifyPlatformTrust: ComponentPlatformTrustVerifier
  readonly #checkAvailableCapacity: ComponentCapacityCheck
  readonly #downloadTimeoutMs: number
  readonly #probeTimeoutMs: number
  readonly #lockTimeoutMs: number
  readonly #lockStaleMs: number
  readonly #onPhase?: (event: ComponentPhaseEvent) => void
  #installPromise: Promise<InstalledEnhancedModeComponent> | undefined
  #removePromise: Promise<void> | undefined

  constructor(options: EnhancedModeComponentManagerOptions) {
    if (!options || typeof options.cacheRoot !== 'string' || !isAbsolute(options.cacheRoot)) {
      throw new TypeError('enhanced_mode_cache_must_be_absolute')
    }
    this.#cacheRoot = resolve(options.cacheRoot)
    this.#manifest = parseCodexComponentManifest(options.manifest)
    this.#platform = options.platform ?? process.platform
    this.#arch = options.arch ?? process.arch
    this.#fetch = options.fetchImplementation ?? globalThis.fetch
    this.#probeVersion = options.probeVersion ?? defaultProbeVersion
    this.#verifyPlatformTrust = options.verifyPlatformTrust ?? defaultVerifyPlatformTrust
    this.#checkAvailableCapacity = options.checkAvailableCapacity ?? defaultCheckAvailableCapacity
    this.#downloadTimeoutMs = this.#duration(options.downloadTimeoutMs, DEFAULT_DOWNLOAD_TIMEOUT_MS)
    this.#probeTimeoutMs = this.#duration(options.probeTimeoutMs, DEFAULT_PROBE_TIMEOUT_MS)
    this.#lockTimeoutMs = this.#duration(options.lockTimeoutMs, DEFAULT_LOCK_TIMEOUT_MS)
    this.#lockStaleMs = this.#duration(
      options.lockStaleMs,
      this.#downloadTimeoutMs + Math.max(5 * 60_000, this.#probeTimeoutMs * 2),
    )
    this.#onPhase = options.onPhase
  }

  async status(): Promise<EnhancedModeComponentStatus> {
    const asset = this.#asset()
    if (!asset) return { state: 'unsupported', supported: false, version: PINNED_VERSION }
    try {
      await ensurePrivateCacheRoot(this.#cacheRoot)
    } catch {
      return { state: 'invalid', supported: true, version: PINNED_VERSION }
    }
    const installPath = this.#installPath(asset)
    if (!(await pathExists(installPath))) {
      return { state: 'missing', supported: true, version: PINNED_VERSION }
    }
    try {
      await this.#verifyInstalled(asset)
      return { state: 'ready', supported: true, version: PINNED_VERSION }
    } catch {
      return { state: 'invalid', supported: true, version: PINNED_VERSION }
    }
  }

  install(
    options: { readonly signal?: AbortSignal } = {},
  ): Promise<InstalledEnhancedModeComponent> {
    if (this.#installPromise) return this.#installPromise
    const promise = this.#install(options.signal)
    this.#installPromise = promise
    void promise.then(
      () => {
        if (this.#installPromise === promise) this.#installPromise = undefined
      },
      () => {
        if (this.#installPromise === promise) this.#installPromise = undefined
      },
    )
    return promise
  }

  async resolveExecutable(options: { readonly signal?: AbortSignal } = {}): Promise<string> {
    const asset = this.#asset()
    if (!asset) fail('enhanced_mode_unsupported')
    await ensurePrivateCacheRoot(this.#cacheRoot).catch(() =>
      fail('enhanced_mode_integrity_failed'),
    )
    if (!(await pathExists(this.#installPath(asset)))) fail('enhanced_mode_install_required')
    try {
      return (await this.#verifyInstalled(asset, options.signal)).executablePath
    } catch (error) {
      if (error instanceof EnhancedModeComponentError) throw error
      fail('enhanced_mode_integrity_failed')
    }
  }

  remove(options: { readonly signal?: AbortSignal } = {}): Promise<void> {
    if (this.#removePromise) return this.#removePromise
    const promise = this.#remove(options.signal)
    this.#removePromise = promise
    void promise.then(
      () => {
        if (this.#removePromise === promise) this.#removePromise = undefined
      },
      () => {
        if (this.#removePromise === promise) this.#removePromise = undefined
      },
    )
    return promise
  }

  async #install(signal?: AbortSignal): Promise<InstalledEnhancedModeComponent> {
    const asset = this.#asset()
    if (!asset) fail('enhanced_mode_unsupported')
    if (signal?.aborted) fail('enhanced_mode_cancelled')
    await ensurePrivateCacheRoot(this.#cacheRoot)
    const release = await acquireLock(
      safeOwnedChild(this.#cacheRoot, '.install.lock'),
      this.#lockTimeoutMs,
      this.#lockStaleMs,
      signal,
    )
    const token = randomBytes(16).toString('hex')
    const archivePath = safeOwnedChild(this.#cacheRoot, `.download.${token}.part`)
    const stagingPath = safeOwnedChild(this.#cacheRoot, `.staging.${token}`)
    let published = false
    let verified = false
    const updating = await this.#hasOlderVersion()
    if (updating) this.#phase({ phase: 'update', outcome: 'started' })
    try {
      if (await pathExists(this.#installPath(asset))) {
        try {
          return await this.#verifyInstalled(asset, signal)
        } catch (error) {
          if (
            signal?.aborted ||
            (error instanceof EnhancedModeComponentError &&
              error.code === 'enhanced_mode_cancelled')
          ) {
            fail('enhanced_mode_cancelled')
          }
          await this.#removeOwnedInstall(asset)
        }
      }
      const requiredCapacity = asset.bytes + asset.archive.maxExtractedBytes + CAPACITY_MARGIN_BYTES
      await this.#checkAvailableCapacity(this.#cacheRoot, requiredCapacity).catch(() =>
        fail('enhanced_mode_capacity_failed'),
      )
      await downloadVerifiedArchive(
        asset,
        archivePath,
        this.#fetch,
        this.#downloadTimeoutMs,
        signal,
        (event) => this.#phase(event),
      )
      await extractVerifiedArchive(archivePath, stagingPath, asset)
      const executablePath = safeOwnedChild(stagingPath, asset.layout.entrypoint)
      const executablePaths = asset.layout.files
        .filter((file) => file.install && file.mode === 'executable')
        .map((file) => safeOwnedChild(stagingPath, file.path))
      this.#phase({ phase: 'signature', outcome: 'started' })
      await this.#verifyPlatformTrust(executablePaths, asset.trust, {
        signal,
        timeoutMs: this.#probeTimeoutMs,
      }).catch((error: unknown) => {
        this.#phase({ phase: 'signature', outcome: 'failed' })
        failPlatformTrust(error, signal)
      })
      this.#phase({ phase: 'signature', outcome: 'succeeded' })
      const output = await this.#probeVersion(executablePath, {
        signal,
        timeoutMs: this.#probeTimeoutMs,
      }).catch(() =>
        fail(signal?.aborted ? 'enhanced_mode_cancelled' : 'enhanced_mode_version_probe_failed'),
      )
      if (output.trim() !== PINNED_VERSION_OUTPUT) fail('enhanced_mode_version_mismatch')
      const installPath = this.#installPath(asset)
      await ensureOwnedDirectory(safeOwnedChild(this.#cacheRoot, PINNED_VERSION))
      if (signal?.aborted) fail('enhanced_mode_cancelled')
      this.#phase({ phase: 'promote', outcome: 'started' })
      await rename(stagingPath, installPath).catch((error) => {
        this.#phase({ phase: 'promote', outcome: 'failed' })
        throw error
      })
      this.#phase({ phase: 'promote', outcome: 'succeeded' })
      published = true
      await this.#garbageCollectOldVersions()
      const installed = await this.#verifyInstalled(asset, signal)
      verified = true
      return installed
    } finally {
      await rm(archivePath, { force: true }).catch(() => undefined)
      await rm(stagingPath, { recursive: true, force: true }).catch(() => undefined)
      let cleanupFailed = false
      if (published && !verified) {
        await this.#removeOwnedInstall(asset).catch(() => {
          cleanupFailed = true
        })
      }
      await release().catch(() => {
        cleanupFailed = true
      })
      if (updating)
        this.#phase({
          phase: 'update',
          outcome: verified && !cleanupFailed ? 'succeeded' : 'failed',
        })
      if (cleanupFailed) fail('enhanced_mode_cache_unsafe')
    }
  }

  async #remove(signal?: AbortSignal): Promise<void> {
    const asset = this.#asset()
    if (!asset) fail('enhanced_mode_unsupported')
    if (signal?.aborted) fail('enhanced_mode_cancelled')
    await ensurePrivateCacheRoot(this.#cacheRoot)
    const release = await acquireLock(
      safeOwnedChild(this.#cacheRoot, '.install.lock'),
      this.#lockTimeoutMs,
      this.#lockStaleMs,
      signal,
    )
    try {
      this.#phase({ phase: 'remove', outcome: 'started' })
      await this.#removeOwnedInstall(asset).then(
        () => this.#phase({ phase: 'remove', outcome: 'succeeded' }),
        (error) => {
          this.#phase({ phase: 'remove', outcome: 'failed' })
          throw error
        },
      )
    } finally {
      await release()
    }
  }

  #phase(event: ComponentPhaseEvent): void {
    try {
      this.#onPhase?.(Object.freeze(event))
    } catch {}
  }

  async #verifyInstalled(
    asset: CodexComponentAssetManifest,
    signal?: AbortSignal,
  ): Promise<InstalledEnhancedModeComponent> {
    if (signal?.aborted) fail('enhanced_mode_cancelled')
    const installPath = this.#installPath(asset)
    const installInfo = await lstat(installPath).catch(() => fail('enhanced_mode_install_required'))
    if (
      !installInfo.isDirectory() ||
      installInfo.isSymbolicLink() ||
      (process.platform !== 'win32' && (installInfo.mode & 0o777) !== 0o700) ||
      (typeof process.getuid === 'function' && installInfo.uid !== process.getuid())
    ) {
      fail('enhanced_mode_integrity_failed')
    }
    if ((await realpath(installPath)) !== installPath) fail('enhanced_mode_integrity_failed')
    const installedFiles = asset.layout.files.filter((file) => file.install)
    const installedDirectories = installedDirectorySet(installedFiles)
    const expectedEntries = new Set([
      ...installedDirectories,
      ...installedFiles.map((file) => file.path),
    ])
    const actualEntries = new Set<string>()
    const pending = [{ absolute: installPath, relative: '' }]
    while (pending.length > 0) {
      const current = pending.pop()!
      for (const entry of await readdir(current.absolute, { withFileTypes: true })) {
        const relativePath =
          current.relative === '' ? entry.name : `${current.relative}/${entry.name}`
        if (!safeRelativePath(relativePath) || entry.isSymbolicLink()) {
          fail('enhanced_mode_integrity_failed')
        }
        actualEntries.add(relativePath)
        if (entry.isDirectory()) {
          pending.push({
            absolute: safeOwnedChild(installPath, relativePath),
            relative: relativePath,
          })
        } else if (!entry.isFile()) {
          fail('enhanced_mode_integrity_failed')
        }
      }
    }
    if (
      actualEntries.size !== expectedEntries.size ||
      [...actualEntries].some((path) => !expectedEntries.has(path))
    ) {
      fail('enhanced_mode_integrity_failed')
    }
    for (const directory of installedDirectories) {
      const path = safeOwnedChild(installPath, directory)
      const info = await lstat(path).catch(() => fail('enhanced_mode_integrity_failed'))
      if (
        !info.isDirectory() ||
        info.isSymbolicLink() ||
        (await realpath(path)) !== path ||
        (process.platform !== 'win32' && (info.mode & 0o777) !== 0o700) ||
        (typeof process.getuid === 'function' && info.uid !== process.getuid())
      ) {
        fail('enhanced_mode_integrity_failed')
      }
    }
    for (const file of installedFiles) {
      const path = safeOwnedChild(installPath, file.path)
      const info = await lstat(path).catch(() => fail('enhanced_mode_integrity_failed'))
      if (
        !info.isFile() ||
        info.isSymbolicLink() ||
        info.size !== file.bytes ||
        (process.platform !== 'win32' &&
          (info.mode & 0o777) !== (file.mode === 'executable' ? 0o700 : 0o600)) ||
        (typeof process.getuid === 'function' && info.uid !== process.getuid()) ||
        (await sha256File(path)) !== file.sha256
      ) {
        fail('enhanced_mode_integrity_failed')
      }
    }
    const executablePath = safeOwnedChild(installPath, asset.layout.entrypoint)
    await this.#verifyPlatformTrust(
      installedFiles
        .filter((file) => file.mode === 'executable')
        .map((file) => safeOwnedChild(installPath, file.path)),
      asset.trust,
      { signal, timeoutMs: this.#probeTimeoutMs },
    ).catch((error: unknown) => failPlatformTrust(error, signal))
    const output = await this.#probeVersion(executablePath, {
      signal,
      timeoutMs: this.#probeTimeoutMs,
    }).catch(() =>
      fail(signal?.aborted ? 'enhanced_mode_cancelled' : 'enhanced_mode_version_probe_failed'),
    )
    if (output.trim() !== PINNED_VERSION_OUTPUT) fail('enhanced_mode_version_mismatch')
    return {
      executablePath,
      version: PINNED_VERSION,
      platform: asset.platform,
      arch: asset.arch,
    }
  }

  async #removeOwnedInstall(asset: CodexComponentAssetManifest): Promise<void> {
    const installPath = this.#installPath(asset)
    if (!(await pathExists(installPath))) return
    const info = await lstat(installPath)
    if (!info.isDirectory() || info.isSymbolicLink()) fail('enhanced_mode_cache_unsafe')
    await rm(installPath, { recursive: true, force: true })
    const versionPath = safeOwnedChild(this.#cacheRoot, PINNED_VERSION)
    const entries = await readdir(versionPath).catch(() => ['not-empty'])
    if (entries.length === 0) await rmdir(versionPath)
  }

  async #garbageCollectOldVersions(): Promise<void> {
    const entries = await readdir(this.#cacheRoot, { withFileTypes: true })
    for (const entry of entries) {
      if (
        entry.name === PINNED_VERSION ||
        !/^\d+\.\d+\.\d+$/.test(entry.name) ||
        !entry.isDirectory() ||
        entry.isSymbolicLink()
      ) {
        continue
      }
      const candidate = safeOwnedChild(this.#cacheRoot, entry.name)
      const info = await lstat(candidate)
      if (info.isDirectory() && !info.isSymbolicLink()) {
        await rm(candidate, { recursive: true, force: true })
      }
    }
  }

  async #hasOlderVersion(): Promise<boolean> {
    const target = PINNED_VERSION.split('.').map(Number)
    const entries = await readdir(this.#cacheRoot, { withFileTypes: true })
    return entries.some((entry) => {
      if (!entry.isDirectory() || entry.isSymbolicLink() || !/^\d+\.\d+\.\d+$/.test(entry.name))
        return false
      const candidate = entry.name.split('.').map(Number)
      for (let index = 0; index < target.length; index += 1) {
        if (candidate[index]! < target[index]!) return true
        if (candidate[index]! > target[index]!) return false
      }
      return false
    })
  }

  #asset(): CodexComponentAssetManifest | undefined {
    return supportedAsset(this.#manifest, this.#platform, this.#arch)
  }

  #installPath(asset: CodexComponentAssetManifest): string {
    return safeOwnedChild(this.#cacheRoot, PINNED_VERSION, `${asset.platform}-${asset.arch}`)
  }

  #duration(value: number | undefined, fallback: number): number {
    const result = value ?? fallback
    if (!Number.isSafeInteger(result) || result <= 0) throw new TypeError('invalid_timeout')
    return result
  }
}
