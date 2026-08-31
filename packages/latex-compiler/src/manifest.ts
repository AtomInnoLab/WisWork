import { manifestError } from './errors.js'

export const TECTONIC_ASSET_HOST_ALLOWLIST = Object.freeze([
  'github.com',
  'release-assets.githubusercontent.com',
  'relay.fullyjustified.net',
  'data1.fullyjustified.net',
] as const)
export const TECTONIC_LICENSE_HOST_ALLOWLIST = Object.freeze(['github.com', 'tug.org'] as const)
export const TECTONIC_REMOTE_INDEXED_BUNDLE_URL =
  'https://relay.fullyjustified.net/default_bundle_v33.tar'

export function isRemoteIndexedBundleUrl(value: string): boolean {
  return value === TECTONIC_REMOTE_INDEXED_BUNDLE_URL
}

export interface AssetLicense {
  readonly spdx: string
  readonly sourceUrl: string
}

export interface TectonicPlatformAsset {
  readonly id: string
  readonly platform: string
  readonly url: string
  readonly bytes: number
  readonly sha256: string
  readonly archive: {
    readonly format: 'tar.gz' | 'zip'
    readonly executable: 'tectonic' | 'tectonic.exe'
  }
}

export interface TectonicBundleAsset {
  readonly id: string
  readonly url: string
  readonly bytes: number
  readonly sha256: string
  readonly license: AssetLicense
}

export interface TectonicManifest {
  readonly schemaVersion: 1
  readonly tectonic: {
    readonly version: string
    readonly license: AssetLicense
    readonly assets: readonly TectonicPlatformAsset[]
  }
  readonly bundle: TectonicBundleAsset
}

const record = (value: unknown, label: string): Record<string, unknown> => {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    manifestError(`${label} must be an object`)
  return value as Record<string, unknown>
}

const exactKeys = (value: Record<string, unknown>, allowed: readonly string[], label: string) => {
  const unexpected = Object.keys(value).filter((key) => !allowed.includes(key))
  if (unexpected.length > 0) manifestError(`${label} contains unknown fields`)
}

const string = (value: unknown, label: string): string => {
  if (typeof value !== 'string' || value.length === 0)
    manifestError(`${label} must be a non-empty string`)
  return value
}

const id = (value: unknown, label: string): string => {
  const parsed = string(value, label)
  if (!/^[a-zA-Z0-9][a-zA-Z0-9._-]{0,127}$/.test(parsed)) manifestError(`${label} is invalid`)
  return parsed
}

const digest = (value: unknown, label: string): string => {
  const parsed = string(value, label)
  if (!/^[a-f0-9]{64}$/.test(parsed)) manifestError(`${label} must be a lowercase SHA-256 digest`)
  return parsed
}

const bytes = (value: unknown, label: string): number => {
  if (!Number.isSafeInteger(value) || (value as number) <= 0)
    manifestError(`${label} must be a positive safe integer`)
  return value as number
}

const pinnedUrl = (value: unknown, label: string, allowedHosts: readonly string[]): string => {
  const parsed = string(value, label)
  let url: URL
  try {
    url = new URL(parsed)
  } catch {
    manifestError(`${label} must be a valid URL`)
  }
  if (url.protocol !== 'https:') manifestError(`${label} must use HTTPS`)
  if (!allowedHosts.includes(url.hostname)) {
    manifestError(`${label} host is not approved`)
  }
  if (url.username || url.password || url.search || url.hash)
    manifestError(`${label} must not contain credentials, query, or fragment`)
  if (url.pathname.toLowerCase().split('/').includes('latest'))
    manifestError(`${label} must not use latest`)
  return url.href
}

const license = (value: unknown, label: string): AssetLicense => {
  const item = record(value, label)
  exactKeys(item, ['spdx', 'sourceUrl'], label)
  return Object.freeze({
    spdx: string(item.spdx, `${label}.spdx`),
    sourceUrl: pinnedUrl(item.sourceUrl, `${label}.sourceUrl`, TECTONIC_LICENSE_HOST_ALLOWLIST),
  })
}

export function parseTectonicManifest(value: unknown): TectonicManifest {
  const root = record(value, 'manifest')
  exactKeys(root, ['schemaVersion', 'tectonic', 'bundle'], 'manifest')
  if (root.schemaVersion !== 1) manifestError('manifest.schemaVersion must be 1')
  const tectonic = record(root.tectonic, 'manifest.tectonic')
  exactKeys(tectonic, ['version', 'license', 'assets'], 'manifest.tectonic')
  const version = string(tectonic.version, 'manifest.tectonic.version')
  if (!/^\d+\.\d+\.\d+$/.test(version)) manifestError('manifest.tectonic.version must be exact')
  if (!Array.isArray(tectonic.assets) || tectonic.assets.length === 0) {
    manifestError('manifest.tectonic.assets must be a non-empty array')
  }
  const platforms = new Set<string>()
  const ids = new Set<string>()
  const assets = tectonic.assets.map((candidate, index): TectonicPlatformAsset => {
    const item = record(candidate, `manifest.tectonic.assets[${index}]`)
    exactKeys(item, ['id', 'platform', 'url', 'bytes', 'sha256', 'archive'], 'Tectonic asset')
    const assetId = id(item.id, `manifest.tectonic.assets[${index}].id`)
    const platform = id(item.platform, `manifest.tectonic.assets[${index}].platform`)
    if (ids.has(assetId) || platforms.has(platform))
      manifestError('duplicate Tectonic asset ID or platform')
    ids.add(assetId)
    platforms.add(platform)
    const archive = record(item.archive, `manifest.tectonic.assets[${index}].archive`)
    exactKeys(archive, ['format', 'executable'], 'Tectonic archive')
    const archiveFormat = string(archive.format, 'archive.format')
    const archiveExecutable = id(archive.executable, 'archive.executable')
    let parsedArchive: TectonicPlatformAsset['archive']
    if (archiveFormat === 'tar.gz' && archiveExecutable === 'tectonic') {
      parsedArchive = Object.freeze({ format: 'tar.gz', executable: 'tectonic' })
    } else if (archiveFormat === 'zip' && archiveExecutable === 'tectonic.exe') {
      parsedArchive = Object.freeze({ format: 'zip', executable: 'tectonic.exe' })
    } else {
      manifestError('Tectonic archive layout is invalid')
    }
    return Object.freeze({
      id: assetId,
      platform,
      url: pinnedUrl(item.url, `manifest.tectonic.assets[${index}].url`, ['github.com']),
      bytes: bytes(item.bytes, `manifest.tectonic.assets[${index}].bytes`),
      sha256: digest(item.sha256, `manifest.tectonic.assets[${index}].sha256`),
      archive: parsedArchive,
    })
  })
  const bundleValue = record(root.bundle, 'manifest.bundle')
  exactKeys(bundleValue, ['id', 'url', 'bytes', 'sha256', 'license'], 'manifest.bundle')
  const bundleId = id(bundleValue.id, 'manifest.bundle.id')
  if (ids.has(bundleId)) manifestError('duplicate asset ID')
  const bundle = Object.freeze({
    id: bundleId,
    url: pinnedUrl(bundleValue.url, 'manifest.bundle.url', [
      'relay.fullyjustified.net',
      'data1.fullyjustified.net',
    ]),
    bytes: bytes(bundleValue.bytes, 'manifest.bundle.bytes'),
    sha256: digest(bundleValue.sha256, 'manifest.bundle.sha256'),
    license: license(bundleValue.license, 'manifest.bundle.license'),
  })
  return Object.freeze({
    schemaVersion: 1 as const,
    tectonic: Object.freeze({
      version,
      license: license(tectonic.license, 'manifest.tectonic.license'),
      assets: Object.freeze(assets),
    }),
    bundle,
  })
}
