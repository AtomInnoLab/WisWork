import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { basename, join, resolve } from 'node:path'

const MAX_INVENTORY_ENTRIES = 20_000
const MAX_INVENTORY_DEPTH = 16
const prohibitedArtifact =
  /(?:codex|app-server).*(?:\.tar\.gz|\.zip|\.exe|\.app|\.dll|\.dylib|\.node|\.so|\.wasm)$|^(?:codex|codex-app-server|app-server)$/i
const prohibitedConfigReference =
  /(?:from|files?)\s*:\s*(?:\[[^\]]*)?['"][^'"]*(?:codex|app-server)|(?:codex|app-server)[^'"\s]*(?:\.tar\.gz|\.zip|\.exe|\.app|\.dll|\.dylib|\.node|\.so|\.wasm)(?=['"\s,}])/i

function assertSafeName(path) {
  if (prohibitedArtifact.test(basename(path))) {
    throw new Error(`optional Codex artifact must not be bundled: ${path}`)
  }
}

function inventory(directory, state, depth = 0) {
  if (!existsSync(directory)) return
  if (depth > MAX_INVENTORY_DEPTH) throw new Error('package inventory exceeds depth limit')
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    state.entries += 1
    if (state.entries > MAX_INVENTORY_ENTRIES)
      throw new Error('package inventory exceeds entry limit')
    const path = join(directory, entry.name)
    assertSafeName(path)
    if (entry.isDirectory()) inventory(path, state, depth + 1)
  }
}

function configuredEntries(config) {
  return [
    ...(Array.isArray(config.files) ? config.files : []),
    ...(Array.isArray(config.extraResources) ? config.extraResources : []),
    ...(Array.isArray(config.mac?.extraResources) ? config.mac.extraResources : []),
    ...(Array.isArray(config.win?.extraResources) ? config.win.extraResources : []),
    ...(Array.isArray(config.linux?.extraResources) ? config.linux.extraResources : []),
  ]
}

function sourceOf(entry) {
  if (typeof entry === 'string') return entry
  if (entry && typeof entry === 'object' && Object.getPrototypeOf(entry) === Object.prototype) {
    return typeof entry.from === 'string' ? entry.from : null
  }
  throw new Error('packaging entry must be a string or data-only from record')
}

export function assertOptionalRuntimePackagingPolicy({ root, packagingConfig } = {}) {
  if (typeof root !== 'string') throw new TypeError('root is required')
  const shellRoot = join(root, 'apps/shell')
  let config = packagingConfig
  if (config === undefined) {
    const configSource = readFileSync(join(shellRoot, 'electron-builder.cjs'), 'utf8')
    if (prohibitedConfigReference.test(configSource)) {
      throw new Error('electron-builder configuration references a bundled Codex artifact')
    }
    config = { files: ['out/**'] }
  }
  const state = { entries: 0 }

  for (const entry of configuredEntries(config)) {
    const source = sourceOf(entry)
    if (source === null) continue
    assertSafeName(source)
    const staticPrefix = source.split(/[*!?[\]{}]/, 1)[0].replace(/[\\/]$/, '')
    if (staticPrefix) inventory(resolve(shellRoot, staticPrefix), state)
  }
  inventory(join(shellRoot, 'release'), state)
}
