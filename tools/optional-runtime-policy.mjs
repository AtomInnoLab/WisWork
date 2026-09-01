import { createRequire } from 'node:module'
import { existsSync, lstatSync, opendirSync, readlinkSync, realpathSync } from 'node:fs'
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const require = createRequire(import.meta.url)
const MAX_INVENTORY_ENTRIES = 20_000
const MAX_INVENTORY_DEPTH = 16
const prohibitedArtifact =
  /(?:codex|app-server).*(?:\.tar\.gz|\.zip|\.exe|\.app|\.dll|\.dylib|\.node|\.so|\.wasm)$|^(?:codex|codex-app-server|app-server)$/i

function assertSafeName(path) {
  if (prohibitedArtifact.test(basename(path))) {
    throw new Error(`optional Codex artifact must not be bundled: ${path}`)
  }
}

function contained(root, path) {
  const child = relative(root, path)
  return child === '' || (!child.startsWith('..') && !isAbsolute(child))
}

function inventory(path, state, { allowMissing, allowSymlinks, inventoryRoot }, depth = 0) {
  if (!existsSync(path)) {
    if (allowMissing) return
    throw new Error(`expected package artifact is missing: ${path}`)
  }
  if (depth > MAX_INVENTORY_DEPTH) throw new Error('package inventory exceeds depth limit')
  assertSafeName(path)
  const info = lstatSync(path)
  if (info.isSymbolicLink()) {
    if (!allowSymlinks) throw new Error(`package input must not be a symlink: ${path}`)
    const target = realpathSync(resolve(dirname(path), readlinkSync(path)))
    if (!contained(inventoryRoot, target))
      throw new Error(`package symlink escapes inventory: ${path}`)
    assertSafeName(target)
    return inventory(
      target,
      state,
      { allowMissing: false, allowSymlinks, inventoryRoot },
      depth + 1,
    )
  }
  if (!info.isDirectory()) return
  const real = realpathSync(path)
  if (state.visited.has(real)) return
  state.visited.add(real)
  const directory = opendirSync(path)
  try {
    for (;;) {
      const entry = directory.readSync()
      if (entry === null) break
      state.entries += 1
      if (state.entries > MAX_INVENTORY_ENTRIES) {
        throw new Error('package inventory exceeds entry limit')
      }
      inventory(
        join(path, entry.name),
        state,
        { allowMissing: false, allowSymlinks, inventoryRoot },
        depth + 1,
      )
    }
  } finally {
    directory.closeSync()
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

function packageInputsConfig(inputs) {
  return {
    files: inputs.files,
    extraResources: inputs.extraResources,
    mac: { extraResources: inputs.macExtraResources },
    win: { extraResources: inputs.winExtraResources },
    linux: { extraResources: inputs.linuxExtraResources },
  }
}

export function assertOptionalRuntimePackagingPolicy({
  root,
  mode = 'source',
  packagingConfig,
  packageInputs,
  artifactDirectories,
} = {}) {
  if (typeof root !== 'string') throw new TypeError('root is required')
  if (mode !== 'source' && mode !== 'post-package') throw new TypeError('invalid policy mode')
  const state = { entries: 0, visited: new Set() }
  if (mode === 'post-package') {
    if (!Array.isArray(artifactDirectories) || artifactDirectories.length === 0) {
      throw new Error('post-package mode requires artifact directories')
    }
    for (const directory of artifactDirectories) {
      const absolute = resolve(directory)
      inventory(absolute, state, {
        allowMissing: false,
        allowSymlinks: true,
        inventoryRoot: absolute,
      })
    }
    return
  }

  const shellRoot = join(root, 'apps/shell')
  const config =
    packagingConfig ??
    (packageInputs
      ? packageInputsConfig(packageInputs)
      : require(join(shellRoot, 'electron-builder.cjs')))
  for (const entry of configuredEntries(config)) {
    const source = sourceOf(entry)
    if (source === null) continue
    assertSafeName(source)
    const staticPrefix = source.split(/[*!?[\]{}]/, 1)[0].replace(/[\\/]$/, '')
    if (!staticPrefix) continue
    const absolute = resolve(shellRoot, staticPrefix)
    if (
      absolute.split(/[\\/]/).includes('node_modules') &&
      existsSync(absolute) &&
      lstatSync(absolute).isDirectory()
    ) {
      throw new Error('broad node_modules package inputs are not auditable')
    }
    inventory(absolute, state, {
      allowMissing: true,
      allowSymlinks: false,
      inventoryRoot: existsSync(absolute) ? realpathSync(absolute) : absolute,
    })
  }
}

function parseCli(argv) {
  let mode = 'source'
  const artifactDirectories = []
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === '--mode' && argv[index + 1]) mode = argv[++index]
    else if (argv[index] === '--artifact-dir' && argv[index + 1])
      artifactDirectories.push(argv[++index])
    else throw new Error(`unknown or incomplete argument: ${argv[index]}`)
  }
  return { mode, artifactDirectories }
}

if (process.argv[1] && realpathSync(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const options = parseCli(process.argv.slice(2))
  assertOptionalRuntimePackagingPolicy({ root: process.cwd(), ...options })
}
