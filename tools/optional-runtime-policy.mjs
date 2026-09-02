import { createRequire } from 'node:module'
import { createHash } from 'node:crypto'
import {
  closeSync,
  existsSync,
  lstatSync,
  openSync,
  opendirSync,
  readFileSync,
  readlinkSync,
  readSync,
  realpathSync,
} from 'node:fs'
import { basename, dirname, isAbsolute, join, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const require = createRequire(import.meta.url)
const MAX_INVENTORY_ENTRIES = 20_000
const MAX_INVENTORY_DEPTH = 16
const MAX_INVENTORY_BYTES = 20 * 1024 * 1024 * 1024
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

function executableMagic(path) {
  const handle = openSync(path, 'r')
  const prefix = Buffer.alloc(4)
  try {
    const count = readSync(handle, prefix, 0, prefix.length, 0)
    if (count < 2) return false
    if (prefix[0] === 0x4d && prefix[1] === 0x5a) return true
    if (count < 4) return false
    const magic = prefix.readUInt32BE(0)
    return (
      magic === 0x7f454c46 ||
      magic === 0xfeedface ||
      magic === 0xfeedfacf ||
      magic === 0xcefaedfe ||
      magic === 0xcffaedfe ||
      magic === 0xcafebabe ||
      magic === 0xbebafeca
    )
  } finally {
    closeSync(handle)
  }
}

function sha256File(path) {
  const hash = createHash('sha256')
  const handle = openSync(path, 'r')
  const chunk = Buffer.allocUnsafe(1024 * 1024)
  try {
    for (;;) {
      const count = readSync(handle, chunk, 0, chunk.length, null)
      if (count === 0) break
      hash.update(chunk.subarray(0, count))
    }
  } finally {
    closeSync(handle)
  }
  return hash.digest('hex')
}

function inventory(
  path,
  state,
  { allowMissing, allowSymlinks, inventoryRoot, allowExecutableMagic },
  depth = 0,
) {
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
      { allowMissing: false, allowSymlinks, inventoryRoot, allowExecutableMagic },
      depth + 1,
    )
  }
  if (!info.isDirectory()) {
    if (!info.isFile()) throw new Error(`package inventory contains special file: ${path}`)
    state.bytes += info.size
    if (state.bytes > MAX_INVENTORY_BYTES) throw new Error('package inventory exceeds byte limit')
    if (executableMagic(path) && !allowExecutableMagic) {
      throw new Error(`executable package input is not allowlisted: ${path}`)
    }
    if (state.knownHashes.has(sha256File(path))) {
      throw new Error(`known optional Codex component bytes must not be bundled: ${path}`)
    }
    return
  }
  if (/\.app$/i.test(basename(path)) || /(?:^|-)unpacked$/i.test(basename(path))) {
    state.sawUnpackedRoot = true
  }
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
        { allowMissing: false, allowSymlinks, inventoryRoot, allowExecutableMagic },
        depth + 1,
      )
    }
  } finally {
    directory.closeSync()
  }
}

export function optionalRuntimeKnownHashes(root) {
  const path = join(root, 'tools/codex/manifest.json')
  if (!existsSync(path)) return new Set()
  const manifest = JSON.parse(readFileSync(path, 'utf8'))
  const hashes = new Set()
  for (const asset of manifest?.component?.assets ?? []) {
    if (typeof asset.sha256 === 'string') hashes.add(asset.sha256)
    for (const file of asset?.layout?.files ?? []) {
      if (file?.install === true && typeof file.sha256 === 'string') hashes.add(file.sha256)
    }
  }
  return hashes
}

function executableInputAllowed(source, shellRoot) {
  const candidate = resolve(shellRoot, source)
  return new Set([
    resolve(shellRoot, '../sheets/native/xlsx-engine/target/release/xlsx-sidecar'),
    resolve(shellRoot, '../sheets/native/xlsx-engine/target/release/xlsx-sidecar.exe'),
    resolve(shellRoot, '../latex/native/tectonic'),
    resolve(shellRoot, '../latex/native/tectonic.exe'),
    resolve(shellRoot, '../latex/native/tectonic-ci'),
    resolve(shellRoot, '../latex/native/tectonic-ci.exe'),
  ]).has(candidate)
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
  knownHashes,
} = {}) {
  if (typeof root !== 'string') throw new TypeError('root is required')
  if (mode !== 'source' && mode !== 'post-package') throw new TypeError('invalid policy mode')
  const resolvedKnownHashes = knownHashes ?? optionalRuntimeKnownHashes(root)
  if (
    !(resolvedKnownHashes instanceof Set) ||
    [...resolvedKnownHashes].some(
      (hash) => typeof hash !== 'string' || !/^[a-f0-9]{64}$/.test(hash),
    )
  ) {
    throw new TypeError('knownHashes must be a Set of SHA-256 digests')
  }
  const state = {
    bytes: 0,
    entries: 0,
    visited: new Set(),
    knownHashes: resolvedKnownHashes,
    sawUnpackedRoot: false,
  }
  if (mode === 'post-package') {
    if (!Array.isArray(artifactDirectories) || artifactDirectories.length === 0) {
      throw new Error('post-package mode requires artifact directories')
    }
    // ASAR/DMG/ZIP/NSIS files are treated as opaque, bounded regular files and hashed as a whole.
    // Their inputs are independently constrained by source mode; release jobs must also provide
    // electron-builder's unpacked output in the same artifact directory for recursive inspection.
    for (const directory of artifactDirectories) {
      const absolute = resolve(directory)
      inventory(absolute, state, {
        allowMissing: false,
        allowSymlinks: true,
        inventoryRoot: absolute,
        allowExecutableMagic: true,
      })
    }
    if (!state.sawUnpackedRoot) {
      throw new Error('post-package inventory requires an unpacked application artifact')
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
      allowExecutableMagic: executableInputAllowed(source, shellRoot),
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
