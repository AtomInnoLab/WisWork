import { createHash } from 'node:crypto'
import { lstat, mkdir, mkdtemp, readFile, rename, rm, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { basename, dirname, join, resolve, win32 } from 'node:path'
import { dump, load } from 'js-yaml'

const argumentNames = new Map([
  ['--windows', 'windows'],
  ['--mac-arm64', 'macArm64'],
  ['--mac-x64', 'macX64'],
  ['--artifacts-dir', 'artifactsDir'],
  ['--output-dir', 'outputDir'],
  ['--version', 'version'],
])
const additionalArtifactOption = '--additional-artifact'

function fail(message) {
  throw new Error(message)
}

function isObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

function requireDocument(document, name, version) {
  if (!isObject(document)) fail(`${name} metadata root must be an object`)
  if (document.version !== version) fail(`${name} metadata version must equal ${version}`)
  if (!Array.isArray(document.files) || document.files.length === 0) {
    fail(`${name} metadata must contain file records`)
  }
  if (!isBasenameUrl(document.path) || typeof document.sha512 !== 'string') {
    fail(`${name} metadata top-level path or SHA-512 is invalid`)
  }
  if (
    !document.files.some(
      (record) =>
        isObject(record) && record.url === document.path && record.sha512 === document.sha512,
    )
  ) {
    fail(`${name} metadata top-level path and SHA-512 must match a file record`)
  }
}

function requireOneRecord(records, predicate, description) {
  const matches = records.filter((record) => isObject(record) && predicate(record.url))
  if (matches.length !== 1) fail(`${description} metadata must contain exactly one matching record`)
  return matches[0]
}

function assertUniqueUrls(records) {
  const urls = new Set()
  for (const record of records) {
    if (!isObject(record) || typeof record.url !== 'string') {
      fail('metadata file record URL is invalid')
    }
    if (urls.has(record.url)) fail(`duplicate metadata file URL: ${record.url}`)
    urls.add(record.url)
  }
}

function isBasenameUrl(url) {
  return (
    typeof url === 'string' &&
    url.length > 0 &&
    url !== '.' &&
    url !== '..' &&
    url === basename(url) &&
    url === win32.basename(url) &&
    !url.includes('\u0000')
  )
}

export function parseArgs(argv) {
  const parsed = { additionalArtifacts: [] }
  for (let index = 0; index < argv.length; index += 2) {
    const option = argv[index]
    const key = argumentNames.get(option)
    const value = argv[index + 1]
    if (typeof value !== 'string' || value.length === 0) {
      fail('expected exactly six required arguments')
    }
    if (option === additionalArtifactOption) {
      parsed.additionalArtifacts.push(value)
      continue
    }
    if (!key || Object.hasOwn(parsed, key)) fail('expected exactly six required arguments')
    parsed[key] = value
  }
  if (Object.keys(parsed).length !== argumentNames.size + 1) fail('expected exactly six required arguments')
  return parsed
}

export async function validateRecord(record, artifactsDir) {
  if (!isObject(record)) fail('metadata file record must be an object')
  if (!isBasenameUrl(record.url)) fail(`metadata file URL must be a basename: ${record.url}`)
  if (!Number.isSafeInteger(record.size) || record.size < 0) {
    fail(`metadata file size is invalid: ${record.url}`)
  }
  if (typeof record.sha512 !== 'string' || record.sha512.length === 0) {
    fail(`metadata file SHA-512 is invalid: ${record.url}`)
  }

  const artifactPath = join(resolve(artifactsDir), record.url)
  let stats
  try {
    stats = await lstat(artifactPath)
  } catch {
    fail(`metadata artifact is missing: ${record.url}`)
  }
  if (!stats.isFile()) fail(`metadata artifact must be a regular file: ${record.url}`)

  const bytes = await readFile(artifactPath)
  if (bytes.length !== record.size) fail(`metadata artifact size does not match: ${record.url}`)
  const sha512 = createHash('sha512').update(bytes).digest('base64')
  if (sha512 !== record.sha512) fail(`metadata artifact SHA-512 does not match: ${record.url}`)

  return { ...record }
}

async function validateAdditionalArtifact(url, artifactsDir) {
  if (!isBasenameUrl(url)) fail(`additional artifact URL must be a basename: ${url}`)
  const artifactPath = join(resolve(artifactsDir), url)
  let stats
  try {
    stats = await lstat(artifactPath)
  } catch {
    fail(`additional artifact is missing: ${url}`)
  }
  if (!stats.isFile()) fail(`additional artifact must be a regular file: ${url}`)
  return { url }
}

export function mergeMetadata({ windows, macArm64, macX64, version }) {
  requireDocument(windows, 'Windows', version)
  requireDocument(macArm64, 'macOS arm64', version)
  requireDocument(macX64, 'macOS x64', version)

  const windowsRecord = requireOneRecord(
    windows.files,
    (url) => typeof url === 'string' && url.endsWith('.exe'),
    'Windows',
  )
  const arm64Record = requireOneRecord(
    macArm64.files,
    (url) => typeof url === 'string' && url.endsWith('-arm64.zip'),
    'macOS arm64',
  )
  const x64Record = requireOneRecord(
    macX64.files,
    (url) => typeof url === 'string' && url.endsWith('.zip') && !url.endsWith('-arm64.zip'),
    'macOS x64',
  )

  assertUniqueUrls([...windows.files, ...macArm64.files, ...macX64.files])

  return {
    windows: { ...windows, files: windows.files.map((record) => ({ ...record })) },
    mac: {
      ...macArm64,
      files: [{ ...arm64Record }, { ...x64Record }],
      path: arm64Record.url,
      sha512: arm64Record.sha512,
    },
    outputRecords: [windowsRecord, arm64Record, x64Record],
    windowsRecord,
  }
}

async function loadMetadata(path, name) {
  let source
  try {
    source = await readFile(path, 'utf8')
  } catch (error) {
    fail(`could not read ${name} metadata: ${error.message}`)
  }
  let document
  try {
    document = load(source)
  } catch (error) {
    fail(`could not parse ${name} metadata: ${error.message}`)
  }
  if (!isObject(document)) fail(`${name} metadata root must be an object`)
  return document
}

async function pathExists(path) {
  try {
    await lstat(path)
    return true
  } catch {
    return false
  }
}

export async function writeReleaseMetadata(options) {
  const write = options.writeFileImplementation ?? writeFile
  const move = options.renameImplementation ?? rename
  const [windows, macArm64, macX64] = await Promise.all([
    loadMetadata(options.windows, 'Windows'),
    loadMetadata(options.macArm64, 'macOS arm64'),
    loadMetadata(options.macX64, 'macOS x64'),
  ])
  const merged = mergeMetadata({ windows, macArm64, macX64, version: options.version })

  const validatedByUrl = new Map()
  // electron-builder creates latest-mac.yml before Apple notarization is stapled.
  // Stapling mutates DMG bytes, so any DMG records in that intermediate metadata
  // are stale by design. Only ZIP/EXE records are updater authority; DMGs are
  // validated below as final additional release artifacts and receive fresh sums.
  for (const record of merged.outputRecords) {
    validatedByUrl.set(record.url, await validateRecord(record, options.artifactsDir))
  }
  const outputRecords = merged.outputRecords.map((record) => validatedByUrl.get(record.url))
  const outputUrls = new Set(outputRecords.map((record) => record.url))
  const additionalRecords = []
  for (const url of options.additionalArtifacts ?? []) {
    if (outputUrls.has(url)) fail(`duplicate release artifact basename: ${url}`)
    outputUrls.add(url)
    additionalRecords.push(await validateAdditionalArtifact(url, options.artifactsDir))
  }

  const outputDir = resolve(options.outputDir)
  const outputParent = dirname(outputDir)
  if (await pathExists(outputDir)) fail(`output directory already exists: ${outputDir}`)
  await mkdir(outputParent, { recursive: true })
  const stagingDir = await mkdtemp(join(outputParent, `.${basename(outputDir)}.staging-`))

  try {
    const sums = await Promise.all(
      [...outputRecords, ...additionalRecords].map(async (record) => {
        const bytes = await readFile(join(resolve(options.artifactsDir), record.url))
        return `${createHash('sha256').update(bytes).digest('hex')}  ${record.url}`
      }),
    )
    await write(join(stagingDir, 'latest.yml'), dump(merged.windows, { lineWidth: -1 }))
    await write(join(stagingDir, 'latest-mac.yml'), dump(merged.mac, { lineWidth: -1 }))
    await write(join(stagingDir, 'SHA256SUMS.txt'), `${sums.join('\n')}\n`)
    await move(stagingDir, outputDir)
  } catch (error) {
    await rm(stagingDir, { recursive: true, force: true })
    throw error
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    await writeReleaseMetadata(parseArgs(process.argv.slice(2)))
  } catch (error) {
    console.error(error.message)
    process.exitCode = 1
  }
}
