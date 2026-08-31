import assert from 'node:assert/strict'
import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync } from 'node:fs'
import { mkdtemp, mkdir, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import test from 'node:test'
import { load } from 'js-yaml'
import { writeReleaseMetadata } from './merge-update-metadata.mjs'

const root = join(import.meta.dirname, '..', '..')
const tool = join(root, 'tools/release/merge-update-metadata.mjs')

function digest(algorithm, value) {
  return createHash(algorithm).update(value).digest('base64')
}

function fileRecord(url, contents) {
  return {
    url,
    sha512: digest('sha512', contents),
    size: Buffer.byteLength(contents),
  }
}

function metadata(record, version = '0.6.1') {
  return [
    `version: ${version}`,
    'files:',
    `  - url: ${record.url}`,
    `    sha512: ${record.sha512}`,
    `    size: ${record.size}`,
    `path: ${record.url}`,
    `sha512: ${record.sha512}`,
    "releaseDate: '2026-08-30T00:00:00.000Z'",
    '',
  ].join('\n')
}

async function fixture(context) {
  const directory = await mkdtemp(join(tmpdir(), 'merge-update-metadata-'))
  context.after(() => rm(directory, { recursive: true, force: true }))

  const artifactsDir = join(directory, 'artifacts')
  const metadataDir = join(directory, 'metadata')
  const outputDir = join(directory, 'publish')
  await Promise.all([mkdir(artifactsDir), mkdir(metadataDir)])

  const contents = {
    windows: 'windows installer',
    arm64: 'arm64 archive',
    x64: 'x64 archive',
  }
  const records = {
    windows: fileRecord('WisWork-Setup-0.6.1.exe', contents.windows),
    arm64: fileRecord('WisWork-0.6.1-arm64.zip', contents.arm64),
    x64: fileRecord('WisWork-0.6.1.zip', contents.x64),
  }
  await Promise.all(
    Object.entries(records).map(([name, record]) =>
      writeFile(join(artifactsDir, record.url), contents[name]),
    ),
  )

  const paths = {
    windows: join(metadataDir, 'latest.yml'),
    arm64: join(metadataDir, 'latest-arm64.yml'),
    x64: join(metadataDir, 'latest-x64.yml'),
  }
  await Promise.all([
    writeFile(paths.windows, metadata(records.windows)),
    writeFile(paths.arm64, metadata(records.arm64)),
    writeFile(paths.x64, metadata(records.x64)),
  ])

  return { artifactsDir, outputDir, paths, records }
}

function run({ artifactsDir, outputDir, paths, additionalArtifacts = [] }) {
  const additionalArgs = additionalArtifacts.flatMap((path) => ['--additional-artifact', path])
  return spawnSync(
    process.execPath,
    [
      tool,
      '--windows',
      paths.windows,
      '--mac-arm64',
      paths.arm64,
      '--mac-x64',
      paths.x64,
      '--artifacts-dir',
      artifactsDir,
      '--output-dir',
      outputDir,
      '--version',
      '0.6.1',
      ...additionalArgs,
    ],
    { cwd: root, encoding: 'utf8' },
  )
}

async function yaml(path) {
  return load(await readFile(path, 'utf8'))
}

test('merges architecture-local macOS files into one metadata document', async (context) => {
  const options = await fixture(context)

  const result = run(options)

  assert.equal(result.status, 0, result.stderr)
  const mac = await yaml(join(options.outputDir, 'latest-mac.yml'))
  assert.equal(mac.version, '0.6.1')
  assert.deepEqual(
    mac.files.map(({ url }) => url),
    ['WisWork-0.6.1-arm64.zip', 'WisWork-0.6.1.zip'],
  )
  assert.equal(mac.path, 'WisWork-0.6.1-arm64.zip')
  assert.equal(mac.sha512, options.records.arm64.sha512)
})

test('preserves and validates the Windows latest.yml record', async (context) => {
  const options = await fixture(context)

  const result = run(options)

  assert.equal(result.status, 0, result.stderr)
  const windows = await yaml(join(options.outputDir, 'latest.yml'))
  assert.equal(windows.files[0].url, 'WisWork-Setup-0.6.1.exe')
  assert.equal(windows.files[0].sha512, options.records.windows.sha512)
  const sums = await readFile(join(options.outputDir, 'SHA256SUMS.txt'), 'utf8')
  assert.equal(
    sums,
    [
      `${createHash('sha256').update('windows installer').digest('hex')}  WisWork-Setup-0.6.1.exe`,
      `${createHash('sha256').update('arm64 archive').digest('hex')}  WisWork-0.6.1-arm64.zip`,
      `${createHash('sha256').update('x64 archive').digest('hex')}  WisWork-0.6.1.zip`,
      '',
    ].join('\n'),
  )
})

test('includes explicit additional release artifacts in SHA256SUMS', async (context) => {
  const options = await fixture(context)
  const arm64Dmg = 'WisWork-0.6.1-arm64.dmg'
  const x64Dmg = 'WisWork-0.6.1.dmg'
  await Promise.all([
    writeFile(join(options.artifactsDir, arm64Dmg), 'arm64 disk image'),
    writeFile(join(options.artifactsDir, x64Dmg), 'x64 disk image'),
  ])
  options.additionalArtifacts = [arm64Dmg, x64Dmg]

  const result = run(options)

  assert.equal(result.status, 0, result.stderr)
  const sums = await readFile(join(options.outputDir, 'SHA256SUMS.txt'), 'utf8')
  assert.equal(
    sums,
    [
      `${createHash('sha256').update('windows installer').digest('hex')}  WisWork-Setup-0.6.1.exe`,
      `${createHash('sha256').update('arm64 archive').digest('hex')}  WisWork-0.6.1-arm64.zip`,
      `${createHash('sha256').update('x64 archive').digest('hex')}  WisWork-0.6.1.zip`,
      `${createHash('sha256').update('arm64 disk image').digest('hex')}  WisWork-0.6.1-arm64.dmg`,
      `${createHash('sha256').update('x64 disk image').digest('hex')}  WisWork-0.6.1.dmg`,
      '',
    ].join('\n'),
  )
})

test('rejects additional release artifacts that collide with updater metadata', async (context) => {
  const options = await fixture(context)
  options.additionalArtifacts = [options.records.windows.url]

  const result = run(options)

  assert.notEqual(result.status, 0)
  assert.match(result.stderr, /duplicate/i)
  assert.equal(existsSync(join(options.outputDir, 'latest.yml')), false)
})

test('rejects mismatched metadata version before writing output', async (context) => {
  const options = await fixture(context)
  await writeFile(options.paths.x64, metadata(options.records.x64, '0.6.2'))

  const result = run(options)

  assert.notEqual(result.status, 0)
  assert.match(result.stderr, /version/i)
  assert.equal(existsSync(join(options.outputDir, 'latest.yml')), false)
  assert.equal(existsSync(join(options.outputDir, 'latest-mac.yml')), false)
  assert.equal(existsSync(join(options.outputDir, 'SHA256SUMS.txt')), false)
})

test('rejects metadata that references a missing artifact or duplicate macOS URL', async (context) => {
  const missing = await fixture(context)
  await rm(join(missing.artifactsDir, missing.records.x64.url))

  const missingResult = run(missing)

  assert.notEqual(missingResult.status, 0)
  assert.match(missingResult.stderr, /artifact/i)
  assert.equal(existsSync(join(missing.outputDir, 'latest.yml')), false)

  const duplicate = await fixture(context)
  const x64 = metadata(duplicate.records.x64).replace(
    'path: WisWork-0.6.1.zip',
    `  - url: ${duplicate.records.arm64.url}\n    sha512: ${duplicate.records.arm64.sha512}\n    size: ${duplicate.records.arm64.size}\npath: WisWork-0.6.1.zip`,
  )
  await writeFile(duplicate.paths.x64, x64)

  const duplicateResult = run(duplicate)

  assert.notEqual(duplicateResult.status, 0)
  assert.match(duplicateResult.stderr, /duplicate/i)
  assert.equal(existsSync(join(duplicate.outputDir, 'latest.yml')), false)
})

test('removes staged metadata and leaves no final output when a staging write fails', async (context) => {
  const options = await fixture(context)

  await assert.rejects(
    writeReleaseMetadata({
      windows: options.paths.windows,
      macArm64: options.paths.arm64,
      macX64: options.paths.x64,
      artifactsDir: options.artifactsDir,
      outputDir: options.outputDir,
      version: '0.6.1',
      writeFileImplementation: async (path, contents) => {
        if (path.endsWith('latest-mac.yml')) throw new Error('staging write failed')
        return writeFile(path, contents)
      },
    }),
    /staging write failed/,
  )

  assert.equal(existsSync(options.outputDir), false)
  assert.equal(
    (await readdir(dirname(options.outputDir))).some((name) => name.startsWith('.publish.staging-')),
    false,
  )
})

test('removes staged metadata and leaves no partial final output when the final rename fails', async (context) => {
  const options = await fixture(context)

  await assert.rejects(
    writeReleaseMetadata({
      windows: options.paths.windows,
      macArm64: options.paths.arm64,
      macX64: options.paths.x64,
      artifactsDir: options.artifactsDir,
      outputDir: options.outputDir,
      version: '0.6.1',
      renameImplementation: async (from, to) => {
        if (to === options.outputDir) throw new Error('final rename failed')
        return rename(from, to)
      },
    }),
    /final rename failed/,
  )

  assert.equal(existsSync(options.outputDir), false)
  assert.equal(
    (await readdir(dirname(options.outputDir))).some((name) => name.startsWith('.publish.staging-')),
    false,
  )
})
