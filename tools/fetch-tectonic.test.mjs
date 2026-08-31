import assert from 'node:assert/strict'
import { createHash } from 'node:crypto'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import {
  extractVerifiedTectonic,
  fetchVerifiedAsset,
  parseArguments,
  publishExecutable,
  selectPlatformAsset,
  syncDirectory,
} from './fetch-tectonic.mjs'

const asset = {
  id: 'tectonic-test-darwin-arm64',
  archive: { format: 'tar.gz', executable: 'tectonic' },
}

const windowsAsset = {
  id: 'tectonic-test-win32-x64',
  archive: { format: 'zip', executable: 'tectonic.exe' },
}

test('CLI accepts an explicit output path for the verified executable', () => {
  const parsed = parseArguments(['--platform', 'darwin-arm64', '--output', '/tmp/wiswork-tectonic'])
  assert.equal(parsed.platform, 'darwin-arm64')
  assert.equal(parsed.outputPath, '/tmp/wiswork-tectonic')
})

test('publishExecutable copies only a regular verified executable to the explicit output', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'publish-tectonic-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const source = join(root, 'cache', 'tectonic')
  const output = join(root, 'ci', 'tectonic')
  await mkdir(join(root, 'cache'))
  await writeFile(source, 'verified-binary')
  await publishExecutable(source, output)
  assert.equal(await readFile(output, 'utf8'), 'verified-binary')
})

test('syncDirectory skips unsupported Windows directory handles', async () => {
  let opened = false
  await syncDirectory('C:\\runner\\cache', {
    platform: 'win32',
    openImplementation: async () => {
      opened = true
      throw new Error('Windows directories cannot be opened for fsync')
    },
  })
  assert.equal(opened, false)
})

test('extractVerifiedTectonic validates one executable and its exact version', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'fetch-tectonic-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const archive = join(root, 'tectonic.tar.gz')
  await writeFile(archive, 'fixture')
  const execute = async (file, args) => {
    if (file === 'tar' && args[0] === '-tzf') return { stdout: 'tectonic\n', stderr: '' }
    if (file === 'tar' && args[0] === '-xzf') {
      const destination = args[3]
      await mkdir(destination, { recursive: true })
      await writeFile(join(destination, 'tectonic'), 'binary')
      return { stdout: '', stderr: '' }
    }
    if (file.endsWith('/tectonic') && args[0] === '--version') {
      return { stdout: 'tectonic 0.16.9\n', stderr: '' }
    }
    throw new Error('unexpected command')
  }
  const executable = await extractVerifiedTectonic(asset, archive, root, '0.16.9', {
    execFileImplementation: execute,
  })
  assert.equal(await readFile(executable, 'utf8'), 'binary')
})

test('extractVerifiedTectonic validates the single Windows ZIP executable', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'fetch-tectonic-windows-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const archive = join(root, 'tectonic.zip')
  await writeFile(archive, 'fixture')
  const execute = async (file, args) => {
    if (file === 'tar' && args[0] === '-tf') return { stdout: 'tectonic.exe\n', stderr: '' }
    if (file === 'tar' && args[0] === '-xf') {
      await writeFile(join(args[3], 'tectonic.exe'), 'windows-binary')
      return { stdout: '', stderr: '' }
    }
    if (file.endsWith('/tectonic.exe') && args[0] === '--version') {
      return { stdout: 'tectonic 0.16.9\n', stderr: '' }
    }
    throw new Error('unexpected command')
  }

  const executable = await extractVerifiedTectonic(windowsAsset, archive, root, '0.16.9', {
    execFileImplementation: execute,
  })

  assert.equal(await readFile(executable, 'utf8'), 'windows-binary')
})

test('selectPlatformAsset accepts only the pinned Windows ZIP layout', () => {
  const selected = selectPlatformAsset(
    {
      schemaVersion: 1,
      tectonic: {
        version: '0.16.9',
        assets: [
          {
            ...windowsAsset,
            platform: 'win32-x64',
            url: 'https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%400.16.9/tectonic.zip',
            bytes: 20_035_039,
            sha256: 'a'.repeat(64),
          },
        ],
      },
    },
    'win32-x64',
  )

  assert.equal(selected.archive.format, 'zip')
  assert.equal(selected.archive.executable, 'tectonic.exe')
})

test('extractVerifiedTectonic replaces a poisoned version-spoofing cache', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'fetch-tectonic-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const archive = join(root, 'tectonic.tar.gz')
  const target = join(root, asset.id)
  await writeFile(archive, 'fixture')
  await mkdir(target)
  await writeFile(join(target, 'tectonic'), 'poisoned')
  const execute = async (file, args) => {
    if (file === 'tar' && args[0] === '-tzf') return { stdout: 'tectonic\n', stderr: '' }
    if (file === 'tar' && args[0] === '-xzf') {
      await writeFile(join(args[3], 'tectonic'), 'verified-archive-binary')
      return { stdout: '', stderr: '' }
    }
    return { stdout: 'tectonic 0.16.9\n', stderr: '' }
  }
  const executable = await extractVerifiedTectonic(asset, archive, root, '0.16.9', {
    execFileImplementation: execute,
  })
  assert.equal(await readFile(executable, 'utf8'), 'verified-archive-binary')
})

test('extractVerifiedTectonic rejects traversal or extra archive entries', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'fetch-tectonic-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const archive = join(root, 'tectonic.tar.gz')
  await writeFile(archive, 'fixture')
  await assert.rejects(
    extractVerifiedTectonic(asset, archive, root, '0.16.9', {
      execFileImplementation: async () => ({ stdout: 'tectonic\n../escape\n', stderr: '' }),
    }),
    /layout/i,
  )
})

test('selectPlatformAsset rejects path traversal IDs', () => {
  assert.throws(
    () =>
      selectPlatformAsset(
        {
          schemaVersion: 1,
          tectonic: {
            version: '0.16.9',
            assets: [
              {
                ...asset,
                id: '../../victim',
                platform: 'darwin-arm64',
                url: 'https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%400.16.9/fixed.tar.gz',
                bytes: 1,
                sha256: 'a'.repeat(64),
              },
            ],
          },
        },
        'darwin-arm64',
      ),
    /invalid/i,
  )
})

test('fetchVerifiedAsset stops oversized responses', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'fetch-tectonic-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const target = join(root, 'asset.tar.gz')
  await assert.rejects(
    fetchVerifiedAsset(
      { url: 'https://github.com/example/fixed', bytes: 3, sha256: 'a'.repeat(64) },
      target,
      { fetchImplementation: async () => new Response('too-large'), maxAttempts: 1 },
    ),
    /expected size/i,
  )
})

test('fetchVerifiedAsset retries a transient download and still verifies integrity', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'fetch-tectonic-retry-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const target = join(root, 'asset.tar.gz')
  const body = 'verified-archive'
  const retries = []
  let attempts = 0

  const result = await fetchVerifiedAsset(
    {
      url: 'https://github.com/example/fixed',
      bytes: Buffer.byteLength(body),
      sha256: createHash('sha256').update(body).digest('hex'),
    },
    target,
    {
      maxAttempts: 2,
      retryDelayMs: 0,
      onRetry: (attempt) => retries.push(attempt),
      fetchImplementation: async () => {
        attempts += 1
        if (attempts === 1) throw new Error('transient upstream failure')
        return new Response(body)
      },
    },
  )

  assert.equal(attempts, 2)
  assert.deepEqual(retries, [2])
  assert.equal(await readFile(target, 'utf8'), body)
  assert.equal(result.bytes, Buffer.byteLength(body))
})

test('extractVerifiedTectonic restores the prior sidecar when publish fails', async (context) => {
  const root = await mkdtemp(join(tmpdir(), 'fetch-tectonic-'))
  context.after(() => rm(root, { recursive: true, force: true }))
  const archive = join(root, 'tectonic.tar.gz')
  const target = join(root, asset.id)
  await writeFile(archive, 'fixture')
  await mkdir(target)
  await writeFile(join(target, 'tectonic'), 'old-safe-binary')
  let renames = 0
  const execute = async (file, args) => {
    if (file === 'tar' && args[0] === '-tzf') return { stdout: 'tectonic\n', stderr: '' }
    if (file === 'tar' && args[0] === '-xzf') {
      await writeFile(join(args[3], 'tectonic'), 'new-binary')
      return { stdout: '', stderr: '' }
    }
    return { stdout: 'tectonic 0.16.9\n', stderr: '' }
  }
  await assert.rejects(
    extractVerifiedTectonic(asset, archive, root, '0.16.9', {
      execFileImplementation: execute,
      renameImplementation: async (from, to) => {
        renames += 1
        if (renames === 2) throw new Error('publish failed')
        const { rename } = await import('node:fs/promises')
        return rename(from, to)
      },
    }),
    /publish failed/,
  )
  assert.equal(await readFile(join(target, 'tectonic'), 'utf8'), 'old-safe-binary')
})
