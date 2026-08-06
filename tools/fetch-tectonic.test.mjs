import assert from 'node:assert/strict'
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import {
  extractVerifiedTectonic,
  fetchVerifiedAsset,
  selectPlatformAsset,
} from './fetch-tectonic.mjs'

const asset = {
  id: 'tectonic-test-darwin-arm64',
  archive: { format: 'tar.gz', executable: 'tectonic' },
}

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
      { fetchImplementation: async () => new Response('too-large') },
    ),
    /expected size/i,
  )
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
