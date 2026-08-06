import { createHash } from 'node:crypto'
import { access, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { BundleInstaller, type BundleDownload } from '../src/bundle-installer.js'
import type { TectonicBundleAsset } from '../src/manifest.js'

const payload = Buffer.from('verified bundle')
const asset: TectonicBundleAsset = {
  id: 'bundle-v33',
  url: 'https://relay.fullyjustified.net/default_bundle_v33.tar?token=secret',
  bytes: payload.byteLength,
  sha256: createHash('sha256').update(payload).digest('hex'),
  license: { spdx: 'LicenseRef-Tectonic-Bundle', sourceUrl: 'https://relay.fullyjustified.net/' },
}

describe('BundleInstaller', () => {
  const roots: string[] = []
  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function setup(download: BundleDownload, log = vi.fn()) {
    const root = await mkdtemp(join(tmpdir(), 'latex-bundle-'))
    roots.push(root)
    return { root, log, installer: new BundleInstaller(root, asset, { download, log }) }
  }

  it('moves through missing/downloading/ready and atomically publishes verified bytes', async () => {
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    const { root, installer } = await setup(async ({ destination, onBytes }) => {
      await gate
      await writeFile(destination, payload)
      onBytes(payload.byteLength)
    })
    await expect(installer.status()).resolves.toMatchObject({ state: 'missing' })
    const pending = installer.install()
    expect(installer.current).toMatchObject({ state: 'downloading', receivedBytes: 0 })
    release()
    const result = await pending
    expect(result.path).toBe(join(root, `${asset.id}.tar`))
    await expect(readFile(result.path)).resolves.toEqual(payload)
    await expect(installer.status()).resolves.toMatchObject({ state: 'ready', path: result.path })
    expect((await readdir(root)).filter((name) => name.includes('.part'))).toEqual([])
  })

  it('rejects tampered or truncated content and never marks it ready', async () => {
    const { root, installer } = await setup(async ({ destination }) => {
      await writeFile(destination, 'tampered')
    })
    await expect(installer.install()).rejects.toMatchObject({ code: 'BUNDLE_INTEGRITY_FAILED' })
    expect(installer.current).toMatchObject({ state: 'error', code: 'BUNDLE_INTEGRITY_FAILED' })
    await expect(access(join(root, `${asset.id}.tar`))).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('supports cancellation, retains a resumable temporary file, and retries', async () => {
    let attempts = 0
    let partialWritten!: () => void
    const partial = new Promise<void>((resolve) => (partialWritten = resolve))
    const download: BundleDownload = async ({ destination, offset, signal }) => {
      attempts += 1
      if (attempts === 1) {
        await writeFile(destination, payload.subarray(0, 4))
        partialWritten()
        await new Promise<void>((_resolve, reject) =>
          signal.addEventListener('abort', () => reject(signal.reason), { once: true }),
        )
      } else {
        expect(offset).toBe(4)
        await writeFile(destination, payload.subarray(offset), { flag: 'a' })
      }
    }
    const { installer } = await setup(download)
    const pending = installer.install()
    await partial
    installer.cancel()
    await expect(pending).rejects.toMatchObject({ code: 'BUNDLE_DOWNLOAD_CANCELLED' })
    expect(installer.current.state).toBe('missing')
    await expect(installer.install()).resolves.toMatchObject({ bytes: payload.byteLength })
    expect(attempts).toBe(2)
  })

  it('single-flights concurrent installs', async () => {
    let calls = 0
    const { installer } = await setup(async ({ destination }) => {
      calls += 1
      await writeFile(destination, payload)
    })
    const first = installer.install()
    const second = installer.install()
    expect(second).toBe(first)
    await Promise.all([first, second])
    expect(calls).toBe(1)
  })

  it('single-flights across installer instances sharing one target', async () => {
    let calls = 0
    const root = await mkdtemp(join(tmpdir(), 'latex-bundle-'))
    roots.push(root)
    const download: BundleDownload = async ({ destination }) => {
      calls += 1
      await writeFile(destination, payload)
    }
    const first = new BundleInstaller(root, asset, { download })
    const second = new BundleInstaller(root, asset, { download })
    await Promise.all([first.install(), second.install()])
    expect(calls).toBe(1)
  })

  it('shares downloading state and cancellation across installer instances', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-bundle-'))
    roots.push(root)
    let started!: () => void
    const download: BundleDownload = async ({ signal }) => {
      started()
      await new Promise<void>((_resolve, reject) =>
        signal.addEventListener('abort', () => reject(signal.reason), { once: true }),
      )
    }
    const first = new BundleInstaller(root, asset, { download })
    const second = new BundleInstaller(root, asset, { download })
    const began = new Promise<void>((resolve) => (started = resolve))
    const pending = first.install()
    await began
    const shared = second.install()
    await expect(second.status()).resolves.toMatchObject({ state: 'downloading' })
    second.cancel()
    await expect(Promise.all([pending, shared])).rejects.toMatchObject({
      code: 'BUNDLE_DOWNLOAD_CANCELLED',
    })
  })

  it('publishes an already complete verified temporary file without downloading', async () => {
    const download = vi.fn<BundleDownload>()
    const { root, installer } = await setup(download)
    await writeFile(join(root, `.${asset.id}.tar.part`), payload)
    await expect(installer.install()).resolves.toMatchObject({ bytes: payload.byteLength })
    expect(download).not.toHaveBeenCalled()
  })

  it('reports ready when directory sync fails after a verified rename committed', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-bundle-'))
    roots.push(root)
    const installer = new BundleInstaller(root, asset, {
      download: async ({ destination }) => writeFile(destination, payload),
      syncDirectory: async () => {
        throw new Error('directory fsync unsupported')
      },
    })
    await expect(installer.install()).resolves.toMatchObject({ bytes: payload.byteLength })
    await expect(readFile(installer.targetPath)).resolves.toEqual(payload)
  })

  it('uses stable redacted logs without URL or token data', async () => {
    const { installer, log } = await setup(async () => {
      throw new Error(`network failed for ${asset.url}`)
    })
    await expect(installer.install()).rejects.toMatchObject({ code: 'BUNDLE_DOWNLOAD_FAILED' })
    const serialized = JSON.stringify(log.mock.calls)
    expect(serialized).not.toContain('https://')
    expect(serialized).not.toContain('secret')
    expect(serialized).toContain(asset.id)
    expect(serialized).toContain('BUNDLE_DOWNLOAD_FAILED')
  })
})
