import { createHash } from 'node:crypto'
import { chmod, lstat, mkdir, readFile, rm, symlink, utimes, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Readable } from 'node:stream'
import { mkdtemp } from 'node:fs/promises'
import { create as createTar } from 'tar'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  EnhancedModeComponentError,
  EnhancedModeComponentManager,
  parseCodexComponentManifest,
  validateCodexArchiveEntry,
  type CodexComponentManifest,
} from '../src/component-manager.js'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function temporaryRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), 'wiswork-enhanced-component-test-'))
  roots.push(root)
  return root
}

async function sha256(path: string): Promise<string> {
  return createHash('sha256')
    .update(await readFile(path))
    .digest('hex')
}

async function createFixtureArchive(options: { symlink?: boolean; duplicate?: boolean } = {}) {
  const root = await temporaryRoot()
  const source = join(root, 'source')
  await mkdir(join(source, 'bin'), { recursive: true })
  await writeFile(join(source, 'bin', 'codex'), '#!/bin/sh\necho codex-cli 0.147.0\n')
  await chmod(join(source, 'bin', 'codex'), 0o755)
  await writeFile(
    join(source, 'codex-package.json'),
    JSON.stringify({
      layoutVersion: 1,
      version: '0.147.0',
      target: 'aarch64-apple-darwin',
      entrypoint: 'bin/codex',
    }),
  )
  if (options.symlink) await symlink('codex', join(source, 'bin', 'unsafe-link'))
  const archive = join(root, 'component.tar.gz')
  const paths = ['bin/', 'bin/codex', 'codex-package.json']
  if (options.symlink) paths.push('bin/unsafe-link')
  if (options.duplicate) paths.push('bin/codex')
  await createTar(
    { cwd: source, file: archive, gzip: true, portable: true, noDirRecurse: true },
    paths,
  )

  const executablePath = join(source, 'bin', 'codex')
  const packagePath = join(source, 'codex-package.json')
  const archiveInfo = await lstat(archive)
  const executableInfo = await lstat(executablePath)
  const packageInfo = await lstat(packagePath)
  const manifest: CodexComponentManifest = {
    schemaVersion: 1,
    component: {
      version: '0.147.0',
      license: {
        spdx: 'Apache-2.0',
        sourceUrl: 'https://github.com/openai/codex/blob/rust-v0.147.0/LICENSE',
      },
      assets: [
        {
          id: 'codex-cli-0.147.0-darwin-arm64',
          platform: 'darwin',
          arch: 'arm64',
          target: 'aarch64-apple-darwin',
          url: 'https://github.com/openai/codex/releases/download/rust-v0.147.0/codex-package-aarch64-apple-darwin.tar.gz',
          bytes: archiveInfo.size,
          sha256: await sha256(archive),
          archive: { format: 'tar.gz' },
          layout: {
            entrypoint: 'bin/codex',
            directories: ['bin'],
            files: [
              {
                path: 'bin/codex',
                bytes: executableInfo.size,
                sha256: await sha256(executablePath),
                mode: 'executable',
                install: true,
              },
              {
                path: 'codex-package.json',
                bytes: packageInfo.size,
                sha256: await sha256(packagePath),
                mode: 'data',
                install: true,
              },
            ],
          },
        },
      ],
    },
  }
  return { archive, manifest }
}

function responseFor(bytes: Buffer): Response {
  return new Response(Readable.toWeb(Readable.from(bytes)) as ReadableStream)
}

describe('optional Enhanced mode component manager', () => {
  it('locks the repository manifest to the inspected official archive and minimal runtime subset', async () => {
    const manifestPath = join(
      import.meta.dirname,
      '..',
      '..',
      '..',
      'tools',
      'codex',
      'manifest.json',
    )
    const parsed = parseCodexComponentManifest(JSON.parse(await readFile(manifestPath, 'utf8')))
    const asset = parsed.component.assets[0]!
    expect(asset).toMatchObject({
      platform: 'darwin',
      arch: 'arm64',
      bytes: 107_229_164,
      sha256: '17b2984eb22b607e3d0c25728252fc90f510e476bad39a6d9f45cdb1aa685432',
    })
    expect(asset.layout.files.filter((file) => file.install).map((file) => file.path)).toEqual([
      'bin/codex',
      'bin/codex-code-mode-host',
    ])
    expect(asset.layout.files.reduce((total, file) => total + file.bytes, 0)).toBe(274_773_992)
  })

  it('parses only the exact pinned release metadata and supported darwin-arm64 asset', async () => {
    const { manifest } = await createFixtureArchive()
    const parsed = parseCodexComponentManifest(manifest)
    expect(parsed.component.version).toBe('0.147.0')
    expect(parsed.component.assets).toHaveLength(1)

    expect(() =>
      parseCodexComponentManifest({
        ...manifest,
        component: { ...manifest.component, version: 'latest' },
      }),
    ).toThrowError('enhanced_mode_manifest_invalid')
    expect(() =>
      parseCodexComponentManifest({
        ...manifest,
        component: {
          ...manifest.component,
          assets: [{ ...manifest.component.assets[0]!, platform: 'win32' }],
        },
      }),
    ).toThrowError('enhanced_mode_manifest_invalid')
    expect(() =>
      parseCodexComponentManifest({
        ...manifest,
        component: {
          ...manifest.component,
          assets: [
            {
              ...manifest.component.assets[0]!,
              url: 'https://example.com/codex.tar.gz',
            },
          ],
        },
      }),
    ).toThrowError('enhanced_mode_manifest_invalid')
  })

  it('rejects traversal, absolute, link, device, and duplicate archive entries', () => {
    const seen = new Set<string>()
    expect(() => validateCodexArchiveEntry('../bin/codex', 'File', 1, seen)).toThrow(
      'enhanced_mode_archive_unsafe',
    )
    expect(() => validateCodexArchiveEntry('/bin/codex', 'File', 1, seen)).toThrow(
      'enhanced_mode_archive_unsafe',
    )
    expect(() => validateCodexArchiveEntry('bin\\codex', 'File', 1, seen)).toThrow(
      'enhanced_mode_archive_unsafe',
    )
    expect(() => validateCodexArchiveEntry('bin/link', 'SymbolicLink', 0, seen)).toThrow(
      'enhanced_mode_archive_unsafe',
    )
    expect(() => validateCodexArchiveEntry('bin/device', 'CharacterDevice', 0, seen)).toThrow(
      'enhanced_mode_archive_unsafe',
    )
    validateCodexArchiveEntry('bin/codex', 'File', 1, seen)
    expect(() => validateCodexArchiveEntry('bin/codex', 'File', 1, seen)).toThrow(
      'enhanced_mode_archive_unsafe',
    )
  })

  it('streams, verifies, atomically installs, probes, and resolves the pinned component', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const archiveBytes = await readFile(archive)
    const fetchImplementation = vi.fn(async () => responseFor(archiveBytes))
    const probeVersion = vi.fn(async () => 'codex-cli 0.147.0')
    const verifyPlatformTrust = vi.fn(async () => undefined)
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation,
      probeVersion,
      verifyPlatformTrust,
    })

    await expect(manager.status()).resolves.toMatchObject({ state: 'missing', supported: true })
    const installed = await manager.install()
    expect(installed.executablePath).toBe(
      join(cacheRoot, '0.147.0', 'darwin-arm64', 'bin', 'codex'),
    )
    expect(fetchImplementation).toHaveBeenCalledTimes(1)
    expect(probeVersion).toHaveBeenCalledWith(installed.executablePath, expect.anything())
    expect(verifyPlatformTrust).toHaveBeenCalledWith([installed.executablePath], expect.anything())
    await expect(manager.status()).resolves.toMatchObject({ state: 'ready', supported: true })
    await expect(manager.resolveExecutable()).resolves.toBe(installed.executablePath)
    expect((await lstat(installed.executablePath)).isSymbolicLink()).toBe(false)
  })

  it('coalesces concurrent installs and garbage-collects only old version directories', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const archiveBytes = await readFile(archive)
    const fetchImplementation = vi.fn(async () => responseFor(archiveBytes))
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation,
      probeVersion: async () => 'codex-cli 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    await mkdir(join(cacheRoot, '0.146.0', 'darwin-arm64'), { recursive: true })
    await writeFile(join(cacheRoot, 'unowned-file'), 'keep')

    const [first, second] = await Promise.all([manager.install(), manager.install()])
    expect(first).toEqual(second)
    expect(fetchImplementation).toHaveBeenCalledTimes(1)
    await expect(lstat(join(cacheRoot, '0.146.0'))).rejects.toMatchObject({ code: 'ENOENT' })
    await expect(readFile(join(cacheRoot, 'unowned-file'), 'utf8')).resolves.toBe('keep')
  })

  it('recovers an expired valid lock even when its crashed PID has been reused', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const archiveBytes = await readFile(archive)
    await mkdir(cacheRoot, { mode: 0o700 })
    const lockPath = join(cacheRoot, '.install.lock')
    await writeFile(lockPath, JSON.stringify({ pid: process.pid, token: 'a'.repeat(32) }), {
      mode: 0o600,
    })
    const expired = new Date(Date.now() - 1_000)
    await utimes(lockPath, expired, expired)
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(archiveBytes),
      probeVersion: async () => 'codex-cli 0.147.0',
      verifyPlatformTrust: async () => undefined,
      lockTimeoutMs: 100,
      lockStaleMs: 10,
    })

    await expect(manager.install()).resolves.toMatchObject({ version: '0.147.0' })
  })

  it('fails closed for missing, unsupported, symlinked, and tampered components', async () => {
    const root = await temporaryRoot()
    const { archive, manifest } = await createFixtureArchive()
    const archiveBytes = await readFile(archive)
    const manager = new EnhancedModeComponentManager({
      cacheRoot: join(root, 'cache'),
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(archiveBytes),
      probeVersion: async () => 'codex-cli 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    await expect(manager.resolveExecutable()).rejects.toMatchObject({
      code: 'enhanced_mode_install_required',
    })
    await manager.install()
    const executable = await manager.resolveExecutable()
    await writeFile(executable, 'tampered')
    await expect(manager.status()).resolves.toMatchObject({ state: 'invalid' })
    await expect(manager.resolveExecutable()).rejects.toMatchObject({
      code: 'enhanced_mode_integrity_failed',
    })

    const unsupported = new EnhancedModeComponentManager({
      cacheRoot: join(root, 'other'),
      manifest,
      platform: 'win32',
      arch: 'x64',
      fetchImplementation: async () => responseFor(archiveBytes),
      probeVersion: async () => 'codex-cli 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    await expect(unsupported.status()).resolves.toMatchObject({ state: 'unsupported' })
    await expect(unsupported.install()).rejects.toMatchObject({ code: 'enhanced_mode_unsupported' })

    const realCache = join(root, 'real-cache')
    const linkedCache = join(root, 'linked-cache')
    await mkdir(realCache)
    await symlink(realCache, linkedCache)
    const symlinked = new EnhancedModeComponentManager({
      cacheRoot: linkedCache,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(archiveBytes),
      probeVersion: async () => 'codex-cli 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    await expect(symlinked.status()).resolves.toMatchObject({ state: 'invalid' })
  })

  it('rejects unexpected nested files and relaxed installed executable modes', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const archiveBytes = await readFile(archive)
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(archiveBytes),
      probeVersion: async () => 'codex-cli 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    const installed = await manager.install()
    const binDirectory = join(cacheRoot, '0.147.0', 'darwin-arm64', 'bin')

    await writeFile(join(binDirectory, 'unexpected-helper'), 'not allowlisted')
    await expect(manager.status()).resolves.toMatchObject({ state: 'invalid' })
    await rm(join(binDirectory, 'unexpected-helper'))

    await chmod(installed.executablePath, 0o755)
    await expect(manager.resolveExecutable()).rejects.toMatchObject({
      code: 'enhanced_mode_integrity_failed',
    })
  })

  it('rejects unsafe archive content and removes every staging/download artifact', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const fixture = await createFixtureArchive({ symlink: true })
    const bytes = await readFile(fixture.archive)
    const asset = fixture.manifest.component.assets[0]!
    const unsafeManifest: CodexComponentManifest = {
      ...fixture.manifest,
      component: {
        ...fixture.manifest.component,
        assets: [
          {
            ...asset,
            layout: {
              ...asset.layout,
              files: [
                ...asset.layout.files,
                {
                  path: 'bin/unsafe-link',
                  bytes: 5,
                  sha256: '0'.repeat(64),
                  mode: 'data',
                  install: false,
                },
              ],
            },
          },
        ],
      },
    }
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest: unsafeManifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(bytes),
      probeVersion: async () => 'codex-cli 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    await expect(manager.install()).rejects.toBeInstanceOf(EnhancedModeComponentError)
    await expect(manager.status()).resolves.toMatchObject({ state: 'missing' })
    const entries = await import('node:fs/promises').then(({ readdir }) =>
      readdir(cacheRoot).catch(() => []),
    )
    expect(entries.filter((name) => name.includes('.part') || name.includes('.staging'))).toEqual(
      [],
    )
  })

  it('honors cancellation, bounds downloads, retries cleanly, and removes an installed component', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const bytes = await readFile(archive)
    const controller = new AbortController()
    let calls = 0
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => {
        calls += 1
        if (calls === 1) {
          controller.abort()
          return responseFor(Buffer.concat([bytes, Buffer.from('too-large')]))
        }
        return responseFor(bytes)
      },
      probeVersion: async () => 'codex-cli 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })

    await expect(manager.install({ signal: controller.signal })).rejects.toMatchObject({
      code: 'enhanced_mode_cancelled',
    })
    await manager.install()
    await manager.remove()
    await expect(manager.status()).resolves.toMatchObject({ state: 'missing' })
  })

  it('never leaves a published component when cancellation wins after the staging probe', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const bytes = await readFile(archive)
    const controller = new AbortController()
    const probeVersion = vi.fn(async () => {
      controller.abort()
      return 'codex-cli 0.147.0'
    })
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(bytes),
      probeVersion,
      verifyPlatformTrust: async () => undefined,
    })

    await expect(manager.install({ signal: controller.signal })).rejects.toMatchObject({
      code: 'enhanced_mode_cancelled',
    })
    await expect(manager.status()).resolves.toMatchObject({ state: 'missing' })
  })

  it('does not delete a valid existing component when revalidation is cancelled', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const bytes = await readFile(archive)
    const controller = new AbortController()
    let cancelProbe = false
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(bytes),
      probeVersion: async () => {
        if (cancelProbe) {
          controller.abort()
          throw new Error('private cancelled probe detail')
        }
        return 'codex-cli 0.147.0'
      },
      verifyPlatformTrust: async () => undefined,
    })
    await manager.install()
    cancelProbe = true

    await expect(manager.install({ signal: controller.signal })).rejects.toMatchObject({
      code: 'enhanced_mode_cancelled',
    })
    cancelProbe = false
    await expect(manager.status()).resolves.toMatchObject({ state: 'ready' })
  })

  it('rejects an extracted executable whose bounded version probe is not exact', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const bytes = await readFile(archive)
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(bytes),
      probeVersion: async () => 'codex-cli 0.148.0',
      verifyPlatformTrust: async () => undefined,
    })

    await expect(manager.install()).rejects.toMatchObject({
      code: 'enhanced_mode_version_mismatch',
    })
    await expect(manager.status()).resolves.toMatchObject({ state: 'missing' })
  })

  it('does not publish a component that fails the platform trust assessment', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const bytes = await readFile(archive)
    const probeVersion = vi.fn(async () => 'codex-cli 0.147.0')
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(bytes),
      probeVersion,
      verifyPlatformTrust: async () => {
        throw new Error('private signature detail')
      },
    })

    await expect(manager.install()).rejects.toMatchObject({
      code: 'enhanced_mode_platform_trust_failed',
    })
    expect(probeVersion).not.toHaveBeenCalled()
    await expect(manager.status()).resolves.toMatchObject({ state: 'missing' })
  })
})
