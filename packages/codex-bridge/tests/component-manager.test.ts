import { createHash } from 'node:crypto'
import {
  chmod,
  lstat,
  mkdir,
  readFile,
  realpath,
  rm,
  symlink,
  utimes,
  writeFile,
} from 'node:fs/promises'
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
  platformTrustCommands,
  validateCodexArchiveEntry,
  validateZipEntryMode,
  type ComponentPhaseEvent,
  type CodexComponentManifest,
} from '../src/component-manager.js'

const roots: string[] = []

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
})

async function temporaryRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), 'wiswork-enhanced-component-test-'))
  roots.push(root)
  return realpath(root)
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
  await writeFile(join(source, 'bin', 'codex'), '#!/bin/sh\necho codex-app-server 0.147.0\n')
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
          id: 'codex-app-server-0.147.0-darwin-arm64',
          platform: 'darwin',
          arch: 'arm64',
          target: 'aarch64-apple-darwin',
          primaryUrl:
            'https://downloads.wiswork.com/components/codex/0.147.0/codex-app-server-package-aarch64-apple-darwin.tar.gz',
          fallbackUrl:
            'https://github.com/openai/codex/releases/download/rust-v0.147.0/codex-app-server-package-aarch64-apple-darwin.tar.gz',
          bytes: archiveInfo.size,
          sha256: await sha256(archive),
          archive: { format: 'tar.gz', maxExtractedBytes: executableInfo.size + packageInfo.size },
          trust: { policy: 'macos', teamIdentifier: 'ABCDE12345', requireNotarization: true },
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
  const mutableAssets = manifest.component
    .assets as CodexComponentManifest['component']['assets'][number][]
  const arm64 = mutableAssets[0]!
  mutableAssets.push(
    {
      ...arm64,
      id: 'codex-app-server-0.147.0-darwin-x64',
      arch: 'x64',
      target: 'x86_64-apple-darwin',
      primaryUrl: arm64.primaryUrl.replace('aarch64', 'x86_64'),
      fallbackUrl: arm64.fallbackUrl.replace('aarch64', 'x86_64'),
    },
    {
      ...arm64,
      id: 'codex-app-server-0.147.0-win32-x64',
      platform: 'win32',
      arch: 'x64',
      target: 'x86_64-pc-windows-msvc',
      primaryUrl: arm64.primaryUrl.replace('aarch64-apple-darwin', 'x86_64-pc-windows-msvc'),
      fallbackUrl: arm64.fallbackUrl.replace('aarch64-apple-darwin', 'x86_64-pc-windows-msvc'),
      trust: {
        policy: 'windows',
        publisher: 'OpenAI OpCo, LLC',
        publisherThumbprint: '8B0ADFB840E141DAD3044D2B5AC819873DDE3590',
      },
    },
  )
  return { archive, manifest }
}

function responseFor(bytes: Buffer): Response {
  return new Response(Readable.toWeb(Readable.from(bytes)) as ReadableStream)
}

function storedZip(
  entries: readonly { name: string; contents: Buffer; directory?: boolean; unixMode?: number }[],
): Buffer {
  const locals: Buffer[] = []
  const centrals: Buffer[] = []
  let localOffset = 0
  for (const entry of entries) {
    const name = Buffer.from(entry.name)
    const contents = entry.directory ? Buffer.alloc(0) : entry.contents
    const local = Buffer.alloc(30 + name.length)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt32LE(contents.length, 18)
    local.writeUInt32LE(contents.length, 22)
    local.writeUInt16LE(name.length, 26)
    name.copy(local, 30)
    locals.push(local, contents)
    const central = Buffer.alloc(46 + name.length)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt16LE(0x031e, 4)
    central.writeUInt16LE(20, 6)
    central.writeUInt32LE(contents.length, 20)
    central.writeUInt32LE(contents.length, 24)
    central.writeUInt16LE(name.length, 28)
    central.writeUInt32LE(
      ((entry.unixMode ?? (entry.directory ? 0o40755 : 0o100644)) << 16) >>> 0,
      38,
    )
    central.writeUInt32LE(localOffset, 42)
    name.copy(central, 46)
    centrals.push(central)
    localOffset += local.length + contents.length
  }
  const central = Buffer.concat(centrals)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(entries.length, 8)
  end.writeUInt16LE(entries.length, 10)
  end.writeUInt32LE(central.length, 12)
  end.writeUInt32LE(localOffset, 16)
  return Buffer.concat([...locals, central, end])
}

describe('optional Enhanced mode component manager', () => {
  it('builds fail-closed notarization and Authenticode publisher checks', () => {
    expect(
      platformTrustCommands('darwin', '/private/app-server', {
        policy: 'macos',
        teamIdentifier: '2DC432GLL2',
        requireNotarization: true,
      }),
    ).toEqual([
      {
        file: '/usr/bin/codesign',
        args: ['--verify', '--strict', '--verbose=4', '/private/app-server'],
      },
      { file: '/usr/bin/codesign', args: ['-dv', '--verbose=4', '/private/app-server'] },
      {
        file: '/usr/sbin/spctl',
        args: ['--assess', '--type', 'execute', '--verbose=4', '/private/app-server'],
      },
    ])
    const windows = platformTrustCommands('win32', 'C:\\private\\app-server.exe', {
      policy: 'windows',
      publisher: 'OpenAI OpCo, LLC',
      publisherThumbprint: '8B0ADFB840E141DAD3044D2B5AC819873DDE3590',
    })
    expect(windows).toHaveLength(7)
    expect(windows[0]?.file).toBe('powershell.exe')
    expect(windows[0]?.args.at(-1)).toContain('$null -eq $s.SignerCertificate')
    expect(windows[1]?.args.at(-1)).toContain("Status -eq 'HashMismatch'")
    expect(windows[2]?.args.at(-1)).toContain("Status -eq 'NotTrusted'")
    expect(windows[3]?.args.at(-1)).toContain("Status -eq 'UnknownError'")
    expect(windows[4]?.args.at(-1)).toContain("Status -ne 'Valid'")
    expect(windows[5]?.args.at(-1)).toContain(
      "Thumbprint -ne '8B0ADFB840E141DAD3044D2B5AC819873DDE3590'",
    )
    expect(windows[6]?.args.at(-1)).toContain("Subject -notlike '*OpenAI OpCo, LLC*'")
    expect(() =>
      platformTrustCommands('linux', '/tmp/app-server', {
        policy: 'macos',
        teamIdentifier: '2DC432GLL2',
        requireNotarization: true,
      }),
    ).toThrowError('enhanced_mode_platform_trust_failed')
    expect(
      platformTrustCommands('darwin', '/private/app-server', {
        policy: 'macos',
        teamIdentifier: '2DC432GLL2',
        requireNotarization: false,
      }),
    ).toEqual([
      {
        file: '/usr/bin/codesign',
        args: ['--verify', '--strict', '--verbose=4', '/private/app-server'],
      },
      { file: '/usr/bin/codesign', args: ['-dv', '--verbose=4', '/private/app-server'] },
    ])
  })
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
    expect(
      parsed.component.assets.map((asset) => [
        asset.platform,
        asset.arch,
        asset.bytes,
        asset.sha256,
      ]),
    ).toEqual([
      [
        'darwin',
        'arm64',
        92_265_540,
        '552ec417bd679532a531175e705979b9908e575b54ca267f461620b77441a2ad',
      ],
      [
        'darwin',
        'x64',
        99_846_280,
        '3124de77fcb1f2eed35e959453ebd49cc4e786cf9236414948f02d58026728f8',
      ],
      [
        'win32',
        'x64',
        110_054_928,
        'c8908d687cf7caa3074921479726db32f96a295372c3544f1e96919a7254951f',
      ],
    ])
    for (const asset of parsed.component.assets) {
      expect(asset.layout.files.filter((file) => file.install).map((file) => file.path)).toEqual([
        asset.layout.entrypoint,
        asset.platform === 'win32' ? 'bin/codex-code-mode-host.exe' : 'bin/codex-code-mode-host',
        'codex-package.json',
      ])
    }
  })

  it('parses only the exact pinned release metadata and supported darwin-arm64 asset', async () => {
    const { manifest } = await createFixtureArchive()
    const parsed = parseCodexComponentManifest(manifest)
    expect(parsed.component.version).toBe('0.147.0')
    expect(parsed.component.assets.map((asset) => `${asset.platform}-${asset.arch}`)).toEqual([
      'darwin-arm64',
      'darwin-x64',
      'win32-x64',
    ])
    for (const [platform, arch] of [
      ['darwin', 'arm64'],
      ['darwin', 'x64'],
      ['win32', 'x64'],
    ] as const) {
      const manager = new EnhancedModeComponentManager({
        cacheRoot: join(await temporaryRoot(), 'cache'),
        manifest: parsed,
        platform,
        arch,
        verifyPlatformTrust: async () => undefined,
      })
      await expect(manager.status()).resolves.toMatchObject({ state: 'missing', supported: true })
    }

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
          assets: manifest.component.assets.map((asset, index) =>
            index === 0 ? { ...asset, platform: 'win32' } : asset,
          ),
        },
      }),
    ).toThrowError('enhanced_mode_manifest_invalid')
    expect(() =>
      parseCodexComponentManifest({
        ...manifest,
        component: {
          ...manifest.component,
          assets: manifest.component.assets.map((asset, index) =>
            index === 0
              ? {
                  ...asset,
                  primaryUrl: 'https://example.com/codex.tar.gz',
                }
              : asset,
          ),
        },
      }),
    ).toThrowError('enhanced_mode_manifest_invalid')
  })

  it('rejects platform basename substitution and Windows canonical path collisions', async () => {
    const { manifest } = await createFixtureArchive()
    expect(() =>
      parseCodexComponentManifest({
        ...manifest,
        component: {
          ...manifest.component,
          assets: manifest.component.assets.map((asset, index) =>
            index === 0
              ? {
                  ...asset,
                  primaryUrl: asset.primaryUrl.replace('.tar.gz', '-renamed.tar.gz'),
                  fallbackUrl: asset.fallbackUrl.replace('.tar.gz', '-renamed.tar.gz'),
                }
              : asset,
          ),
        },
      }),
    ).toThrowError('enhanced_mode_manifest_invalid')
    const windows = manifest.component.assets[2]!
    expect(() =>
      parseCodexComponentManifest({
        ...manifest,
        component: {
          ...manifest.component,
          assets: manifest.component.assets.map((asset, index) =>
            index === 2
              ? {
                  ...windows,
                  archive: {
                    ...windows.archive,
                    maxExtractedBytes: windows.archive.maxExtractedBytes + 1,
                  },
                  layout: {
                    ...windows.layout,
                    files: [
                      ...windows.layout.files,
                      { ...windows.layout.files[0]!, path: 'BIN/CODEX', bytes: 1 },
                    ],
                  },
                }
              : asset,
          ),
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
    for (const type of ['Link', 'BlockDevice', 'FIFO', 'Socket']) {
      expect(() => validateCodexArchiveEntry(`bin/${type}`, type, 0, new Set())).toThrow(
        'enhanced_mode_archive_unsafe',
      )
    }
    validateCodexArchiveEntry('bin/codex', 'File', 1, seen)
    expect(() => validateCodexArchiveEntry('bin/codex', 'File', 1, seen)).toThrow(
      'enhanced_mode_archive_unsafe',
    )
    expect(() => validateCodexArchiveEntry('CON.txt', 'File', 1, new Set(), 'win32')).toThrow(
      'enhanced_mode_archive_unsafe',
    )
    expect(() => validateCodexArchiveEntry('bin/agent. ', 'File', 1, new Set(), 'win32')).toThrow(
      'enhanced_mode_archive_unsafe',
    )
  })

  it('rejects non-regular Unix ZIP entry metadata', () => {
    for (const mode of [0o120777, 0o060600, 0o020600, 0o010600, 0o140600, 0]) {
      const archive = storedZip([{ name: 'bin/codex', contents: Buffer.from('x'), unixMode: mode }])
      const central = archive.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]))
      expect(() =>
        validateZipEntryMode(
          archive.readUInt16LE(central + 4),
          archive.readUInt32LE(central + 38),
          false,
        ),
      ).toThrowError('enhanced_mode_archive_unsafe')
    }
  })

  it('streams, verifies, atomically installs, probes, and resolves the pinned component', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const archiveBytes = await readFile(archive)
    const fetchImplementation = vi.fn(async () => responseFor(archiveBytes))
    const probeVersion = vi.fn(async () => 'codex-app-server 0.147.0')
    const verifyPlatformTrust = vi.fn(async () => undefined)
    const phases = vi.fn()
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation,
      probeVersion,
      verifyPlatformTrust,
      onPhase: phases,
    })

    await expect(manager.status()).resolves.toMatchObject({ state: 'missing', supported: true })
    const installed = await manager.install()
    expect(installed.executablePath).toBe(
      join(cacheRoot, '0.147.0', 'darwin-arm64', 'bin', 'codex'),
    )
    expect(fetchImplementation).toHaveBeenCalledTimes(1)
    expect(probeVersion).toHaveBeenCalledWith(installed.executablePath, expect.anything())
    expect(verifyPlatformTrust).toHaveBeenCalledWith(
      [installed.executablePath],
      expect.objectContaining({ policy: 'macos' }),
      expect.anything(),
    )
    await expect(manager.status()).resolves.toMatchObject({ state: 'ready', supported: true })
    await expect(manager.resolveExecutable()).resolves.toBe(installed.executablePath)
    expect((await lstat(installed.executablePath)).isSymbolicLink()).toBe(false)
    expect(phases.mock.calls.map(([event]) => event)).toEqual([
      { phase: 'download', outcome: 'started' },
      { phase: 'download', outcome: 'succeeded' },
      { phase: 'digest', outcome: 'started' },
      { phase: 'digest', outcome: 'succeeded' },
      { phase: 'signature', outcome: 'started' },
      { phase: 'signature', outcome: 'succeeded' },
      { phase: 'promote', outcome: 'started' },
      { phase: 'promote', outcome: 'succeeded' },
    ])
  })

  it('coalesces concurrent installs and garbage-collects only old version directories', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const archiveBytes = await readFile(archive)
    const fetchImplementation = vi.fn(async () => responseFor(archiveBytes))
    const phases = vi.fn()
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation,
      probeVersion: async () => 'codex-app-server 0.147.0',
      verifyPlatformTrust: async () => undefined,
      onPhase: phases,
    })
    await mkdir(join(cacheRoot, '0.146.0', 'darwin-arm64'), { recursive: true })
    await writeFile(join(cacheRoot, 'unowned-file'), 'keep')

    const [first, second] = await Promise.all([manager.install(), manager.install()])
    expect(first).toEqual(second)
    expect(fetchImplementation).toHaveBeenCalledTimes(1)
    expect(
      phases.mock.calls.map(([event]) => event).filter(({ phase }) => phase === 'update'),
    ).toEqual([
      { phase: 'update', outcome: 'started' },
      { phase: 'update', outcome: 'succeeded' },
    ])
    await expect(lstat(join(cacheRoot, '0.146.0'))).rejects.toMatchObject({ code: 'ENOENT' })
    await expect(readFile(join(cacheRoot, 'unowned-file'), 'utf8')).resolves.toBe('keep')
  })

  it('reports a failed update only when an older compatible cache version exists', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { manifest } = await createFixtureArchive()
    await mkdir(join(cacheRoot, '0.146.0', 'darwin-arm64'), { recursive: true })
    const phases = vi.fn((_event: ComponentPhaseEvent) => {
      throw new Error('sink unavailable')
    })
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: vi.fn(async () => {
        throw new Error('offline')
      }),
      probeVersion: async () => 'codex-app-server 0.147.0',
      verifyPlatformTrust: async () => undefined,
      onPhase: phases,
    })
    await expect(manager.install()).rejects.toMatchObject({ code: 'enhanced_mode_download_failed' })
    expect(
      phases.mock.calls.map(([event]) => event).filter(({ phase }) => phase === 'update'),
    ).toEqual([
      { phase: 'update', outcome: 'started' },
      { phase: 'update', outcome: 'failed' },
    ])
  })

  it('falls back from the pinned WisWork mirror to the same pinned official bytes', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const archiveBytes = await readFile(archive)
    const fetchImplementation = vi.fn(async (input: string | URL | Request) => {
      const url = String(input)
      if (url.startsWith('https://downloads.wiswork.com/')) throw new Error('mirror unavailable')
      return responseFor(archiveBytes)
    })
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: fetchImplementation as unknown as typeof fetch,
      probeVersion: async () => 'codex-app-server 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    await expect(manager.install()).resolves.toMatchObject({ platform: 'darwin', arch: 'arm64' })
    expect(fetchImplementation).toHaveBeenCalledTimes(2)
    expect(String(fetchImplementation.mock.calls[0]![0])).toMatch(
      /^https:\/\/downloads\.wiswork\.com\//,
    )
    expect(String(fetchImplementation.mock.calls[1]![0])).toMatch(
      /^https:\/\/github\.com\/openai\/codex\/releases\//,
    )
  })

  it('never follows an unapproved redirect and rejects substituted bytes from both sources', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { manifest } = await createFixtureArchive()
    const calls: string[] = []
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: (async (input: string | URL | Request) => {
        const url = String(input)
        calls.push(url)
        if (url.startsWith('https://downloads.wiswork.com/')) {
          return new Response(null, {
            status: 302,
            headers: { location: 'https://evil.example/component' },
          })
        }
        return responseFor(Buffer.from('substituted'))
      }) as typeof fetch,
      probeVersion: async () => 'codex-app-server 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    await expect(manager.install()).rejects.toMatchObject({
      code: 'enhanced_mode_integrity_failed',
    })
    expect(calls).toHaveLength(2)
    expect(calls).not.toContain('https://evil.example/component')
    await expect(manager.status()).resolves.toMatchObject({ state: 'missing' })
  })

  it('fails before download when app-private capacity is insufficient and leaves no partials', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { manifest } = await createFixtureArchive()
    const fetchImplementation = vi.fn()
    const checkAvailableCapacity = vi.fn(async () => {
      throw new Error('insufficient fixture capacity')
    })
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: fetchImplementation as unknown as typeof fetch,
      checkAvailableCapacity,
      probeVersion: async () => 'codex-app-server 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    await expect(manager.install()).rejects.toMatchObject({ code: 'enhanced_mode_capacity_failed' })
    expect(fetchImplementation).not.toHaveBeenCalled()
    const asset = manifest.component.assets[0]!
    expect(checkAvailableCapacity).toHaveBeenCalledWith(
      cacheRoot,
      asset.bytes + asset.archive.maxExtractedBytes + 64 * 1024 * 1024,
    )
    const entries = await import('node:fs/promises').then(({ readdir }) => readdir(cacheRoot))
    expect(entries.filter((name) => name.includes('.part') || name.includes('.staging'))).toEqual(
      [],
    )
  })

  it('extracts a bounded ZIP fixture and retains only the declared app-server entrypoint', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const executable = Buffer.from('#!/bin/sh\necho codex-app-server 0.147.0\n')
    const metadata = Buffer.from('{}')
    const archive = storedZip([
      { name: 'bin/', contents: Buffer.alloc(0), directory: true },
      { name: 'bin/codex', contents: executable },
      { name: 'codex-package.json', contents: metadata },
    ])
    const { manifest } = await createFixtureArchive()
    const asset = manifest.component.assets[0]!
    const zipAsset = {
      ...asset,
      bytes: archive.length,
      sha256: createHash('sha256').update(archive).digest('hex'),
      archive: { format: 'zip' as const, maxExtractedBytes: executable.length + metadata.length },
      layout: {
        ...asset.layout,
        files: [
          {
            path: 'bin/codex',
            bytes: executable.length,
            sha256: createHash('sha256').update(executable).digest('hex'),
            mode: 'executable' as const,
            install: true,
          },
          {
            path: 'codex-package.json',
            bytes: metadata.length,
            sha256: createHash('sha256').update(metadata).digest('hex'),
            mode: 'data' as const,
            install: false,
          },
        ],
      },
    }
    const zipManifest = {
      ...manifest,
      component: {
        ...manifest.component,
        assets: manifest.component.assets.map((item, index) => (index === 0 ? zipAsset : item)),
      },
    }
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest: zipManifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(archive),
      probeVersion: async () => 'codex-app-server 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    const installed = await manager.install()
    await expect(readFile(installed.executablePath, 'utf8')).resolves.toContain('codex-app-server')
    await expect(
      lstat(join(installed.executablePath, '..', '..', 'codex-package.json')),
    ).rejects.toMatchObject({ code: 'ENOENT' })
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
      probeVersion: async () => 'codex-app-server 0.147.0',
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
      probeVersion: async () => 'codex-app-server 0.147.0',
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
      platform: 'linux',
      arch: 'x64',
      fetchImplementation: async () => responseFor(archiveBytes),
      probeVersion: async () => 'codex-app-server 0.147.0',
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
      probeVersion: async () => 'codex-app-server 0.147.0',
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
      probeVersion: async () => 'codex-app-server 0.147.0',
      verifyPlatformTrust: async () => undefined,
    })
    const installed = await manager.install()
    const binDirectory = join(cacheRoot, '0.147.0', 'darwin-arm64', 'bin')

    await writeFile(join(binDirectory, 'unexpected-helper'), 'not allowlisted')
    await expect(manager.status()).resolves.toMatchObject({ state: 'invalid' })
    await rm(join(binDirectory, 'unexpected-helper'))

    if (process.platform !== 'win32') {
      await chmod(installed.executablePath, 0o755)
      await expect(manager.resolveExecutable()).rejects.toMatchObject({
        code: 'enhanced_mode_integrity_failed',
      })
    }
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
        assets: fixture.manifest.component.assets.map((current, index) =>
          index === 0
            ? {
                ...asset,
                archive: {
                  ...asset.archive,
                  maxExtractedBytes: asset.archive.maxExtractedBytes + 5,
                },
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
              }
            : current,
        ),
      },
    }
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest: unsafeManifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(bytes),
      probeVersion: async () => 'codex-app-server 0.147.0',
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
      probeVersion: async () => 'codex-app-server 0.147.0',
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
      return 'codex-app-server 0.147.0'
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
        return 'codex-app-server 0.147.0'
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
    const probeVersion = vi.fn(async () => 'codex-app-server 0.147.0')
    const phases = vi.fn()
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
      onPhase: phases,
    })

    await expect(manager.install()).rejects.toMatchObject({
      code: 'enhanced_mode_platform_trust_failed',
    })
    expect(probeVersion).not.toHaveBeenCalled()
    expect(phases.mock.calls.map(([event]) => event).at(-1)).toEqual({
      phase: 'signature',
      outcome: 'failed',
    })
    expect(phases.mock.calls.flat()).not.toContainEqual(
      expect.objectContaining({ phase: 'promote' }),
    )
    await expect(manager.status()).resolves.toMatchObject({ state: 'missing' })
  })

  it('preserves bounded macOS trust phase diagnostics without exposing command output', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const bytes = await readFile(archive)
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'darwin',
      arch: 'arm64',
      fetchImplementation: async () => responseFor(bytes),
      probeVersion: async () => 'codex-app-server 0.147.0',
      verifyPlatformTrust: async () => {
        throw new EnhancedModeComponentError('enhanced_mode_macos_notarization_failed')
      },
    })

    await expect(manager.install()).rejects.toMatchObject({
      code: 'enhanced_mode_macos_notarization_failed',
      message: 'enhanced_mode_macos_notarization_failed',
    })
    await expect(manager.status()).resolves.toMatchObject({ state: 'missing' })
  })

  it('preserves bounded Windows trust phase diagnostics without exposing command output', async () => {
    const cacheRoot = join(await temporaryRoot(), 'cache')
    const { archive, manifest } = await createFixtureArchive()
    const bytes = await readFile(archive)
    const manager = new EnhancedModeComponentManager({
      cacheRoot,
      manifest,
      platform: 'win32',
      arch: 'x64',
      fetchImplementation: async () => responseFor(bytes),
      probeVersion: async () => 'codex-app-server 0.147.0',
      verifyPlatformTrust: async () => {
        throw new EnhancedModeComponentError('enhanced_mode_windows_thumbprint_mismatch')
      },
    })

    await expect(manager.install()).rejects.toMatchObject({
      code: 'enhanced_mode_windows_thumbprint_mismatch',
      message: 'enhanced_mode_windows_thumbprint_mismatch',
    })
    await expect(manager.status()).resolves.toMatchObject({ state: 'missing' })
  })
})
