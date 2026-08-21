import JSZip from 'jszip'
import { deflateSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { parseSkillArchive, SKILL_PACKAGE_LIMITS } from '../src/skills/shared/skill-package.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import { SkillRegistry } from '../src/skills/shared/skill-registry.js'
import {
  SkillPackageWorkerRuntime,
  type PackageWorkerLike,
} from '../src/skills/shared/skill-package-runtime.js'

const manifest = '---\nname: writer\ndescription: Helps edit prose\n---\nBe concise.'
const png = (() => {
  const table = Uint32Array.from({ length: 256 }, (_, value) => {
    let crc = value
    for (let bit = 0; bit < 8; bit++) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1
    return crc >>> 0
  })
  const chunk = (type: string, data: Uint8Array) => {
    const value = new Uint8Array(12 + data.length)
    const view = new DataView(value.buffer)
    view.setUint32(0, data.length)
    value.set(
      [...type].map((letter) => letter.charCodeAt(0)),
      4,
    )
    value.set(data, 8)
    let crc = 0xffffffff
    for (const byte of value.subarray(4, 8 + data.length))
      crc = table[(crc ^ byte) & 0xff] ^ (crc >>> 8)
    view.setUint32(8 + data.length, (crc ^ 0xffffffff) >>> 0)
    return value
  }
  const header = new Uint8Array(13)
  const headerView = new DataView(header.buffer)
  headerView.setUint32(0, 1)
  headerView.setUint32(4, 1)
  header.set([8, 2, 0, 0, 0], 8)
  const parts = [
    Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk('IHDR', header),
    chunk('IDAT', deflateSync(Uint8Array.from([0, 0, 0, 0]))),
    chunk('IEND', new Uint8Array()),
  ]
  const value = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0))
  let offset = 0
  for (const part of parts) {
    value.set(part, offset)
    offset += part.length
  }
  return value
})()
const jpeg = Uint8Array.from([
  0xff, 0xd8, 0xff, 0xc0, 0, 11, 8, 0, 1, 0, 1, 1, 1, 0x11, 0, 0xff, 0xda, 0, 8, 1, 1, 0, 0, 63, 0,
  0, 0xff, 0xd9,
])
const metadataOnlyWebp = Uint8Array.from([
  82, 73, 70, 70, 22, 0, 0, 0, 87, 69, 66, 80, 86, 80, 56, 88, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
  0, 0,
])
const webp = Uint8Array.from([
  82, 73, 70, 70, 36, 0, 0, 0, 87, 69, 66, 80, 86, 80, 56, 88, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
  0, 0, 86, 80, 56, 76, 6, 0, 0, 0, 47, 0, 0, 0, 0, 0,
])

async function archive(
  entries: Array<{ path: string; value: string | Uint8Array; options?: JSZip.JSZipFileOptions }>,
) {
  const zip = new JSZip()
  for (const entry of entries) zip.file(entry.path, entry.value, entry.options)
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer', platform: 'UNIX' }))
}

function findSignature(bytes: Uint8Array, signature: number): number {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  for (let offset = 0; offset + 4 <= bytes.length; offset++)
    if (view.getUint32(offset, true) === signature) return offset
  throw new Error('fixture_signature_missing')
}

describe('bounded skill archive parser', () => {
  it('accepts exactly one root SKILL.md plus allowlisted text and image assets', async () => {
    const result = await parseSkillArchive(
      await archive([
        { path: 'SKILL.md', value: manifest },
        { path: 'references/style.txt', value: 'Plain language.' },
        { path: 'assets/example.PNG', value: png },
      ]),
    )
    expect(result.skill.metadata.name).toBe('writer')
    expect(result.files.map((file) => file.path)).toEqual([
      'SKILL.md',
      'assets/example.PNG',
      'references/style.txt',
    ])
  })

  it.each([
    ['missing manifest', [{ path: 'notes.md', value: 'x' }]],
    ['nested manifest', [{ path: 'docs/SKILL.md', value: manifest }]],
    [
      'duplicate manifest by case',
      [
        { path: 'SKILL.md', value: manifest },
        { path: 'skill.md', value: manifest },
      ],
    ],
    [
      'traversal',
      [
        { path: 'SKILL.md', value: manifest },
        { path: '../escape.txt', value: 'x' },
      ],
    ],
    [
      'absolute path',
      [
        { path: 'SKILL.md', value: manifest },
        { path: '/escape.txt', value: 'x' },
      ],
    ],
    [
      'backslash path',
      [
        { path: 'SKILL.md', value: manifest },
        { path: 'a\\b.txt', value: 'x' },
      ],
    ],
    [
      'case collision',
      [
        { path: 'SKILL.md', value: manifest },
        { path: 'A.txt', value: 'x' },
        { path: 'a.TXT', value: 'y' },
      ],
    ],
    [
      'unicode collision',
      [
        { path: 'SKILL.md', value: manifest },
        { path: 'café.txt', value: 'x' },
        { path: 'café.txt', value: 'y' },
      ],
    ],
    [
      'executable extension',
      [
        { path: 'SKILL.md', value: manifest },
        { path: 'run.js', value: 'x' },
      ],
    ],
    [
      'unsupported image',
      [
        { path: 'SKILL.md', value: manifest },
        { path: 'asset.gif', value: 'GIF89a' },
      ],
    ],
  ])('rejects %s', async (_label, entries) => {
    await expect(parseSkillArchive(await archive(entries as never))).rejects.toThrow(
      'invalid_skill_package',
    )
  })

  it('rejects symlink and executable-mode entries', async () => {
    for (const unixPermissions of [0o120777, 0o100755]) {
      const bytes = await archive([
        { path: 'SKILL.md', value: manifest },
        { path: 'notes.txt', value: 'x', options: { unixPermissions } },
      ])
      await expect(parseSkillArchive(bytes)).rejects.toThrow('invalid_skill_package')
    }
  })

  it.each([
    'flag mismatch',
    'version mismatch',
    'ZIP64 sentinel',
    'local name mismatch',
    'overlapping range',
  ] as const)('rejects inconsistent raw ZIP metadata: %s', async (mode) => {
    const bytes = await archive([{ path: 'SKILL.md', value: manifest }])
    const view = new DataView(bytes.buffer)
    const central = findSignature(bytes, 0x02014b50)
    const local = findSignature(bytes, 0x04034b50)
    if (mode === 'flag mismatch')
      view.setUint16(local + 6, view.getUint16(local + 6, true) ^ 0x8, true)
    if (mode === 'version mismatch')
      view.setUint16(local + 4, view.getUint16(local + 4, true) + 1, true)
    if (mode === 'ZIP64 sentinel') view.setUint32(central + 24, 0xffffffff, true)
    if (mode === 'local name mismatch') bytes[local + 30] ^= 1
    if (mode === 'overlapping range') view.setUint32(central + 42, 1, true)
    await expect(parseSkillArchive(bytes)).rejects.toThrow('invalid_skill_package')
  })

  it('rejects entry count and compressed, per-file, and aggregate uncompressed limits', async () => {
    const many = [{ path: 'SKILL.md', value: manifest }]
    for (let index = 0; index < SKILL_PACKAGE_LIMITS.maxEntries; index++)
      many.push({ path: `r/${index}.txt`, value: '' })
    await expect(parseSkillArchive(await archive(many))).rejects.toThrow('skill_package_limit')
    await expect(
      parseSkillArchive(new Uint8Array(SKILL_PACKAGE_LIMITS.maxCompressedBytes + 1)),
    ).rejects.toThrow('skill_package_limit')
    await expect(
      parseSkillArchive(
        await archive([
          { path: 'SKILL.md', value: manifest },
          { path: 'large.txt', value: 'x'.repeat(SKILL_PACKAGE_LIMITS.maxFileBytes + 1) },
        ]),
      ),
    ).rejects.toThrow('skill_package_limit')
    const aggregate = [{ path: 'SKILL.md', value: manifest }]
    for (let index = 0; index < 9; index++)
      aggregate.push({
        path: `${index}.txt`,
        value: 'x'.repeat(Math.ceil(SKILL_PACKAGE_LIMITS.maxUncompressedBytes / 8)),
      })
    await expect(parseSkillArchive(await archive(aggregate))).rejects.toThrow('skill_package_limit')
  })

  it('rejects overlong entry paths', async () => {
    await expect(
      parseSkillArchive(
        await archive([
          { path: 'SKILL.md', value: manifest },
          { path: `${'a'.repeat(SKILL_PACKAGE_LIMITS.maxPathBytes)}.txt`, value: 'x' },
        ]),
      ),
    ).rejects.toThrow('invalid_skill_package')
  })

  it('rejects invalid UTF-8 in manifest or text assets', async () => {
    await expect(
      parseSkillArchive(await archive([{ path: 'SKILL.md', value: Uint8Array.from([0xff]) }])),
    ).rejects.toThrow('invalid_skill_package')
    await expect(
      parseSkillArchive(
        await archive([
          { path: 'SKILL.md', value: manifest },
          { path: 'bad.txt', value: Uint8Array.from([0xff]) },
        ]),
      ),
    ).rejects.toThrow('invalid_skill_package')
  })

  it('rejects truncated and excessive-dimension images', async () => {
    await expect(
      parseSkillArchive(
        await archive([
          { path: 'SKILL.md', value: manifest },
          { path: 'bad.png', value: png.slice(0, 20) },
        ]),
      ),
    ).rejects.toThrow('invalid_skill_package')
    const huge = png.slice()
    new DataView(huge.buffer).setUint32(16, 100_000)
    new DataView(huge.buffer).setUint32(20, 100_000)
    await expect(
      parseSkillArchive(
        await archive([
          { path: 'SKILL.md', value: manifest },
          { path: 'huge.png', value: huge },
        ]),
      ),
    ).rejects.toThrow('invalid_skill_package')
  })

  it.each([
    ['jpeg', jpeg],
    ['webp', webp],
  ])('accepts a structurally bounded %s image', async (extension, value) => {
    await expect(
      parseSkillArchive(
        await archive([
          { path: 'SKILL.md', value: manifest },
          { path: `asset.${extension}`, value },
        ]),
      ),
    ).resolves.toMatchObject({ skill: { metadata: { name: 'writer' } } })
  })

  it('rejects inconsistent WebP RIFF/chunk lengths', async () => {
    const malformed = webp.slice()
    new DataView(malformed.buffer).setUint32(4, 999, true)
    await expect(
      parseSkillArchive(
        await archive([
          { path: 'SKILL.md', value: manifest },
          { path: 'bad.webp', value: malformed },
        ]),
      ),
    ).rejects.toThrow('invalid_skill_package')
  })

  it('rejects metadata-only WebP and malformed PNG/JPEG semantic structures', async () => {
    const badCrc = png.slice()
    badCrc[32] ^= 1
    const noSos = jpeg.slice(0, 17)
    noSos.set([0xff, 0xd9], noSos.length - 2)
    for (const [path, value] of [
      ['metadata.webp', metadataOnlyWebp],
      ['bad-crc.png', badCrc],
      ['no-sos.jpg', noSos],
    ] as const) {
      await expect(
        parseSkillArchive(
          await archive([
            { path: 'SKILL.md', value: manifest },
            { path, value },
          ]),
        ),
      ).rejects.toThrow('invalid_skill_package')
    }
  })

  it('honors cancellation without returning partial output', async () => {
    const controller = new AbortController()
    controller.abort()
    await expect(
      parseSkillArchive(await archive([{ path: 'SKILL.md', value: manifest }]), controller.signal),
    ).rejects.toThrow('upload_cancelled')
  })
})

describe('terminateable skill package worker runtime', () => {
  it.each(['cancel', 'timeout'] as const)('terminates its worker on %s', async (mode) => {
    let terminated = false
    const worker: PackageWorkerLike = {
      onmessage: null,
      onerror: null,
      postMessage() {},
      terminate() {
        terminated = true
      },
    }
    const runtime = new SkillPackageWorkerRuntime({ workerFactory: () => worker, timeoutMs: 1 })
    const controller = new AbortController()
    const pending = runtime.parse(new ArrayBuffer(1), controller.signal)
    if (mode === 'cancel') controller.abort()
    await expect(pending).rejects.toThrow(
      mode === 'cancel' ? 'upload_cancelled' : 'skill_package_timeout',
    )
    expect(terminated).toBe(true)
  })

  it('terminates active workers on clear and rejects late results', async () => {
    let terminated = false
    const worker: PackageWorkerLike = {
      onmessage: null,
      onerror: null,
      postMessage() {},
      terminate() {
        terminated = true
      },
    }
    const runtime = new SkillPackageWorkerRuntime({ workerFactory: () => worker })
    const pending = runtime.parse(new ArrayBuffer(1))
    runtime.cancelAll()
    await expect(pending).rejects.toThrow('upload_cancelled')
    expect(terminated).toBe(true)
  })

  it('terminates immediately when postMessage throws synchronously', async () => {
    let terminated = false
    const runtime = new SkillPackageWorkerRuntime({
      workerFactory: () => ({
        onmessage: null,
        onerror: null,
        postMessage() {
          throw new Error('clone failed')
        },
        terminate() {
          terminated = true
        },
      }),
    })
    await expect(runtime.parse(new ArrayBuffer(1))).rejects.toThrow('invalid_skill_package')
    expect(terminated).toBe(true)
  })
})

describe('skill package registry lifecycle', () => {
  it('installs atomically, rejects duplicates, uninstalls, and clears logout state', async () => {
    const vfs = new InMemoryVfs()
    const registry = new SkillRegistry(vfs)
    const parsed = await parseSkillArchive(
      await archive([
        { path: 'SKILL.md', value: manifest },
        { path: 'notes.txt', value: 'x' },
      ]),
    )
    registry.installPackage(parsed)
    expect(vfs.readText('/home/skills/writer/notes.txt')).toBe('x')
    expect(registry.prompt()).toContain('## writer\nBe concise.')
    expect(() => registry.installPackage(parsed)).toThrow('skill_already_installed')
    registry.uninstall('writer')
    expect(vfs.list('/home/skills/writer')).toEqual([])
    registry.installPackage(parsed)
    registry.clear()
    expect(vfs.list('/home/skills')).toEqual([])
  })

  it('preserves the exact bounded body in the next Agent context', () => {
    const vfs = new InMemoryVfs()
    const registry = new SkillRegistry(vfs)
    const body = 'First rule.\n\nSecond rule: use references/style.txt.'
    registry.install(`---\nname: exact\ndescription: Exact instructions\n---\n${body}`)
    expect(registry.prompt()).toBe(
      `exact: Exact instructions (/home/skills/exact/SKILL.md)\n## exact\n${body}`,
    )
  })

  it('rolls back registry and VFS state when mounting fails', async () => {
    const vfs = new InMemoryVfs({ maxFiles: 1 })
    const registry = new SkillRegistry(vfs)
    const parsed = await parseSkillArchive(
      await archive([
        { path: 'SKILL.md', value: manifest },
        { path: 'notes.txt', value: 'x' },
      ]),
    )
    expect(() => registry.installPackage(parsed)).toThrow('vfs_limit')
    expect(registry.list()).toEqual([])
    expect(vfs.list('/home/skills')).toEqual([])
  })
})
