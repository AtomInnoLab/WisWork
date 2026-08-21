import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseSkillArchive, SKILL_PACKAGE_LIMITS } from '../src/skills/shared/skill-package.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import { SkillRegistry } from '../src/skills/shared/skill-registry.js'
import {
  SkillPackageWorkerRuntime,
  type PackageWorkerLike,
} from '../src/skills/shared/skill-package-runtime.js'

const manifest = '---\nname: writer\ndescription: Helps edit prose\n---\nBe concise.'

async function archive(
  entries: Array<{ path: string; value: string | Uint8Array; options?: JSZip.JSZipFileOptions }>,
) {
  const zip = new JSZip()
  for (const entry of entries) zip.file(entry.path, entry.value, entry.options)
  return new Uint8Array(await zip.generateAsync({ type: 'arraybuffer', platform: 'UNIX' }))
}

describe('bounded skill archive parser', () => {
  it('accepts exactly one root SKILL.md plus allowlisted text and image assets', async () => {
    const result = await parseSkillArchive(
      await archive([
        { path: 'SKILL.md', value: manifest },
        { path: 'references/style.txt', value: 'Plain language.' },
        { path: 'assets/example.PNG', value: Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]) },
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
    expect(() => registry.installPackage(parsed)).toThrow('skill_already_installed')
    registry.uninstall('writer')
    expect(vfs.list('/home/skills/writer')).toEqual([])
    registry.installPackage(parsed)
    registry.clear()
    expect(vfs.list('/home/skills')).toEqual([])
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
