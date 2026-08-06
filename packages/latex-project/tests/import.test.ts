import { access, mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { importLatexProject } from '../src/import.js'
import { writeZipFixture } from './fixtures/zip.js'

describe('safe ZIP import', () => {
  const roots: string[] = []
  async function sandbox() {
    const root = await mkdtemp(join(tmpdir(), 'latex-import-'))
    roots.push(root)
    return root
  }
  afterEach(async () =>
    Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true }))),
  )

  it('extracts regular files only after the whole archive succeeds', async () => {
    const root = await sandbox()
    const zip = join(root, 'paper.zip')
    const target = join(root, 'paper')
    await writeZipFixture(zip, [
      { name: 'chapters/' },
      { name: 'main.tex', data: '\\documentclass{article}' },
      { name: 'chapters/a.tex', data: 'A' },
    ])
    const result = await importLatexProject(zip, target)
    expect(result.entryCount).toBe(3)
    expect(await readFile(join(target, 'chapters/a.tex'), 'utf8')).toBe('A')
  })

  it.each([
    ['traversal', '../escape.tex'],
    ['absolute', '/escape.tex'],
    ['drive absolute', 'C:\\escape.tex'],
    ['backslash rooted', '\\escape.tex'],
    ['NUL', 'bad\0.tex'],
    ['duplicate normalized', 'a/./b.tex'],
  ])('rejects %s names and leaves no target', async (_label, badName) => {
    const root = await sandbox()
    const zip = join(root, 'bad.zip')
    const target = join(root, 'target')
    const entries =
      badName === 'a/./b.tex'
        ? [
            { name: 'a/b.tex', data: 'one' },
            { name: badName, data: 'two' },
          ]
        : [{ name: badName, data: 'bad' }]
    await writeZipFixture(zip, entries)
    await expect(importLatexProject(zip, target)).rejects.toThrow()
    await expect(access(target)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it.each([
    [0o120777, 'symbolic link'],
    [0o060644, 'device'],
    [0o010644, 'non-regular'],
  ])('rejects Unix mode %o (%s)', async (unixMode) => {
    const root = await sandbox()
    const zip = join(root, 'special.zip')
    const target = join(root, 'target')
    await writeZipFixture(zip, [{ name: 'bad.tex', data: 'payload', unixMode }])
    await expect(importLatexProject(zip, target)).rejects.toThrow(/regular|type/i)
    await expect(access(target)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it.each([
    [
      'entry count',
      { maxEntries: 1 },
      [
        { name: 'a.tex', data: 'a' },
        { name: 'b.tex', data: 'b' },
      ],
    ],
    ['depth', { maxDirectoryDepth: 1 }, [{ name: 'a/b/c.tex', data: 'x' }]],
    ['single file', { maxFileBytes: 1 }, [{ name: 'a.tex', data: 'xx' }]],
    [
      'total size',
      { maxTotalBytes: 2 },
      [
        { name: 'a.tex', data: 'xx' },
        { name: 'b.tex', data: 'x' },
      ],
    ],
    ['compression ratio', { maxCompressionRatio: 0.5 }, [{ name: 'a.tex', data: 'x' }]],
  ])('enforces the %s limit without a partial target', async (_label, limits, entries) => {
    const root = await sandbox()
    const zip = join(root, 'limited.zip')
    const target = join(root, 'target')
    await writeZipFixture(zip, entries)
    await expect(importLatexProject(zip, target, limits)).rejects.toThrow(
      /limit|ratio|entries|depth/i,
    )
    await expect(access(target)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('rejects an oversized archive and an existing target without overwriting it', async () => {
    const root = await sandbox()
    const zip = join(root, 'paper.zip')
    const target = join(root, 'target')
    await writeZipFixture(zip, [{ name: 'main.tex', data: 'new' }])
    await mkdir(target)
    await writeFile(join(target, 'main.tex'), 'keep')
    await expect(importLatexProject(zip, target, { maxArchiveBytes: 1 })).rejects.toThrow(
      /archive/i,
    )
    await expect(importLatexProject(zip, target)).rejects.toThrow(/exists/i)
    expect(await readFile(join(target, 'main.tex'), 'utf8')).toBe('keep')
  })
})
