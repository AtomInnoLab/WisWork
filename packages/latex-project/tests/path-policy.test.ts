import { mkdtemp, mkdir, readFile, rename, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { ProjectPathPolicy } from '../src/path-policy.js'

describe('ProjectPathPolicy', () => {
  let sandbox: string
  let root: string
  let outside: string

  beforeEach(async () => {
    sandbox = await mkdtemp(join(tmpdir(), 'latex-path-policy-'))
    root = join(sandbox, 'project')
    outside = join(sandbox, 'outside.txt')
    await mkdir(join(root, 'chapters'), { recursive: true })
    await writeFile(join(root, 'chapters', 'one.tex'), 'hello')
    await writeFile(outside, 'secret')
  })

  afterEach(async () => {
    await rm(sandbox, { recursive: true, force: true })
  })

  it('normalizes safe relative paths while rejecting absolute, traversal, and NUL paths', async () => {
    const policy = await ProjectPathPolicy.open(root)
    expect(policy.normalize('chapters/./one.tex')).toBe('chapters/one.tex')
    await expect(policy.resolveExisting('/etc/passwd')).rejects.toThrow(/relative/i)
    await expect(policy.resolveExisting(String.raw`\etc\passwd`)).rejects.toThrow(/relative/i)
    await expect(policy.resolveExisting('../outside.txt')).rejects.toThrow(/traversal/i)
    await expect(policy.resolveExisting('chapters/\0one.tex')).rejects.toThrow(/NUL/i)
  })

  it('rejects symlinks in either a directory segment or the file itself', async () => {
    const policy = await ProjectPathPolicy.open(root)
    await symlink(sandbox, join(root, 'linked-dir'), 'dir')
    await symlink(outside, join(root, 'linked-file.tex'), 'file')
    await expect(policy.resolveExisting('linked-dir/outside.txt')).rejects.toThrow(/symbolic link/i)
    await expect(policy.resolveExisting('linked-file.tex')).rejects.toThrow(/symbolic link/i)
  })

  it('rejects missing files and prevents writes through symlinks', async () => {
    const policy = await ProjectPathPolicy.open(root)
    await symlink(outside, join(root, 'linked-file.tex'), 'file')
    await expect(policy.resolveExisting('missing.tex')).rejects.toThrow(/does not exist/i)
    await expect(policy.resolveForWrite('linked-file.tex')).rejects.toThrow(/symbolic link/i)
    expect(await readFile(outside, 'utf8')).toBe('secret')
  })

  it('detects a regular-file leaf replaced after its handle is opened', async () => {
    const original = join(root, 'chapters', 'one.tex')
    const replacement = join(root, 'chapters', 'replacement.tex')
    await writeFile(replacement, 'replacement')
    const policy = await ProjectPathPolicy.open(root, {
      afterReadOpen: async () => {
        await rename(replacement, original)
      },
    })
    await expect(policy.openTextFile('chapters/one.tex')).rejects.toThrow(
      /changed during validation/i,
    )
  })

  it('detects a parent directory replaced before a prepared write', async () => {
    const original = join(root, 'chapters')
    const moved = join(root, 'moved-chapters')
    const policy = await ProjectPathPolicy.open(root)
    const target = await policy.prepareWrite('chapters/new.tex')
    await rename(original, moved)
    await mkdir(original)
    await expect(policy.validateWriteTarget(target, 'before-rename')).rejects.toThrow(
      /parent directory changed/i,
    )
  })
})
