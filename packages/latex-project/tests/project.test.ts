import { mkdtemp, mkdir, readFile, readdir, rename, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { atomicWriteFile } from '../src/atomic-write.js'
import { createLatexProject, openLatexProject, ProjectPathPolicy } from '../src/index.js'

describe('LaTeX directory projects', () => {
  const sandboxes: string[] = []

  afterEach(async () => {
    await Promise.all(sandboxes.splice(0).map((path) => rm(path, { recursive: true, force: true })))
  })

  async function sandbox() {
    const path = await mkdtemp(join(tmpdir(), 'latex-project-'))
    sandboxes.push(path)
    return path
  }

  it('atomically creates a basic article project and opens main.tex', async () => {
    const parent = await sandbox()
    const root = join(parent, 'paper')
    const project = await createLatexProject(root)
    expect(project.mainFile).toBe('main.tex')
    expect(await readFile(join(root, 'main.tex'), 'utf8')).toContain('\\documentclass{article}')
    expect(await project.listTextFiles()).toEqual(['main.tex'])
  })

  it('exclusively creates the target directory without overwriting existing content', async () => {
    const parent = await sandbox()
    const root = join(parent, 'existing')
    await mkdir(root)
    await writeFile(join(root, 'main.tex'), 'keep me')
    await expect(createLatexProject(root)).rejects.toThrow(/already exists/i)
    expect(await readFile(join(root, 'main.tex'), 'utf8')).toBe('keep me')
  })

  it('lists and reads UTF-8 text, enforces size limits, and rejects invalid UTF-8', async () => {
    const root = await sandbox()
    await mkdir(join(root, 'chapters'))
    await writeFile(join(root, 'chapters', 'a.tex'), '你好')
    await writeFile(join(root, 'large.tex'), '12345')
    await writeFile(join(root, 'bad.tex'), Buffer.from([0xc3, 0x28]))
    const project = await openLatexProject(root, { maxTextBytes: 4 })
    expect(await project.listTextFiles()).toEqual(['bad.tex', 'chapters/a.tex', 'large.tex'])
    await expect(project.readText('chapters/a.tex')).rejects.toThrow(/size limit/i)
    await expect(project.readText('large.tex')).rejects.toThrow(/size limit/i)
    await expect(project.readText('bad.tex')).rejects.toThrow(/UTF-8/i)
  })

  it('saves atomically, returns the SHA-256 digest, and never modifies outside files', async () => {
    const parent = await sandbox()
    const root = join(parent, 'project')
    await mkdir(root)
    const outside = join(parent, 'outside.tex')
    await writeFile(outside, 'secret')
    await symlink(outside, join(root, 'escape.tex'), 'file')
    const project = await openLatexProject(root)
    const saved = await project.saveText('new.tex', 'hello')
    expect(saved.sha256).toBe('2cf24dba5fb0a30e26e83b2ac5b9e29e1b161e5c1fa7425e73043362938b9824')
    expect(saved.path).toBe('new.tex')
    await expect(project.saveText('escape.tex', 'owned')).rejects.toThrow(/symbolic link/i)
    await expect(project.saveText('../outside.tex', 'owned')).rejects.toThrow(/traversal/i)
    expect(await readFile(outside, 'utf8')).toBe('secret')
  })

  it('limits handle reads to max + 1 bytes when a file grows after stat', async () => {
    const root = await sandbox()
    await writeFile(join(root, 'growing.tex'), '1234')
    const project = await openLatexProject(root, {
      maxTextBytes: 4,
      pathHooks: {
        afterReadStat: async () => writeFile(join(root, 'growing.tex'), '1234567890'),
      },
    })
    await expect(project.readText('growing.tex')).rejects.toThrow(/size limit/i)
  })

  it('rejects all access through an old project object after the real root is replaced', async () => {
    const parent = await sandbox()
    const root = join(parent, 'project')
    const moved = join(parent, 'moved-project')
    await mkdir(root)
    await writeFile(join(root, 'original.tex'), 'original')
    const project = await openLatexProject(root)
    await rename(root, moved)
    await mkdir(root)
    await writeFile(join(root, 'original.tex'), 'replacement')
    await expect(project.listTextFiles()).rejects.toThrow(/project root changed/i)
    await expect(project.readText('original.tex')).rejects.toThrow(/project root changed/i)
    await expect(project.saveText('new.tex', 'write')).rejects.toThrow(/project root changed/i)
    expect(await readFile(join(root, 'original.tex'), 'utf8')).toBe('replacement')
  })

  it('syncs the parent directory after rename and cleans temp files on failure', async () => {
    const root = await sandbox()
    const target = join(root, 'saved.tex')
    let syncCalls = 0
    await atomicWriteFile(target, Buffer.from('saved'), {
      syncDirectory: async (directory) => {
        expect(directory).toBe(root)
        syncCalls += 1
      },
    })
    expect(syncCalls).toBe(1)

    await expect(
      atomicWriteFile(join(root, 'failed.tex'), Buffer.from('nope'), {
        validateBeforeRename: async () => {
          throw new Error('validation failed')
        },
      }),
    ).rejects.toThrow('validation failed')
    expect((await readdir(root)).sort()).toEqual(['saved.tex'])
  })

  it('reports committed state when the renamed target identity is replaced', async () => {
    const root = await sandbox()
    const targetPath = join(root, 'target.tex')
    const replacementPath = join(root, 'replacement.tex')
    await writeFile(replacementPath, 'attacker')
    const policy = await ProjectPathPolicy.open(root)
    const target = await policy.prepareWrite('target.tex')
    let caught: unknown
    try {
      await atomicWriteFile(targetPath, Buffer.from('committed'), {
        validateAfterRename: async (tempIdentity) => {
          await rename(replacementPath, targetPath)
          await policy.validateWriteTarget(target, 'after-rename', tempIdentity)
        },
      })
    } catch (error) {
      caught = error
    }
    expect(caught).toMatchObject({
      code: 'LATEX_ATOMIC_WRITE_COMMITTED_VALIDATION_FAILED',
      committed: true,
    })
    expect(await readFile(targetPath, 'utf8')).toBe('attacker')
  })

  it('still syncs the parent directory when post-rename validation reports a race', async () => {
    const root = await sandbox()
    const target = join(root, 'raced.tex')
    let syncCalls = 0
    await expect(
      atomicWriteFile(target, Buffer.from('committed'), {
        validateAfterRename: async () => {
          throw new Error('post-rename race')
        },
        syncDirectory: async () => {
          syncCalls += 1
        },
      }),
    ).rejects.toThrow('post-rename race')
    expect(syncCalls).toBe(1)
    expect(await readFile(target, 'utf8')).toBe('committed')
  })
})
