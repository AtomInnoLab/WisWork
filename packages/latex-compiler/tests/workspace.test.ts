import { access, mkdtemp, mkdir, readFile, rm, symlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { createCompileWorkspace } from '../src/workspace.js'

describe('isolated compile workspace', () => {
  const roots: string[] = []
  afterEach(async () =>
    Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true }))),
  )

  async function setup() {
    const root = await mkdtemp(join(tmpdir(), 'latex-workspace-'))
    roots.push(root)
    const project = join(root, 'project')
    await mkdir(project)
    await writeFile(join(project, 'main.tex'), 'main')
    return { root, project }
  }

  it('copies regular files into a distinct isolated input and always cleans it', async () => {
    const { root, project } = await setup()
    await mkdir(join(project, 'chapters'))
    await writeFile(join(project, 'chapters', 'a.tex'), 'chapter')
    const workspace = await createCompileWorkspace(project, join(root, 'tmp'))
    expect(workspace.inputDirectory).not.toBe(project)
    expect(await readFile(join(workspace.inputDirectory, 'chapters/a.tex'), 'utf8')).toBe('chapter')
    await workspace.cleanup()
    await expect(access(workspace.root)).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('rejects links and leaves no isolated workspace', async () => {
    const { root, project } = await setup()
    await symlink('/tmp', join(project, 'escape'))
    await expect(createCompileWorkspace(project, join(root, 'tmp'))).rejects.toThrow(/link/i)
    await expect(access(join(root, 'tmp'))).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it('enforces file, total-byte, entry, and main-file path limits', async () => {
    const { root, project } = await setup()
    await writeFile(join(project, 'large.bin'), '12345')
    await expect(
      createCompileWorkspace(project, join(root, 'a'), { maxFileBytes: 4 }),
    ).rejects.toThrow(/limit/i)
    await expect(
      createCompileWorkspace(project, join(root, 'b'), { maxTotalBytes: 5 }),
    ).rejects.toThrow(/limit/i)
    await expect(
      createCompileWorkspace(project, join(root, 'c'), { maxEntries: 1 }),
    ).rejects.toThrow(/limit/i)
    await expect(
      createCompileWorkspace(project, join(root, 'd'), { mainFile: '../main.tex' }),
    ).rejects.toThrow(/path|traversal/i)
  })
})
