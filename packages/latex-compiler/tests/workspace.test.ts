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

  it('applies replacement and new text overlays only to the isolated input', async () => {
    const { root, project } = await setup()
    await mkdir(join(project, 'chapters'))
    await writeFile(join(project, 'chapters/a.tex'), 'before')
    const workspace = await createCompileWorkspace(project, join(root, 'tmp'), {
      overlay: [
        { path: 'chapters/a.tex', text: 'after' },
        { path: 'chapters/new.tex', text: 'new' },
      ],
    })

    expect(await readFile(join(workspace.inputDirectory, 'chapters/a.tex'), 'utf8')).toBe('after')
    expect(await readFile(join(workspace.inputDirectory, 'chapters/new.tex'), 'utf8')).toBe('new')
    expect(await readFile(join(project, 'chapters/a.tex'), 'utf8')).toBe('before')
    await expect(access(join(project, 'chapters/new.tex'))).rejects.toMatchObject({ code: 'ENOENT' })
  })

  it.each([
    ['absolute', '/escape.tex', 'x'],
    ['drive absolute', 'C:\\escape.tex', 'x'],
    ['backslash rooted', '\\escape.tex', 'x'],
    ['traversal', '../escape.tex', 'x'],
    ['disallowed extension', 'image.png', 'x'],
    ['NUL text', 'main.tex', 'x\0y'],
    ['invalid UTF-8 text', 'main.tex', '\ud800'],
  ])('rejects invalid overlay %s', async (_label, path, text) => {
    const { root, project } = await setup()
    await expect(
      createCompileWorkspace(project, join(root, 'tmp'), { overlay: [{ path, text }] }),
    ).rejects.toMatchObject({ code: 'TECTONIC_WORKSPACE_INVALID' })
  })

  it('rejects duplicate normalized overlay paths and overlay size limits', async () => {
    const { root, project } = await setup()
    await expect(
      createCompileWorkspace(project, join(root, 'duplicate'), {
        overlay: [
          { path: 'a/./b.tex', text: 'one' },
          { path: 'a/b.tex', text: 'two' },
        ],
      }),
    ).rejects.toThrow(/duplicate/i)
    await expect(
      createCompileWorkspace(project, join(root, 'count'), {
        overlay: [
          { path: 'a.tex', text: 'a' },
          { path: 'b.tex', text: 'b' },
        ],
        maxOverlayFiles: 1,
      }),
    ).rejects.toThrow(/file count/i)
    await expect(
      createCompileWorkspace(project, join(root, 'file-bytes'), {
        overlay: [{ path: 'a.tex', text: '123' }],
        maxOverlayFileBytes: 2,
      }),
    ).rejects.toThrow(/file size/i)
    await expect(
      createCompileWorkspace(project, join(root, 'total-bytes'), {
        overlay: [
          { path: 'a.tex', text: '12' },
          { path: 'b.tex', text: '34' },
        ],
        maxOverlayTotalBytes: 3,
      }),
    ).rejects.toThrow(/total size/i)
  })

  it('rejects overlay targets below a linked directory', async () => {
    const { root, project } = await setup()
    const outside = join(root, 'outside')
    await mkdir(outside)
    await symlink(outside, join(project, 'linked'))
    await expect(
      createCompileWorkspace(project, join(root, 'tmp'), {
        overlay: [{ path: 'linked/escape.tex', text: 'owned' }],
      }),
    ).rejects.toThrow(/link/i)
    await expect(access(join(outside, 'escape.tex'))).rejects.toMatchObject({ code: 'ENOENT' })
  })
})
