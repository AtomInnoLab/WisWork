import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { ProjectSessionRegistry } from '../src/main/project-session.js'
import { latexQueryDirty, requestLatexClose } from '../src/main/latex-main.js'

describe('LaTeX dirty close guard', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  async function dirtySession() {
    const root = await mkdtemp(join(tmpdir(), 'latex-close-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before')
    const registry = new ProjectSessionRegistry({ watch: () => ({ close() {} }) })
    const session = await registry.attach(22, projectRoot)
    await session.readText('main.tex')
    session.updateBuffer('main.tex', 'local')
    return { projectRoot, registry, session, contents: { id: 22, isDestroyed: () => false } }
  }

  it('queries clean and dirty state from the sender-owned session', async () => {
    const { registry, session, contents } = await dirtySession()
    expect(await latexQueryDirty(contents, registry)).toBe(true)
    await session.discardAll()
    expect(await latexQueryDirty(contents, registry)).toBe(false)
  })

  it.each([
    ['cancel', 2, false, true],
    ['discard', 1, true, false],
  ] as const)(
    '%s close follows explicit user choice',
    async (_label, response, closes, remainsDirty) => {
      const { registry, contents } = await dirtySession()
      const showMessageBox = vi.fn().mockResolvedValue({ response })
      await expect(requestLatexClose(contents, null, registry, { showMessageBox })).resolves.toBe(
        closes,
      )
      expect(await latexQueryDirty(contents, registry)).toBe(remainsDirty)
    },
  )

  it('save choice persists all dirty buffers before allowing close', async () => {
    const { registry, contents, session } = await dirtySession()
    const showMessageBox = vi.fn().mockResolvedValue({ response: 0 })
    await expect(requestLatexClose(contents, null, registry, { showMessageBox })).resolves.toBe(
      true,
    )
    expect(session.isDirty()).toBe(false)
  })

  it('waits for in-flight autosave then Save persists latest v3', async () => {
    const { projectRoot, registry, contents, session } = await dirtySession()
    const original = session.project.saveText.bind(session.project)
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    vi.spyOn(session.project, 'saveText').mockImplementation(async (...args) => {
      const saved = await original(...args)
      if (args[1] === 'local') await gate
      return saved
    })
    const autosave = session.saveText('main.tex')
    await vi.waitFor(async () =>
      expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('local'),
    )
    session.updateBuffer('main.tex', 'v3')
    const showMessageBox = vi.fn().mockResolvedValue({ response: 0 })
    const closing = requestLatexClose(contents, null, registry, { showMessageBox })
    await Promise.resolve()
    expect(showMessageBox).not.toHaveBeenCalled()
    release()
    await autosave
    await expect(closing).resolves.toBe(true)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('v3')
  })

  it("settles autosave before Don't Save so discard has no late write", async () => {
    const { projectRoot, registry, contents, session } = await dirtySession()
    const original = session.project.saveText.bind(session.project)
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    vi.spyOn(session.project, 'saveText').mockImplementation(async (...args) => {
      const saved = await original(...args)
      await gate
      return saved
    })
    const autosave = session.saveText('main.tex')
    await vi.waitFor(async () =>
      expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('local'),
    )
    session.updateBuffer('main.tex', 'v3')
    const closing = requestLatexClose(contents, null, registry, {
      showMessageBox: vi.fn().mockResolvedValue({ response: 1 }),
    })
    release()
    await autosave
    await expect(closing).resolves.toBe(true)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('local')
    expect(session.isDirty()).toBe(false)
  })

  it('preserves a concurrent edit while close Discard is reading disk', async () => {
    const { registry, contents, session } = await dirtySession()
    const original = session.project.readText.bind(session.project)
    let markStarted!: () => void
    let release!: () => void
    const started = new Promise<void>((resolve) => (markStarted = resolve))
    const gate = new Promise<void>((resolve) => (release = resolve))
    vi.spyOn(session.project, 'readText').mockImplementation(async (...args) => {
      markStarted()
      await gate
      return original(...args)
    })
    const closing = requestLatexClose(contents, null, registry, {
      showMessageBox: vi.fn().mockResolvedValue({ response: 1 }),
    })
    await started
    session.updateBuffer('main.tex', 'v3')
    release()
    await expect(closing).resolves.toBe(false)
    expect(session.getBuffer('main.tex')).toMatchObject({ text: 'v3', dirty: true })
  })

  it('returns false when a new edit arrives during close Save', async () => {
    const { projectRoot, registry, contents, session } = await dirtySession()
    const original = session.project.saveText.bind(session.project)
    let release!: () => void
    const gate = new Promise<void>((resolve) => (release = resolve))
    vi.spyOn(session.project, 'saveText').mockImplementation(async (...args) => {
      const saved = await original(...args)
      await gate
      return saved
    })
    const closing = requestLatexClose(contents, null, registry, {
      showMessageBox: vi.fn().mockResolvedValue({ response: 0 }),
    })
    await vi.waitFor(async () =>
      expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('local'),
    )
    session.updateBuffer('main.tex', 'v3')
    release()
    await expect(closing).resolves.toBe(false)
    expect(session.getBuffer('main.tex')).toMatchObject({ text: 'v3', dirty: true })
  })
})
