import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises'
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
    return { registry, session, contents: { id: 22, isDestroyed: () => false } }
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
})
