import { mkdtemp, readFile, writeFile, symlink, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { createOfficeDesignDocuments } from '../src/main/office-design-document'

const directories: string[] = []
afterEach(async () => {
  await Promise.all(directories.splice(0).map((path) => rm(path, { recursive: true, force: true })))
})
async function setup() {
  const root = await mkdtemp(join(tmpdir(), 'wiswork-design-test-'))
  directories.push(root)
  const openFile = vi.fn(async (_path: string) => {})
  const documents = createOfficeDesignDocuments({ directory: root, openFile })
  const request = (
    body: unknown,
    sessionId = 'session-one',
    signal = new AbortController().signal,
  ) => documents({ sessionId, host: 'PowerPoint', body, signal })
  return { root, openFile, request }
}
const documentId = 'document_12345678'
const markdown = '# DESIGN.md\n\n## Visual direction\n\nOcean blue.'

describe('connected Office DESIGN.md files', () => {
  it('opens a PC-owned Markdown file and reads actual saved changes back', async () => {
    const { request, openFile } = await setup()
    const opened = await request({ action: 'open', documentId, markdown })
    const path = openFile.mock.calls[0]![0]
    expect(path.endsWith('/DESIGN.md')).toBe(true)
    expect(await readFile(path, 'utf8')).toBe(markdown)
    expect(opened.markdown).toBe(markdown)
    await writeFile(path, `${markdown}\nUpdated on PC.`, 'utf8')
    const saved = await request({ action: 'read', documentId })
    expect(saved.markdown).toContain('Updated on PC.')
    expect(saved.revision).not.toBe(opened.revision)
    await request({ action: 'open', documentId, markdown })
    expect(await readFile(path, 'utf8')).toBe(saved.markdown)
  })

  it('refuses a stale open instead of overwriting a file being edited', async () => {
    const { request, openFile } = await setup()
    await request({ action: 'open', documentId, markdown })
    await expect(
      request({ action: 'open', documentId, markdown: '# New contract' }),
    ).rejects.toThrow('design_document_conflict')
    expect(await readFile(openFile.mock.calls[0]![0], 'utf8')).toBe(markdown)
  })

  it('isolates files by paired session and rejects unregistered reads', async () => {
    const { request } = await setup()
    await request({ action: 'open', documentId, markdown })
    await expect(request({ action: 'read', documentId }, 'session-two')).rejects.toThrow(
      'design_document_missing',
    )
  })

  it('rejects paths, unexpected fields, oversized content and cancelled requests', async () => {
    const { request, openFile } = await setup()
    for (const body of [
      { action: 'open', documentId: '../outside', markdown },
      { action: 'open', documentId, markdown, path: '/tmp/outside' },
      { action: 'open', documentId, markdown: 'x'.repeat(100 * 1024) },
    ])
      await expect(request(body)).rejects.toThrow('design_document_invalid')
    await expect(
      request({ action: 'open', documentId, markdown }, 'session-one', AbortSignal.abort()),
    ).rejects.toThrow('cancelled')
    expect(openFile).not.toHaveBeenCalled()
  })

  it('does not read an external file substituted through a symlink', async () => {
    const { request, openFile, root } = await setup()
    await request({ action: 'open', documentId, markdown })
    const path = openFile.mock.calls[0]![0]
    const external = join(root, 'external.md')
    await writeFile(external, 'private content', 'utf8')
    await rm(path)
    await symlink(external, path)
    await expect(request({ action: 'read', documentId })).rejects.toThrow(
      'design_document_unavailable',
    )
  })
})
