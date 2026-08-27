import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { exportPublishedPdf } from '../src/main/export-pdf.js'

describe('PDF export', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('leaves the filesystem unchanged when Save As is cancelled', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-pdf-export-'))
    roots.push(root)
    const source = join(root, 'compiled.pdf')
    await writeFile(source, '%PDF-compiled')
    const dialog = { showSaveDialog: vi.fn().mockResolvedValue({ canceled: true }) }

    await expect(exportPublishedPdf(dialog, source, 'main.pdf')).resolves.toEqual({
      state: 'cancelled',
    })
    expect(dialog.showSaveDialog).toHaveBeenCalledWith(
      expect.objectContaining({ defaultPath: 'main.pdf' }),
    )
  })

  it('atomically copies the published revision to the chosen path', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-pdf-export-'))
    roots.push(root)
    const source = join(root, 'compiled.pdf')
    const target = join(root, 'paper.pdf')
    await writeFile(source, '%PDF-compiled')
    const dialog = {
      showSaveDialog: vi.fn().mockResolvedValue({ canceled: false, filePath: target }),
    }

    await expect(exportPublishedPdf(dialog, source, 'main.pdf')).resolves.toEqual({
      state: 'written',
      path: target,
    })
    await expect(readFile(target, 'utf8')).resolves.toBe('%PDF-compiled')
  })
})
