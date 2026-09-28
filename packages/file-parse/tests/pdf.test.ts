import { describe, expect, it } from 'vitest'
import { parseFileToText } from '../src/index'
import { buildPdfFixture, writeFixture } from './helpers/fixtures'
import { resolve } from 'node:path'

describe('parseFileToText: pdf', () => {
  it('extracts page text via pdfjs', async () => {
    const path = writeFixture('doc.pdf', buildPdfFixture('Hello PDF parsing'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.kind).toBe('text')
    expect(result.text).toContain('Hello PDF parsing')
  })
  it('keeps exact UTF-16 offsets for each PDF page', async () => {
    const path = writeFixture('two-pages.pdf', buildPdfFixture(['Page one', 'Second page']))
    const result = await parseFileToText(path)
    expect(result.text).toBe('Page one\n\nSecond page')
    expect(result.sections).toEqual([
      { locator: '第 1 页', start: 0, end: 8 },
      { locator: '第 2 页', start: 10, end: 21 },
    ])
  })

  it('fails gracefully on a corrupt pdf', async () => {
    const path = writeFixture('broken.pdf', Buffer.from('%PDF-1.4 garbage'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })

  it('reports an image-only research PDF as unreadable and accepts its text fallback', async () => {
    const materials = resolve(import.meta.dirname, '../../../docs/product/ppt-benchmark-materials/PPT-P0-11')
    const imageOnly = await parseFileToText(resolve(materials, 'deardorff-2020-image-only.pdf'))
    expect(imageOnly).toMatchObject({ ok: false, kind: 'text', error: 'pdf_no_extractable_text' })

    const source = await parseFileToText(resolve(materials, 'deardorff-2020-article.pdf'))
    expect(source.ok).toBe(true)
    expect(source.sections).toHaveLength(11)
    expect(source.text).toContain('Assessing the impact of introductory')

    const fallback = await parseFileToText(resolve(materials, 'deardorff-2020-assistive-text.txt'))
    expect(fallback.ok).toBe(true)
    expect(fallback.text).toContain('Assessing the impact of introductory')
  })
})
