import { describe, expect, it } from 'vitest'
import { parseFileToText } from '../src/index'
import { buildImageBackedPdfFixture, buildPdfFixture, writeFixture } from './helpers/fixtures'
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

  it('separates a full-page image with visible text from an invisible text layer', async () => {
    const visible = await parseFileToText(
      writeFixture('image-visible.pdf', buildImageBackedPdfFixture('Visible evidence')),
    )
    expect(visible.ok).toBe(true)
    expect(visible.pagesWithFullPageImage).toEqual([1])
    expect(visible.pagesWithInvisibleTextLayer).toBeUndefined()

    const invisible = await parseFileToText(
      writeFixture('image-invisible.pdf', buildImageBackedPdfFixture('OCR evidence', true)),
    )
    expect(invisible.ok).toBe(true)
    expect(invisible.pagesWithFullPageImage).toEqual([1])
    expect(invisible.pagesWithInvisibleTextLayer).toEqual([1])
  })

  it('finds an unsampled scanned page in a long text PDF', async () => {
    const pages = Array.from({ length: 65 }, (_, index) => `Page ${index + 1} evidence`)
    const result = await parseFileToText(
      writeFixture('mixed-long.pdf', buildPdfFixture(pages, { imagePage: 2, invisibleText: true })),
    )
    expect(result.ok).toBe(true)
    expect(result.sections).toHaveLength(65)
    expect(result.pagesWithFullPageImage).toEqual([2])
    expect(result.pagesWithInvisibleTextLayer).toEqual([2])
  }, 20_000)

  it('does not label visible text as invisible when the mode is set after drawing', async () => {
    const result = await parseFileToText(
      writeFixture(
        'unused-invisible-mode.pdf',
        buildPdfFixture('Visible evidence', { imagePage: 1, unusedInvisibleMode: true }),
      ),
    )
    expect(result.pagesWithFullPageImage).toEqual([1])
    expect(result.pagesWithInvisibleTextLayer).toBeUndefined()
  })

  it('does not count mostly off-page images as full-page coverage', async () => {
    for (const offset of [300, 700]) {
      const result = await parseFileToText(
        writeFixture(
          `off-page-image-${offset}.pdf`,
          buildPdfFixture('Visible evidence', { imagePage: 1, imageOffsetX: offset }),
        ),
      )
      expect(result.pagesWithFullPageImage).toBeUndefined()
    }
  })

  it('fails gracefully on a corrupt pdf', async () => {
    const path = writeFixture('broken.pdf', Buffer.from('%PDF-1.4 garbage'))
    const result = await parseFileToText(path)
    expect(result.ok).toBe(false)
    expect(result.error).toBeTruthy()
  })

  it('reports an image-only research PDF as unreadable and accepts its text fallback', async () => {
    const materials = resolve(
      import.meta.dirname,
      '../../../docs/product/ppt-benchmark-materials/PPT-P0-11',
    )
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

  it('preserves the blank scanned page and source page numbers in a real NACA scan', async () => {
    const path = resolve(
      import.meta.dirname,
      '../../../docs/product/ppt-benchmark-materials/PPT-P0-11/naca-rm-l50b01-1950-real-scan.pdf',
    )
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.sections).toHaveLength(30)
    const blankPage = result.sections?.[1]
    expect(blankPage?.locator).toBe('第 2 页')
    expect(blankPage?.start).toBe(blankPage?.end)
    expect(result.text?.slice(result.sections![2]!.start, result.sections![2]!.end)).toContain(
      'NATIONAL ADVISORY COMMITTEE FOR AERONAUTICS',
    )
  }, 20_000)

  it('locates Apple and Toyota revenue and historical FX in the frozen P0-08 PDFs', async () => {
    const root = resolve(
      import.meta.dirname,
      '../../../docs/product/ppt-benchmark-materials/PPT-P0-08',
    )
    const apple = await parseFileToText(resolve(root, 'apple-fy2024-form10k.pdf'))
    expect(apple.ok).toBe(true)
    expect(apple.sections).toHaveLength(121)
    const appleRevenuePage = apple.sections![31]!
    expect(appleRevenuePage.locator).toBe('第 32 页')
    expect(apple.text!.slice(appleRevenuePage.start, appleRevenuePage.end)).toContain('391,035')
    expect(apple.pagesWithFullPageImage).toBeUndefined()

    const report = await parseFileToText(resolve(root, 'toyota-fy2024-form20f.pdf'))
    expect(report.ok).toBe(true)
    expect(report.sections).toHaveLength(281)
    const revenuePage = report.sections![166]!
    expect(revenuePage.locator).toBe('第 167 页')
    const revenueText = report.text!.slice(revenuePage.start, revenuePage.end)
    expect(revenueText).toContain('45,095,325')
    expect(revenueText).toContain('Financial services')

    const summary = await parseFileToText(resolve(root, 'toyota-fy2024-financial-summary.pdf'))
    expect(summary.ok).toBe(true)
    expect(summary.sections).toHaveLength(29)
    const fxPage = summary.sections![27]!
    expect(fxPage.locator).toBe('第 28 页')
    const fxText = summary.text!.slice(fxPage.start, fxPage.end)
    expect(fxText).toContain('FY2024')
    expect(fxText).toContain('Yen to US Dollar Rate')
  }, 20_000)

  it('identifies full-page raster images in the real NACA scan without certifying OCR text', async () => {
    const path = resolve(
      import.meta.dirname,
      '../../../docs/product/ppt-benchmark-materials/PPT-P0-11/naca-rm-l50b01-1950-real-scan.pdf',
    )
    const result = await parseFileToText(path)
    expect(result.ok).toBe(true)
    expect(result.pagesWithFullPageImage).toEqual(
      Array.from({ length: 30 }, (_, index) => index + 1),
    )
  }, 20_000)
})
