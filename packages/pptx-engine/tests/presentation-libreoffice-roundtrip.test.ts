import { execFileSync, spawnSync } from 'node:child_process'
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { openPptx } from '../src/index'
import { compilePresentationDeck } from '../src/presentation-compiler'
import { benchmarkDeck } from './fixtures/presentation-benchmark'

const sofficeAvailable = spawnSync('soffice', ['--version'], { timeout: 5_000 }).status === 0
if (process.env.WISWORK_REQUIRE_LIBREOFFICE === '1' && !sofficeAvailable)
  throw new Error('LibreOffice is required for the PPTX round-trip gate')

it.skipIf(!sofficeAvailable)(
  'reopens the eight-page editable benchmark through LibreOffice without losing native content',
  async () => {
    const directory = mkdtempSync(join(tmpdir(), 'wiswork-ppt-roundtrip-'))
    const inputDirectory = join(directory, 'input')
    const outputDirectory = join(directory, 'output')
    mkdirSync(inputDirectory)
    mkdirSync(outputDirectory)
    try {
      const { bytes } = await compilePresentationDeck(benchmarkDeck())
      const input = join(inputDirectory, 'benchmark.pptx')
      writeFileSync(input, bytes)
      execFileSync(
        'soffice',
        [
          `-env:UserInstallation=file://${join(directory, 'profile')}`,
          '--headless',
          '--convert-to',
          'pptx:Impress MS PowerPoint 2007 XML',
          '--outdir',
          outputDirectory,
          input,
        ],
        { timeout: 60_000, stdio: 'pipe' },
      )
      const reopened = await JSZip.loadAsync(readFileSync(join(outputDirectory, 'benchmark.pptx')))
      const slides = Object.keys(reopened.files)
        .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
        .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]))
      expect(slides).toHaveLength(8)
      const xml = await Promise.all(slides.map((name) => reopened.file(name)!.async('string')))
      for (const [index, title] of benchmarkDeck()
        .slides.map((slide) => slide.title)
        .entries())
        expect(xml[index]).toContain(title)
      expect(
        xml.some((page) =>
          page.replace(/<[^>]*>/g, '').includes('示例内容：原生文本可在 PowerPoint 中编辑。'),
        ),
      ).toBe(true)
      expect(xml.some((page) => page.includes('<a:tbl>'))).toBe(true)
      expect(xml.some((page) => page.includes('<p:pic>'))).toBe(true)
      expect(
        Object.keys(reopened.files).some((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name)),
      ).toBe(true)
      const original = benchmarkDeck()
      const parsed = (await openPptx(readFileSync(join(outputDirectory, 'benchmark.pptx')))).deck
      expect(parsed.slides).toHaveLength(original.slides.length)
      for (const [index, slide] of parsed.slides.entries()) {
        expect(slide.background).toEqual({ type: 'solid', color: `#${original.style.background}` })
        const title = slide.elements.find(
          (element) =>
            element.type === 'shape' &&
            element.text?.paragraphs.some((paragraph) =>
              paragraph.runs.some((run) => run.text === original.slides[index]!.title),
            ),
        )
        expect(title?.type).toBe('shape')
        if (title?.type !== 'shape') continue
        const run = title.text?.paragraphs[0]?.runs[0]
        expect(run?.fontFamily).toBe(original.style.fontFace)
        expect(run?.fontSize).toBe(32)
        expect(run?.color?.toUpperCase()).toBe(`#${original.style.textColor}`)
      }
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  },
  75_000,
)
