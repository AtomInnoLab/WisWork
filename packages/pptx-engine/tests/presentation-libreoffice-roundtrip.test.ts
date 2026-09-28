import { execFileSync, spawnSync } from 'node:child_process'
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { openPptx } from '../src/index'
import { compilePresentationDeck } from '../src/presentation-compiler'
import { benchmarkDeck } from './fixtures/presentation-benchmark'

const sofficeAvailable = spawnSync('soffice', ['--version'], { timeout: 5_000 }).status === 0
const popplerAvailable = spawnSync('pdftoppm', ['-v'], { timeout: 5_000 }).status === 0
if (process.env.WISWORK_REQUIRE_LIBREOFFICE === '1' && !sofficeAvailable)
  throw new Error('LibreOffice is required for the PPTX round-trip gate')
if (process.env.WISWORK_REQUIRE_LIBREOFFICE === '1' && !popplerAvailable)
  throw new Error('Poppler is required for the PPTX visual round-trip gate')

it.skipIf(!sofficeAvailable)(
  'reopens the eight-page editable benchmark through LibreOffice without losing native content',
  async () => {
    const directory = mkdtempSync(join(tmpdir(), 'wiswork-ppt-roundtrip-'))
    const inputDirectory = join(directory, 'input')
    const outputDirectory = join(directory, 'output')
    mkdirSync(inputDirectory)
    mkdirSync(outputDirectory)
    try {
      const original = benchmarkDeck()
      const image = original.slides[2]!.elements[1]
      if (image?.kind !== 'image') throw new Error('benchmark_image_missing')
      image.altText = '合成像素示意图'
      const { bytes } = await compilePresentationDeck(original)
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
      const parsed = (await openPptx(readFileSync(join(outputDirectory, 'benchmark.pptx')))).deck
      expect(parsed.slides).toHaveLength(original.slides.length)
      const expectFrame = (
        actual: { offset: { x: number; y: number; cx: number; cy: number } },
        source: { x: number; y: number; w: number; h: number },
      ) => {
        for (const [field, expected] of [
          ['x', source.x],
          ['y', source.y],
          ['cx', source.w],
          ['cy', source.h],
        ] as const)
          expect(Math.abs(actual.offset[field] - expected * 914400)).toBeLessThanOrEqual(9144)
      }
      for (const [index, slide] of parsed.slides.entries()) {
        expect(slide.background).toEqual({ type: 'solid', color: `#${original.style.background}` })
        for (const source of original.slides[index]!.elements) {
          if (source.kind === 'text') {
            const native = slide.elements.find(
              (element) =>
                element.type === 'shape' &&
                element.text?.paragraphs
                  .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
                  .join('') === source.text,
            )
            expect(native?.type).toBe('shape')
            if (native?.type !== 'shape') continue
            expectFrame(native.transform, source)
            const run = native.text?.paragraphs[0]?.runs[0]
            expect(run?.fontFamily).toBe(original.style.fontFace)
            expect(run?.fontSize).toBe(source.fontSize ?? 20)
            expect(run?.color?.toUpperCase()).toBe(`#${source.color ?? original.style.textColor}`)
          } else if (source.kind === 'shape') {
            const native = slide.elements.find(
              (element) =>
                element.type === 'shape' &&
                element.presetGeometry === source.shape &&
                element.fill?.type === 'solid',
            )
            expect(native?.type).toBe('shape')
            if (native?.type !== 'shape') continue
            expectFrame(native.transform, source)
            expect(native.fill).toEqual({
              type: 'solid',
              color: `#${source.fill ?? original.style.accentColor}`,
            })
          } else if (source.kind === 'image') {
            const picture = slide.elements.find((element) => element.type === 'picture')
            expect(picture?.type).toBe('picture')
            if (picture?.type !== 'picture') continue
            expectFrame(picture.transform, source)
            expect(picture.descr).toBe('合成像素示意图')
            expect(picture.mediaRef).toMatch(/^ppt\/media\//)
            const media = reopened.file(picture.mediaRef)
            expect(media).not.toBeNull()
            expect((await media!.async('uint8array')).slice(0, 8)).toEqual(
              new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
            )
            const originalPixel = PNG.sync.read(Buffer.from(original.assets[0]!.base64, 'base64'))
            const reopenedPixel = PNG.sync.read(Buffer.from(await media!.async('uint8array')))
            expect(reopenedPixel.width).toBe(originalPixel.width)
            expect(reopenedPixel.height).toBe(originalPixel.height)
            expect(reopenedPixel.data).toEqual(originalPixel.data)
          } else if (source.kind === 'table') {
            const native = slide.elements.find((element) => element.type === 'table')
            expect(native?.type).toBe('table')
            if (native?.type !== 'table') continue
            expectFrame(native.transform, source)
            expect(
              native.rows.map((row) =>
                row.map(
                  (cell) =>
                    cell.text?.paragraphs
                      .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
                      .join('') ?? '',
                ),
              ),
            ).toEqual(source.rows)
          } else {
            const native = slide.elements.find((element) => element.type === 'chart')
            expect(native?.type).toBe('chart')
            if (native?.type !== 'chart') continue
            expectFrame(native.transform, source)
            expect(native.chart.categories).toEqual(source.categories)
            expect(native.chart.series.map((series) => series.values)).toEqual(
              source.series.map((series) => series.values),
            )
          }
        }
      }
      if (popplerAvailable) {
        const render = (pptxPath: string, folder: string) => {
          const pdfDirectory = join(directory, folder)
          mkdirSync(pdfDirectory)
          execFileSync(
            'soffice',
            [
              `-env:UserInstallation=file://${join(directory, `${folder}-profile`)}`,
              '--headless',
              '--convert-to',
              'pdf',
              '--outdir',
              pdfDirectory,
              pptxPath,
            ],
            { timeout: 60_000, stdio: 'pipe' },
          )
          execFileSync(
            'pdftoppm',
            [
              '-f',
              '1',
              '-l',
              '8',
              '-r',
              '72',
              '-png',
              join(pdfDirectory, 'benchmark.pdf'),
              join(pdfDirectory, 'page'),
            ],
            { timeout: 60_000, stdio: 'pipe' },
          )
          return Array.from({ length: 8 }, (_, index) =>
            PNG.sync.read(readFileSync(join(pdfDirectory, `page-${index + 1}.png`))),
          )
        }
        const before = render(input, 'before-render')
        const after = render(join(outputDirectory, 'benchmark.pptx'), 'after-render')
        for (let index = 0; index < 8; index++) {
          const originalPage = before[index]!
          const reopenedPage = after[index]!
          expect([reopenedPage.width, reopenedPage.height]).toEqual([
            originalPage.width,
            originalPage.height,
          ])
          let changed = 0
          let visible = 0
          for (let pixel = 0; pixel < originalPage.data.length; pixel += 4) {
            const rgb = [0, 1, 2] as const
            if (rgb.some((channel) => originalPage.data[pixel + channel]! < 245)) visible++
            if (
              rgb.some(
                (channel) =>
                  Math.abs(
                    originalPage.data[pixel + channel]! - reopenedPage.data[pixel + channel]!,
                  ) > 12,
              )
            )
              changed++
          }
          expect(visible).toBeGreaterThan(100)
          expect(changed / (originalPage.width * originalPage.height)).toBeLessThan(0.02)
        }
      }
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  },
  75_000,
)
