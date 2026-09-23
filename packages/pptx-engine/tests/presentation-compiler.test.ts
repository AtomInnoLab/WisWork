import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { openPptx } from '../src/index'
import { parsePresentationDeck, inspectPresentationGeometry } from '../src/presentation'
import { compilePresentationDeck } from '../src/presentation-compiler'
import { benchmarkDeck } from './fixtures/presentation-benchmark'

describe('presentation contract and compiler', () => {
  it('compiles eight Chinese slides to native editable objects and preserves attribution', async () => {
    const { bytes, report } = await compilePresentationDeck(benchmarkDeck())
    const opened = await openPptx(bytes)
    expect(opened.deck.slides).toHaveLength(8)
    expect(report.slideCount).toBe(8)
    expect(report.checks.render).toBe('not_run')
    expect(report.checks.sources).toBe('not_verified')
    const zip = await JSZip.loadAsync(bytes)
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).toContain('科研汇报')
    expect(await zip.file('ppt/notesSlides/notesSlide1.xml')!.async('string')).toContain('研究报告')
    expect(await zip.file('ppt/slides/slide6.xml')!.async('string')).toContain('<a:tbl>')
    expect(await zip.file('ppt/charts/chart1.xml')!.async('string')).toContain('120')
    expect(Object.keys(zip.files).some((name) => name.startsWith('ppt/media/image'))).toBe(true)
    expect(report.geometry).toEqual([])
    expect(opened.deck.slides[2]!.elements.some((el) => el.type === 'picture')).toBe(true)
    expect(opened.deck.slides[5]!.elements.some((el) => el.type === 'table')).toBe(true)
    expect(opened.deck.slides[6]!.elements.some((el) => el.type === 'chart')).toBe(true)
  })

  it('rejects malformed or unsafe input before compilation', () => {
    const mutations: Array<(d: any) => void> = [
      (d) => {
        d.version = 2
      },
      (d) => {
        d.extra = true
      },
      (d) => {
        d.slides[0].elements[0].x = NaN
      },
      (d) => {
        d.slides[0].elements[0].text = 'x'.repeat(12001)
      },
      (d) => {
        d.slides[0].elements[0].script = 'bad'
      },
      (d) => {
        d.slides[0].claimIds = ['missing']
      },
      (d) => {
        d.slides[1].id = d.slides[0].id
      },
      (d) => {
        d.assets[0].base64 = 'https://example.com/picture.png'
      },
      (d) => {
        d.slides[2].elements[1].assetId = 'missing'
      },
      (d) => {
        d.slides[6].elements[1].series[0].values = [1]
      },
    ]
    for (const mutate of mutations) {
      const deck = benchmarkDeck()
      mutate(deck)
      expect(() => parsePresentationDeck(deck)).toThrow(/presentation_invalid/)
    }
  })

  it('rejects mismatched image metadata and supports cover without external access', async () => {
    const deck = benchmarkDeck()
    deck.assets[0]!.width = 100
    await expect(compilePresentationDeck(deck)).rejects.toThrow(/image_dimensions_or_data/)
    deck.assets[0]!.width = 1
    const image = deck.slides[2]!.elements[1]!
    if (image.kind === 'image') image.fit = 'cover'
    const output = await compilePresentationDeck(deck)
    expect((await openPptx(output.bytes)).deck.slides).toHaveLength(8)
  })

  it('reports content collisions but permits background and explicitly layered elements', async () => {
    const deck = benchmarkDeck()
    deck.slides[0]!.elements.push({
      kind: 'text',
      id: 'overlap',
      x: 1,
      y: 1,
      w: 4,
      h: 1,
      text: '重叠',
    })
    expect(inspectPresentationGeometry(parsePresentationDeck(deck))).toEqual([
      expect.objectContaining({ kind: 'overlap', slideId: 'slide-1' }),
    ])
    expect((await compilePresentationDeck(deck)).report.checks.geometry).toBe('warning')
    deck.slides[0]!.elements.at(-1)!.allowOverlap = true
    expect(inspectPresentationGeometry(parsePresentationDeck(deck))).toEqual([])
    deck.slides[0]!.elements[0]!.x = 13
    await expect(compilePresentationDeck(deck)).rejects.toThrow(/presentation_geometry/)
  })
})
