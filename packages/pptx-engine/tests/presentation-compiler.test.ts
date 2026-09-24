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
    expect(report.assetWarnings).toEqual({ missingSource: 0, unknownLicense: 1, missingAltText: 1 })
    expect(await zip.file('ppt/slides/slide3.xml')!.async('string')).toContain(
      'Image description missing',
    )
    expect(opened.deck.slides[5]!.elements.some((el) => el.type === 'table')).toBe(true)
    expect(opened.deck.slides[6]!.elements.some((el) => el.type === 'chart')).toBe(true)
  })

  it('writes supplied alt text and license provenance to the PPTX', async () => {
    const deck = benchmarkDeck()
    deck.assets[0]!.license = 'licensed'
    const image = deck.slides[2]!.elements[1]!
    if (image.kind !== 'image') throw new Error('invalid fixture')
    image.altText = '显微镜下的细胞图像'
    const { bytes, report } = await compilePresentationDeck(deck)
    expect(report.assetWarnings).toEqual({ missingSource: 0, unknownLicense: 0, missingAltText: 0 })
    const zip = await JSZip.loadAsync(bytes)
    expect(await zip.file('ppt/slides/slide3.xml')!.async('string')).toContain('显微镜下的细胞图像')
    expect(await zip.file('ppt/notesSlides/notesSlide3.xml')!.async('string')).toContain(
      'license: licensed',
    )
  })

  it('preserves every recorded image source in slide notes', async () => {
    const deck = benchmarkDeck()
    deck.assets[0]!.source = 'https://example.com/first.png'
    deck.assets[0]!.sources = ['https://example.com/first.png', 'https://example.org/second.png']
    const { bytes } = await compilePresentationDeck(deck)
    const zip = await JSZip.loadAsync(bytes)
    const notes = await zip.file('ppt/notesSlides/notesSlide3.xml')!.async('string')
    expect(notes).toContain('https://example.com/first.png')
    expect(notes).toContain('https://example.org/second.png')
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
        d.assets[0].license = 'verified by AI'
      },
      (d) => {
        d.assets[0].sources = []
      },
      (d) => {
        d.assets[0].sources = ['x'.repeat(2001)]
      },
      (d) => {
        d.assets[0].sources = ['Different from primary source']
      },
      (d) => {
        d.slides[2].elements[1].altText = ''
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

  it.each([
    { w: 4, h: 2, crop: { l: 0, r: 0, t: 0.25, b: 0.25 } },
    { w: 2, h: 4, crop: { l: 0.25, r: 0.25, t: 0, b: 0 } },
  ])('preserves square-source aspect ratio when covering $w × $h', async ({ w, h, crop }) => {
    const deck = benchmarkDeck()
    const image = deck.slides[2]!.elements[1]!
    if (image.kind !== 'image') throw new Error('fixture_image_missing')
    Object.assign(image, { fit: 'cover', w, h })
    const { bytes } = await compilePresentationDeck(deck)
    const opened = await openPptx(bytes)
    const picture = opened.deck.slides[2]!.elements.find((el) => el.type === 'picture')!
    expect(picture.type === 'picture' && picture.srcRect).toEqual(crop)
    expect(picture.transform.offset).toMatchObject({ cx: w * 914400, cy: h * 914400 })
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
it('accepts compact attachment references but never compiles unresolved asset bytes', async () => {
  const deck: import('../src/presentation').PresentationDeck = benchmarkDeck()
  deck.assets = [{ id: deck.assets[0]!.id, attachmentId: 'a'.repeat(64) }]
  expect(parsePresentationDeck(deck)).toEqual(deck)
  await expect(compilePresentationDeck(deck)).rejects.toThrow(
    'presentation_invalid:unresolved_asset',
  )
  expect(() =>
    parsePresentationDeck({
      ...deck,
      assets: [{ ...deck.assets[0], source: 'https://untrusted.example' }],
    }),
  ).toThrow('schema')
  expect(() =>
    parsePresentationDeck({ ...deck, assets: [{ ...deck.assets[0], attachmentId: '../asset' }] }),
  ).toThrow('schema')
  expect(() =>
    parsePresentationDeck({ ...deck, assets: [deck.assets[0], deck.assets[0]] }),
  ).toThrow('duplicate_asset')
})
it('accepts more than 32 distinct image references in one editable page', () => {
  const deck: import('../src/presentation').PresentationDeck = benchmarkDeck()
  const images = Array.from({ length: 33 }, (_, index) => ({
    id: `image_${index}`,
    attachmentId: index.toString(16).padStart(64, '0'),
  }))
  deck.assets.push(...images)
  deck.slides[0]!.elements = images.map((asset, index) => ({
    kind: 'image' as const,
    id: `element_${index}`,
    assetId: asset.id,
    x: (index % 11) * 1.1,
    y: Math.floor(index / 11) * 1.1,
    w: 1,
    h: 1,
  }))
  expect(parsePresentationDeck(deck).assets).toHaveLength(34)
})

it('returns Office source slide IDs from the final OOXML in deck order on repeated compilation', async () => {
  const deck = benchmarkDeck()
  const first = await compilePresentationDeck(deck)
  const second = await compilePresentationDeck(deck)
  const xml = await (
    await JSZip.loadAsync(first.bytes)
  )
    .file('ppt/presentation.xml')!
    .async('string')
  const ids = [...xml.matchAll(/<p:sldId\s+id="(\d+)"/g)].map((match) => `${match[1]}#`)
  expect(ids).toHaveLength(8)
  expect(first.sourceSlideIds).toEqual(ids)
  expect(second.sourceSlideIds).toEqual(ids)
  expect(new Set(ids).size).toBe(8)
})
