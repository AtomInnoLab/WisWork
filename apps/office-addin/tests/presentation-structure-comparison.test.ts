import { expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { openPptx } from '@wiswork/pptx-engine'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { comparePresentationPageStructure } from '../src/skills/powerpoint/presentation-structure-comparison'
import { createPresentationQaSkill } from '../src/skills/powerpoint/presentation-qa'
import { presentationArtifactContent } from '../src/skills/powerpoint/presentation-page-delivery'
import { InMemoryVfs } from '../src/skills/shared/vfs'

it('detects changed native text in an exported host page package', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const hostPackage = await JSZip.loadAsync(bytes)
  const slide = await hostPackage.file('ppt/slides/slide1.xml')!.async('string')
  hostPackage.file('ppt/slides/slide1.xml', slide.replace('科研汇报', '替换标题'))
  const hostBase64 = await hostPackage.generateAsync({ type: 'base64' })
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    hostBase64,
  )
  expect(result.content).toMatchObject({ status: 'warning', changed: ['title'] })
})

it('reports changed and unverified native slide backgrounds', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const background = xml.match(/<p:bg>[\s\S]*?<\/p:bg>/)?.[0]
  expect(background).toContain('val="FFFFFF"')
  const inspection = {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes,
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: '' },
  }
  for (const [changed, expected] of [
    [
      xml.replace(background!, background!.replace('val="FFFFFF"', 'val="112233"')),
      { status: 'warning', backgroundChanged: true, backgroundUnchecked: false },
    ],
    [
      xml.replace(background!, ''),
      { status: 'incomplete', backgroundChanged: false, backgroundUnchecked: true },
    ],
  ] as const) {
    const host = await JSZip.loadAsync(bytes)
    host.file('ppt/slides/slide1.xml', changed)
    const result = await comparePresentationPageStructure(
      Buffer.from(bytes).toString('base64'),
      0,
      inspection,
      await host.generateAsync({ type: 'base64' }),
    )
    expect(result.content).toMatchObject(expected)
  }

  const inherited =
    '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></p:bgPr></p:bg>'
  for (const folder of ['slideLayouts', 'slideMasters'] as const) {
    const host = await JSZip.loadAsync(bytes)
    host.file('ppt/slides/slide1.xml', xml.replace(background!, ''))
    if (folder === 'slideMasters') {
      const layoutPath = Object.keys(host.files).find((part) =>
        /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(part),
      )!
      const layout = await host.file(layoutPath)!.async('string')
      host.file(layoutPath, layout.replace(/<p:bg>[\s\S]*?<\/p:bg>/, ''))
    }
    const path = Object.keys(host.files).find((part) =>
      new RegExp(`^ppt/${folder}/slide(?:Layout|Master)\\d+\\.xml$`).test(part),
    )
    expect(path).toBeTruthy()
    const parent = await host.file(path!)!.async('string')
    const withoutBackground = parent.replace(/<p:bg>[\s\S]*?<\/p:bg>/, '')
    host.file(
      path!,
      withoutBackground.replace(/<p:cSld\b[^>]*>/, (tag) => `${tag}${inherited}`),
    )
    const result = await comparePresentationPageStructure(
      Buffer.from(bytes).toString('base64'),
      0,
      inspection,
      await host.generateAsync({ type: 'base64' }),
    )
    expect(result.content).toMatchObject({
      status: 'warning',
      backgroundChanged: true,
      backgroundUnchecked: false,
    })
  }
  const themed = await JSZip.loadAsync(bytes)
  themed.file('ppt/slides/slide1.xml', xml.replace(background!, ''))
  const layoutPath = Object.keys(themed.files).find((part) =>
    /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(part),
  )!
  const layout = await themed.file(layoutPath)!.async('string')
  const themeBackground = '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
  themed.file(
    layoutPath,
    layout
      .replace(/<p:bg>[\s\S]*?<\/p:bg>/, '')
      .replace(/<p:cSld\b[^>]*>/, (tag) => `${tag}${themeBackground}`),
  )
  const themedResult = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    inspection,
    await themed.generateAsync({ type: 'base64' }),
  )
  expect(themedResult.content).toMatchObject({ status: 'incomplete', backgroundUnchecked: true })
})

it('detects native text font and size drift while text content remains unchanged', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const title = xml.match(/<p:sp>[^]*?<\/p:sp>/g)?.find((item) => item.includes('name="title"'))
  expect(title).toBeTruthy()
  const altered = title!
    .replace('sz="3200"', 'sz="2800"')
    .replace('typeface="Microsoft YaHei"', 'typeface="Arial"')
  zip.file('ppt/slides/slide1.xml', xml.replace(title!, altered))
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content).toMatchObject({
    status: 'warning',
    changed: [],
    textStyleChanged: ['title'],
  })
})

it('detects changed table cells and leaves chart data explicitly unchecked', async () => {
  for (const [pageIndex, original, replacement, nativeType] of [
    [5, '120', '999', 'Table'],
    [6, '120', '999', 'Chart'],
  ] as const) {
    const deck = benchmarkDeck()
    deck.slides = [deck.slides[pageIndex]!]
    const { bytes } = await compilePresentationDeck(deck)
    const source = (await openPptx(bytes)).deck.slides[0]!
    const shapes = source.elements.map((element, index) => ({
      id: String(index),
      name: element.name!,
      type: element.name === 'table' || element.name === 'chart' ? nativeType : 'TextBox',
      left: (element.transform.offset.x * 72) / 914400,
      top: (element.transform.offset.y * 72) / 914400,
      width: (element.transform.offset.cx * 72) / 914400,
      height: (element.transform.offset.cy * 72) / 914400,
    }))
    const hostPackage = await JSZip.loadAsync(bytes)
    const target = nativeType === 'Table' ? 'ppt/slides/slide1.xml' : 'ppt/charts/chart1.xml'
    const content = await hostPackage.file(target)!.async('string')
    const valueTag = nativeType === 'Table' ? 'a:t' : 'c:v'
    hostPackage.file(
      target,
      content.replace(
        `<${valueTag}>${original}</${valueTag}>`,
        `<${valueTag}>${replacement}</${valueTag}>`,
      ),
    )
    const result = await comparePresentationPageStructure(
      Buffer.from(bytes).toString('base64'),
      0,
      {
        slideId: 'host',
        slideWidth: 960,
        slideHeight: 540,
        shapes,
        shapesTruncated: false,
        overflows: [],
        overlaps: [],
        overlapsTruncated: false,
        screenshot: { mime: 'image/png', base64: '' },
      },
      await hostPackage.generateAsync({ type: 'base64' }),
    )
    expect(result.structureStatus).toBe('passed')
    if (nativeType === 'Table')
      expect(result).toMatchObject({
        status: 'warning',
        content: { status: 'warning', changed: ['table'] },
      })
    else
      expect(result).toMatchObject({
        status: 'warning',
        content: { status: 'warning', cacheChanged: ['chart'], unchecked: ['chart'] },
      })
  }
})

it('detects native table cell fill, border and font drift when cell text is unchanged', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[5]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.name === 'table' ? 'Table' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const table = slide.match(
    /<p:graphicFrame>[^]*?<a:tbl>[^]*?<\/a:tbl>[^]*?<\/p:graphicFrame>/,
  )?.[0]
  expect(table).toBeTruthy()
  for (const altered of [
    table!.replace(/(<a:rPr[^>]*\bsz=")[^"]+/, (_match, prefix: string) => `${prefix}9900`),
    table!.replace('</a:tcPr>', '<a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:tcPr>'),
    table!.replace(
      /(<a:lnL[^>]*>[^]*?<a:srgbClr val=")[^"]+/,
      (_match, prefix: string) => `${prefix}112233`,
    ),
  ]) {
    expect(altered).not.toBe(table)
    zip.file('ppt/slides/slide1.xml', slide.replace(table!, altered))
    const result = await comparePresentationPageStructure(
      Buffer.from(bytes).toString('base64'),
      0,
      {
        slideId: 'host',
        slideWidth: 960,
        slideHeight: 540,
        shapes,
        shapesTruncated: false,
        overflows: [],
        overlaps: [],
        overlapsTruncated: false,
        screenshot: { mime: 'image/png', base64: '' },
      },
      await zip.generateAsync({ type: 'base64' }),
    )
    expect(result.content).toMatchObject({
      status: 'warning',
      tableStyleChanged: ['table'],
      changed: [],
    })
  }
})

it('compares a selected chart cache from a multi-page source deck', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(bytes)).deck.slides[6]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.name === 'chart' ? 'Chart' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const one = benchmarkDeck()
  one.slides = [one.slides[6]!]
  const host = await compilePresentationDeck(one)
  const zip = await JSZip.loadAsync(host.bytes)
  const hostInspection = {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes,
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: '' },
  }
  const unchanged = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    6,
    hostInspection,
    Buffer.from(host.bytes).toString('base64'),
  )
  expect(unchanged.content).toMatchObject({
    status: 'incomplete',
    cacheChanged: [],
    unchecked: ['chart'],
  })
  const chartPath = Object.keys(zip.files).find((path) =>
    /^ppt\/charts\/chart\d+\.xml$/.test(path),
  )!
  const chart = await zip.file(chartPath)!.async('string')
  zip.file(chartPath, chart.replace('<c:v>120</c:v>', '<c:v>999</c:v>'))
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    6,
    hostInspection,
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content).toMatchObject({
    status: 'warning',
    cacheChanged: ['chart'],
    unchecked: ['chart'],
  })
})

it('detects chart category, series label and native plot type drift after host export', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.type === 'chart' ? 'Chart' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const original = Buffer.from(bytes).toString('base64')
  const inspect = {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes,
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: '' },
  }
  const sourceZip = await JSZip.loadAsync(bytes)
  const path = Object.keys(sourceZip.files).find((item) =>
    /^ppt\/charts\/chart\d+\.xml$/.test(item),
  )!
  const chart = await sourceZip.file(path)!.async('string')
  expect(chart).toContain('<c:barChart>')
  for (const [edited, cacheChanged, chartTypeChanged] of [
    [chart.replace('<c:v>甲</c:v>', '<c:v>丙</c:v>'), ['chart'], []],
    [chart.replace('<c:v>示例</c:v>', '<c:v>修改后</c:v>'), ['chart'], []],
    [chart.replace(/<c:ser>[\s\S]*?<\/c:ser>/, ''), ['chart'], []],
    [chart.replaceAll('c:barChart', 'c:lineChart'), [], ['chart']],
  ] as const) {
    expect(edited).not.toBe(chart)
    const zip = await JSZip.loadAsync(bytes)
    zip.file(path, edited)
    const result = await comparePresentationPageStructure(
      original,
      0,
      inspect,
      await zip.generateAsync({ type: 'base64' }),
    )
    expect(result.content).toMatchObject({
      status: 'warning',
      cacheChanged,
      chartTypeChanged,
      unchecked: ['chart'],
    })
  }
})

it('reports embedded chart workbook bytes changing while cached values stay the same', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.type === 'chart' ? 'Chart' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const workbookPath = Object.keys(zip.files).find((path) =>
    /^ppt\/embeddings\/[^/]+\.xlsx$/.test(path),
  )!
  const book = await JSZip.loadAsync(await zip.file(workbookPath)!.async('uint8array'))
  const sheet = await book.file('xl/worksheets/sheet1.xml')!.async('string')
  book.file(
    'xl/worksheets/sheet1.xml',
    sheet.replace('<c r="B2"><v>120</v></c>', '<c r="B2"><v>999</v></c>'),
  )
  zip.file(workbookPath, await book.generateAsync({ type: 'uint8array' }))
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content).toMatchObject({
    status: 'warning',
    cacheChanged: [],
    workbookBytesChanged: ['chart'],
    unchecked: ['chart'],
  })
})

it('checks embedded workbooks for all seventeen charts', async () => {
  const deck = benchmarkDeck()
  const chart = deck.slides[6]!.elements[1]!
  if (chart.kind !== 'chart') throw new Error('benchmark chart missing')
  deck.slides = [deck.slides[6]!]
  deck.slides[0]!.elements = [
    deck.slides[0]!.elements[0]!,
    ...Array.from({ length: 17 }, (_, index) => ({
      ...chart,
      id: `chart-${index + 1}`,
      x: 0.2 + (index % 6) * 2,
      y: 1.7 + Math.floor(index / 6) * 1.5,
      w: 1.8,
      h: 1.3,
    })),
  ]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.type === 'chart' ? 'Chart' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const workbookPaths = Object.keys(zip.files).filter((path) =>
    /^ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/.test(path),
  )
  expect(workbookPaths).toHaveLength(17)
  for (const workbookPath of workbookPaths) {
    const book = await JSZip.loadAsync(await zip.file(workbookPath)!.async('uint8array'))
    const sheet = await book.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<c r="B2"><v>120</v></c>')
    book.file(
      'xl/worksheets/sheet1.xml',
      sheet.replace('<c r="B2"><v>120</v></c>', '<c r="B2"><v>999</v></c>'),
    )
    zip.file(workbookPath, await book.generateAsync({ type: 'uint8array' }))
  }
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content.workbookBytesChanged).toEqual(
    Array.from({ length: 17 }, (_, index) => `chart-${index + 1}`),
  )
})

it('detects replaced embedded picture media in an exported page', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.type === 'picture' ? 'Image' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const mediaPath = Object.keys(zip.files).find((path) =>
    /^ppt\/media\/image[^/]+\.png$/.test(path),
  )!
  zip.file(
    mediaPath,
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII=',
      'base64',
    ),
  )
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content).toMatchObject({
    status: 'warning',
    mediaChanged: ['image'],
    mediaChecked: ['image'],
    mediaUnchecked: [],
    unchecked: ['image'],
  })
})

it('checks media on every image when a page contains more than sixteen pictures', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  deck.assets = Array.from({ length: 17 }, (_, index) => {
    const png = new PNG({ width: 1, height: 1 })
    png.data = Buffer.from([index + 1, 0, 0, 255])
    return {
      id: `pixel-${index + 1}`,
      mime: 'image/png' as const,
      width: 1,
      height: 1,
      base64: PNG.sync.write(png).toString('base64'),
      source: 'Synthetic fixture',
    }
  })
  deck.slides[0]!.elements = [
    deck.slides[0]!.elements[0]!,
    ...deck.assets.map((asset, index) => ({
      kind: 'image' as const,
      id: `image-${index + 1}`,
      assetId: asset.id,
      x: 0.4 + (index % 6) * 2,
      y: 1.8 + Math.floor(index / 6) * 1.6,
      w: 1,
      h: 1,
      fit: 'contain' as const,
    })),
  ]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.type === 'picture' ? 'Image' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const media = Object.keys(zip.files).filter((path) => /^ppt\/media\/image[^/]+\.png$/.test(path))
  expect(media).toHaveLength(17)
  zip.file(media[16]!, Buffer.from(deck.assets[0]!.base64, 'base64'))
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content.mediaChecked).toHaveLength(17)
  expect(result.content.mediaUnchecked).toEqual([])
  expect(result.content.mediaChanged).toEqual(['image-17'])
})

it('detects changed picture alternative text and crop independently of media bytes', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.type === 'picture' ? 'Image' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const picture = xml.match(/<p:pic>[^]*?<\/p:pic>/)?.[0]
  expect(picture).toBeTruthy()
  const changedPicture = picture!
    .replace(/(<p:cNvPr\b[^>]*\bdescr=")[^"]*(")/, '$1Changed description$2')
    .replace(/<p:blipFill>/, '<p:blipFill><a:srcRect l="10000"/>')
  expect(changedPicture).not.toBe(picture)
  zip.file('ppt/slides/slide1.xml', xml.replace(picture!, changedPicture))
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content).toMatchObject({
    status: 'warning',
    mediaChanged: [],
    mediaChecked: [],
    mediaUnchecked: ['image'],
    altTextChanged: ['image'],
    cropChanged: ['image'],
  })
})

it('detects native shape preset, fill and rotation drift after host export', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[3]!]
  const { bytes } = await compilePresentationDeck(deck)
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: 'GeometricShape',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shape = xml.match(/<p:sp>[^]*?<\/p:sp>/g)?.find((item) => item.includes('name="step"'))
  expect(shape).toBeTruthy()
  const altered = shape!
    .replace('prst="roundRect"', 'prst="rect"')
    .replace('val="2255AA"', 'val="AA5522"')
    .replace(/<a:xfrm(?=[ >])/, '<a:xfrm rot="5400000"')
  expect(altered).not.toBe(shape)
  zip.file('ppt/slides/slide1.xml', xml.replace(shape!, altered))
  const result = await comparePresentationPageStructure(
    Buffer.from(bytes).toString('base64'),
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    await zip.generateAsync({ type: 'base64' }),
  )
  expect(result.content).toMatchObject({ status: 'warning', appearanceChanged: ['step'] })
  expect(result.issues).toContainEqual({ name: 'step', kind: 'rotation_changed' })
})

it('compares a selected picture from a multi-page source deck', async () => {
  const full = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(full.bytes)).deck.slides[2]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: element.type === 'picture' ? 'Image' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const one = benchmarkDeck()
  one.slides = [one.slides[2]!]
  const host = await compilePresentationDeck(one)
  const result = await comparePresentationPageStructure(
    Buffer.from(full.bytes).toString('base64'),
    2,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    Buffer.from(host.bytes).toString('base64'),
  )
  expect(result.content).toMatchObject({
    status: 'incomplete',
    mediaChanged: [],
    unchecked: ['image'],
  })
})

it('matches text, shapes, images, tables and charts from compiled pages to Office shape readback', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const opened = await openPptx(bytes)
  for (const [pageIndex, source] of opened.deck.slides.entries()) {
    const shapes = source.elements.map((element, index) => ({
      id: String(index + 1),
      name: element.name!,
      type:
        element.type === 'table'
          ? 'Table'
          : element.type === 'chart'
            ? 'Chart'
            : element.type === 'picture'
              ? 'Image'
              : 'TextBox',
      left: (element.transform.offset.x * 72) / 914400,
      top: (element.transform.offset.y * 72) / 914400,
      width: (element.transform.offset.cx * 72) / 914400,
      height: (element.transform.offset.cy * 72) / 914400,
    }))
    const result = await comparePresentationPageStructure(
      Buffer.from(bytes).toString('base64'),
      pageIndex,
      {
        slideId: 'host',
        slideWidth: 960,
        slideHeight: 540,
        shapes,
        shapesTruncated: false,
        overflows: [],
        overlaps: [],
        overlapsTruncated: false,
        screenshot: { mime: 'image/png', base64: '' },
      },
    )
    expect(result).toMatchObject({
      structureStatus: 'passed',
      sourceCount: 3,
      hostCount: 3,
      issues: [],
    })
  }
})

it('reports missing, retyped and moved editable objects without a false pass', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const opened = await openPptx(bytes)
  const source = opened.deck.slides[6]!
  const chart = source.elements.find((element) => element.type === 'chart')!
  const result = await comparePresentationPageStructure(Buffer.from(bytes).toString('base64'), 6, {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes: [
      { id: '1', name: chart.name!, type: 'Image', left: 100, top: 100, width: 100, height: 100 },
    ],
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png', base64: '' },
  })
  expect(result.structureStatus).toBe('warning')
  expect(result.issues).toEqual(
    expect.arrayContaining([
      { name: 'chart', kind: 'type_changed', sourceType: 'chart', hostType: 'Image' },
      { name: 'title', kind: 'missing' },
    ]),
  )
})

it('marks a truncated host readback incomplete even when known objects match', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(bytes)).deck.slides[0]!
  const result = await comparePresentationPageStructure(Buffer.from(bytes).toString('base64'), 0, {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes: source.elements.map((element, index) => ({
      id: String(index),
      name: element.name!,
      type: 'TextBox',
      left: (element.transform.offset.x * 72) / 914400,
      top: (element.transform.offset.y * 72) / 914400,
      width: (element.transform.offset.cx * 72) / 914400,
      height: (element.transform.offset.cy * 72) / 914400,
    })),
    shapesTruncated: true,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png', base64: '' },
  })
  expect(result.status).toBe('incomplete')
})

it('flags a native object moved beyond the host conversion tolerance', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  shapes[0]!.left += 5
  const result = await comparePresentationPageStructure(Buffer.from(bytes).toString('base64'), 0, {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes,
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png', base64: '' },
  })
  expect(result).toMatchObject({
    structureStatus: 'warning',
    issues: [{ name: 'title', kind: 'geometry_changed' }],
  })
})

it('does not combine inconsistent Office shape and package readbacks into a pass', async () => {
  const { bytes } = await compilePresentationDeck(benchmarkDeck())
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index),
    name: element.name!,
    type: 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  shapes[0]!.left += 5
  const base64 = Buffer.from(bytes).toString('base64')
  const result = await comparePresentationPageStructure(
    base64,
    0,
    {
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    },
    base64,
  )
  expect(result).toMatchObject({ status: 'incomplete', readbackConsistent: false })
  expect(result.content.status).toBe('incomplete')
})

it('exposes the comparison for the exact imported production page without changing QA records', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[5]!]
  const { bytes } = await compilePresentationDeck(deck)
  const base64 = Buffer.from(bytes).toString('base64')
  const source = (await openPptx(bytes)).deck.slides[0]!
  const shapes = source.elements.map((element, index) => ({
    id: String(index + 1),
    name: element.name!,
    type: element.type === 'table' ? 'Table' : 'TextBox',
    left: (element.transform.offset.x * 72) / 914400,
    top: (element.transform.offset.y * 72) / 914400,
    width: (element.transform.offset.cx * 72) / 914400,
    height: (element.transform.offset.cy * 72) / 914400,
  }))
  const artifact = {
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pptxBase64: '',
    pagePptxBase64: [base64],
    planRevision: 1,
    slideCount: 1,
    pages: [{ id: 'page', title: 'Table', sourceSlideId: '256#' }],
  }
  const receipt = {
    state: 'complete' as const,
    documentId: 'doc',
    slideIds: ['host'],
    checkpoint: {
      version: 2 as const,
      artifactDigest: createHash('sha256')
        .update(presentationArtifactContent(artifact))
        .digest('hex'),
      pageIds: ['page'],
      sourceSlideIds: ['256#'],
      baselineSlideIds: [],
      completed: [{ sourceSlideId: '256#', slideId: 'host' }],
    },
  }
  const skill = createPresentationQaSkill({
    available: () => true,
    artifact: () => artifact,
    documentId: async () => 'doc',
    readReceipt: () => receipt,
    inspectPage: async () => ({
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes,
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: '' },
    }),
    exportPage: async (slideId) => ({ slideId, base64 }),
    readQa: () => undefined,
    writeQa: async () => {
      throw new Error('unexpected write')
    },
    vfs: new InMemoryVfs(),
  })
  const result = await skill.executeTool({
    id: 'compare',
    name: 'compare_presentation_page_structure',
    input: { page_id: 'page' },
  })
  expect(result.isError).not.toBe(true)
  expect(JSON.parse(result.output)).toMatchObject({
    status: 'passed',
    pageId: 'page',
    hostSlideId: 'host',
    readbackConsistent: true,
    content: { status: 'passed', changed: [], unchecked: [] },
  })
})
