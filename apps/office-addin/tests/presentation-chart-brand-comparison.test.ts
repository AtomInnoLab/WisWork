import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { openPptx } from '@wiswork/pptx-engine'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark.js'
import { inspectPowerPointComplexPagePackage } from '../src/skills/powerpoint/presentation-complex-page-package.js'
import { comparePresentationPageStructure } from '../src/skills/powerpoint/presentation-structure-comparison.js'
async function fixture(
  chartType: 'bar' | 'line' | 'pie' = 'bar',
  categoryCount = 2,
  withTruncatedTable = false,
  seriesCount = 2,
) {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[6]!]
  const chart = deck.slides[0]!.elements.find((item) => item.kind === 'chart')!
  if (chart.kind !== 'chart') throw Error('chart fixture missing')
  chart.chartType = chartType
  chart.categories = Array.from({ length: categoryCount }, (_, i) => `类${i + 1}`)
  chart.series[0]!.values = Array.from({ length: categoryCount }, (_, i) => 100 + i)
  if (withTruncatedTable) {
    const table = structuredClone(
      benchmarkDeck().slides[5]!.elements.find((item) => item.kind === 'table')!,
    )
    if (table.kind !== 'table') throw Error('table fixture missing')
    table.rows[0]![0] = '长'.repeat(200)
    deck.slides[0]!.elements.unshift(table)
  }
  if (chartType !== 'pie')
    for (let index = 1; index < seriesCount; index++)
      chart.series.push({
        name: `组${index}`,
        values: Array.from({ length: categoryCount }, (_, i) => 80 + i + index),
      })
  const { bytes } = await compilePresentationDeck(deck),
    source = Buffer.from(bytes).toString('base64')
  const page = (await openPptx(bytes)).deck.slides[0]!
  const host = {
    slideId: 'host',
    slideWidth: 960,
    slideHeight: 540,
    shapes: page.elements.map((element, index) => ({
      id: String(index),
      name: element.name!,
      type: element.type === 'chart' ? 'Chart' : element.type === 'table' ? 'Table' : 'TextBox',
      left: (element.transform.offset.x * 72) / 914400,
      top: (element.transform.offset.y * 72) / 914400,
      width: (element.transform.offset.cx * 72) / 914400,
      height: (element.transform.offset.cy * 72) / 914400,
    })),
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: '' },
  }
  const zip = await JSZip.loadAsync(bytes),
    path = Object.keys(zip.files).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))!,
    xml = await zip.file(path)!.async('string')
  const compare = async (edited: string, originalXml?: string) => {
    const exportZip = await JSZip.loadAsync(bytes)
    exportZip.file(path, edited)
    const sourceZip = await JSZip.loadAsync(bytes)
    if (originalXml) sourceZip.file(path, originalXml)
    return comparePresentationPageStructure(
      originalXml ? await sourceZip.generateAsync({ type: 'base64' }) : source,
      0,
      host,
      await exportZip.generateAsync({ type: 'base64' }),
    )
  }
  return { xml, compare, source }
}
it.each(['bar', 'line', 'pie'] as const)(
  'detects actual %s series palette and visible text style drift with unchanged data and geometry',
  async (type) => {
    const f = await fixture(type)
    for (const xml of [
      f.xml.replace(/(<a:srgbClr val=")2255AA/, '$1FF00AA'),
      f.xml.replace(/(<a:latin typeface=")[^"]+/, '$1Courier New'),
    ]) {
      expect(xml).not.toBe(f.xml)
      const result = await f.compare(xml)
      expect(result.readbackConsistent).toBe(true)
      expect(result.issues).toEqual([])
      expect(result.content).toMatchObject({
        chartStyleChanged: ['chart'],
        cacheChanged: [],
        unchecked: ['chart'],
      })
    }
  },
)
it('normalizes RGB case and harmless XML whitespace rather than comparing chart bytes', async () => {
  const f = await fixture()
  const rewritten = f.xml
    .replace(
      /(<a:srgbClr val=")([A-Fa-f0-9]{6})/g,
      (_all, prefix, color) => prefix + color.toLowerCase(),
    )
    .replaceAll('><', '>\n<')
  const result = await f.compare(rewritten)
  expect(result.content.chartStyleChanged).toEqual([])
  expect(result.content.cacheChanged).toEqual([])
  expect(result.content.unchecked).toContain('chart')
})
it('keeps transformed colors and inherited/theme styles unverified', async () => {
  const f = await fixture()
  for (const edited of [
    f.xml.replace(
      /<a:srgbClr val="([A-Fa-f0-9]{6})"\/>/,
      '<a:srgbClr val="$1"><a:alpha val="50000"/></a:srgbClr>',
    ),
    f.xml.replace(/<a:srgbClr val="[A-Fa-f0-9]{6}"\/>/, '<a:schemeClr val="accent1"/>'),
  ]) {
    const result = await f.compare(edited)
    expect(result.content.unchecked).toContain('chart')
    expect(result.content.chartStyleChanged).toEqual([])
    expect(result.content.status).not.toBe('passed')
  }
})
it('detects explicit axis, gridline, background, legend and data-label drift without altering caches', async () => {
  const f = await fixture()
  const segment = (tag: string, replace: (text: string) => string) =>
    f.xml.replace(new RegExp(`<${tag}\\b[^>]*>[\\s\\S]*?<\\/${tag}>`), replace)
  const edits = [
    segment('c:catAx', (text) => text.replace(/(<a:srgbClr val=")[A-Fa-f0-9]{6}/, '$1CC3300')),
    segment('c:majorGridlines', (text) =>
      text.replace(/(<a:srgbClr val=")[A-Fa-f0-9]{6}/, '$1CC3300'),
    ),
    segment('c:legend', (text) => text.replace(/(<a:latin typeface=")[^"]+/, '$1Courier New')),
    segment('c:dLbls', (text) => text.replace(/(<a:srgbClr val=")[A-Fa-f0-9]{6}/, '$1CC3300')),
    f.xml.replace(
      /(<\/c:chart>\s*<c:spPr>\s*<a:solidFill>\s*<a:srgbClr val=")[A-Fa-f0-9]{6}/,
      '$1DDDDDD',
    ),
  ]
  for (const edited of edits) {
    expect(edited).not.toBe(f.xml)
    const result = await f.compare(edited)
    expect(result.content.chartStyleChanged).toEqual(['chart'])
    expect(result.content.cacheChanged).toEqual([])
    expect(result.readbackConsistent).toBe(true)
  }
})
it('keeps all 50 valid pie points and their individual labels within the bounded projection', async () => {
  const f = await fixture('pie', 50)
  const projection = (
    await inspectPowerPointComplexPagePackage(f.source, undefined, { includeVisibleStyle: true })
  ).charts[0]!.visibleStyle!
  expect(projection).toBeDefined()
  expect(projection['series.0.labels.point.49.text.font']).toBe('microsoft yahei')
  expect(projection['series.0.point.49.fill']).toBe('2255AA')
  expect(new TextEncoder().encode(JSON.stringify(projection)).length).toBeLessThanOrEqual(64 * 1024)
  const edited = f.xml.replace(
    /(<c:dLbl>\s*<c:idx val="49"(?:\/?>)[\s\S]*?<a:latin typeface=")[^"]+/,
    '$1Courier New',
  )
  expect(edited).not.toBe(f.xml)
  expect((await f.compare(edited)).content.chartStyleChanged).toEqual(['chart'])
})
it.each(['bar', 'line'] as const)(
  'compares all ten legal %s series independently of cache truncation',
  async (type) => {
    const f = await fixture(type, 50, false, 10)
    const summary = await inspectPowerPointComplexPagePackage(f.source, undefined, {
      includeVisibleStyle: true,
    })
    expect(summary.charts[0]!.truncated).toBe(true)
    expect(
      summary.charts[0]!.visibleStyle!['series.9.line.color'] ??
        summary.charts[0]!.visibleStyle!['series.9.fill'],
    ).toBe('172033')
    const parts = f.xml.match(/<c:ser>[\s\S]*?<\/c:ser>/g)!
    const edited = f.xml.replace(
      parts[9]!,
      parts[9]!.replace(/(<a:srgbClr val=")172033/, '$1FF00AA'),
    )
    expect(edited).not.toBe(f.xml)
    expect((await f.compare(edited)).content.chartStyleChanged).toEqual(['chart'])
  },
)
it('reports known chart palette drift even when a preceding table truncates the legacy text summary', async () => {
  const f = await fixture('bar', 2, true)
  expect((await inspectPowerPointComplexPagePackage(f.source)).truncated).toBe(true)
  const edited = f.xml.replace(/(<a:srgbClr val=")2255AA/, '$1FF00AA')
  expect((await f.compare(edited)).content.chartStyleChanged).toEqual(['chart'])
})

it('keeps visible styles private to explicit comparison inspection', async () => {
  const f = await fixture('pie', 50)
  const summary = await inspectPowerPointComplexPagePackage(f.source)
  expect(summary.charts[0]).not.toHaveProperty('visibleStyle')
  expect(
    (await inspectPowerPointComplexPagePackage(f.source, undefined, { includeVisibleStyle: true }))
      .charts[0],
  ).toHaveProperty('visibleStyle')
})

it('compares added and removed explicit point overrides against known series inheritance', async () => {
  const f = await fixture('bar')
  const add = (color: string) =>
    f.xml.replace(
      '</c:ser>',
      `<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:dPt></c:ser>`,
    )
  expect((await f.compare(add('2255AA'))).content.chartStyleChanged).toEqual([])
  expect((await f.compare(add('EE1122'))).content.chartStyleChanged).toEqual(['chart'])
  const pie = await fixture('pie')
  const removed = pie.xml.replace(/<c:dPt>[\s\S]*?<\/c:dPt>/, '')
  expect((await pie.compare(removed)).content.chartStyleChanged).toEqual([])
  const differing = pie.xml.replace(/(<c:dPt>[\s\S]*?<a:srgbClr val=")[^"]+/, '$1EE1122')
  expect((await pie.compare(differing)).content.chartStyleChanged).toEqual(['chart'])
})

it('detects deletion of a differing point override and preserves known group label and marker inheritance', async () => {
  const bar = await fixture('bar')
  const point =
    '<c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="EE1122"/></a:solidFill></c:spPr></c:dPt>'
  expect(
    (await bar.compare(bar.xml, bar.xml.replace('</c:ser>', `${point}</c:ser>`))).content
      .chartStyleChanged,
  ).toEqual(['chart'])
  const parentLabel = bar.xml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/)![0]
  const tx = parentLabel.match(/<c:txPr>[\s\S]*?<\/c:txPr>/)![0]
  const override = `<c:dLbl><c:idx val="0"/>${tx.replace(/(<a:latin typeface=")[^"]+/, '$1Courier New')}</c:dLbl>`
  const labelChanged = bar.xml.replace(
    parentLabel,
    parentLabel.replace('</c:dLbls>', `${override}</c:dLbls>`),
  )
  expect((await bar.compare(labelChanged)).content.chartStyleChanged).toEqual(['chart'])
  expect((await bar.compare(labelChanged, labelChanged)).content.chartStyleChanged).toEqual([])
  const line = await fixture('line')
  const marker = line.xml.match(/<c:marker>[\s\S]*?<\/c:marker>/)![0]
  const changedMarker = `<c:dPt><c:idx val="0"/>${marker.replace(/(<a:srgbClr val=")2255AA/, '$1EE1122')}</c:dPt>`
  expect(
    (await line.compare(line.xml.replace('</c:ser>', `${changedMarker}</c:ser>`))).content
      .chartStyleChanged,
  ).toEqual(['chart'])
  const sameMarker = `<c:dPt><c:idx val="0"/>${marker}</c:dPt>`
  expect(
    (await line.compare(line.xml.replace('</c:ser>', `${sameMarker}</c:ser>`))).content
      .chartStyleChanged,
  ).toEqual([])
})
