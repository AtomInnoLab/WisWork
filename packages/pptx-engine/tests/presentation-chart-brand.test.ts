import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { XMLBuilder, XMLParser } from 'fast-xml-parser'
import { parseChartXml } from '../src/chart'
import {
  compilePresentationDeck,
  verifyCompiledPresentationStructure,
} from '../src/presentation-compiler'
import { benchmarkDeck } from './fixtures/presentation-benchmark'
const parser = new XMLParser({ ignoreAttributes: false, parseTagValue: false })
type Xml = Record<string, any>
type ChartType = 'bar' | 'line' | 'pie'
const items = (value: Xml | Xml[]): Xml[] => (Array.isArray(value) ? value : [value])
function fixture(chartType: ChartType, multiple = chartType !== 'pie') {
  const deck = benchmarkDeck()
  deck.style = {
    ...deck.style,
    fontFace: 'Courier New',
    textColor: 'AB12CD',
    background: '102030',
    accentColor: '459ABC',
  }
  deck.slides = [deck.slides[6]!]
  const chart = deck.slides[0]!.elements.find((e) => e.kind === 'chart')!
  if (chart.kind !== 'chart') throw Error('fixture')
  chart.chartType = chartType
  if (multiple) chart.series.push({ name: 'Second', values: [100, 80] })
  return deck
}
const plot = (root: Xml) => root['c:chart']['c:plotArea']
const series = (root: Xml, type: ChartType) => items(plot(root)[`c:${type}Chart`]['c:ser'])
const label = (node: Xml) => node['c:txPr']['a:p']['a:pPr']['a:defRPr']
const rgb = (node: Xml) => node['a:solidFill']['a:srgbClr']
async function compiled(type: ChartType, multiple = type !== 'pie') {
  const deck = fixture(type, multiple),
    result = await compilePresentationDeck(deck),
    zip = await JSZip.loadAsync(result.bytes)
  const path = Object.keys(zip.files).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))!
  const xml = await zip.file(path)!.async('string'),
    document = parser.parse(xml)
  return { deck, zip, path, xml, document, root: document['c:chartSpace'] as Xml }
}
it.each(['bar', 'line', 'pie'] as const)(
  'emits all supported %s visible roles using the style without a brand plan',
  async (type) => {
    const { deck, zip, root } = await compiled(type)
    for (const area of [root['c:spPr'], plot(root)['c:spPr']])
      expect(rgb(area)['@_val']).toBe(deck.style.background)
    const chart = plot(root)[`c:${type}Chart`]
    for (const node of [chart['c:dLbls'], ...series(root, type).map((s) => s['c:dLbls'])].filter(
      Boolean,
    )) {
      expect(rgb(label(node))['@_val']).toBe(deck.style.textColor)
      expect(label(node)['a:latin']['@_typeface']).toBe(deck.style.fontFace)
      for (const point of node['c:dLbl'] ? items(node['c:dLbl']) : []) {
        expect(rgb(label(point))['@_val']).toBe(deck.style.textColor)
        expect(label(point)['a:latin']['@_typeface']).toBe(deck.style.fontFace)
      }
    }
    for (const [index, s] of series(root, type).entries()) {
      const expectedSeriesColor =
        type === 'pie' || index % 2 === 0 ? deck.style.accentColor : deck.style.textColor
      expect(rgb(s['c:spPr'])['@_val']).toBe(expectedSeriesColor)
      if (type === 'pie')
        for (const point of items(s['c:dPt'])) {
          expect(rgb(point['c:spPr'])['@_val']).toBe(deck.style.accentColor)
          expect(rgb(point['c:spPr']['a:ln'])['@_val']).toBe(deck.style.accentColor)
        }
      if (type === 'line')
        expect(rgb(s['c:marker']['c:spPr']['a:ln'])['@_val']).toBe(expectedSeriesColor)
    }
    if (type === 'pie') {
      expect(root['c:chart']['c:legend']).toBeUndefined()
      expect(plot(root)['c:catAx']).toBeUndefined()
      expect(plot(root)['c:valAx']).toBeUndefined()
    } else {
      expect(rgb(label(root['c:chart']['c:legend']))['@_val']).toBe(deck.style.textColor)
      for (const axis of [plot(root)['c:catAx'], plot(root)['c:valAx']]) {
        expect(rgb(label(axis))['@_val']).toBe(deck.style.textColor)
        expect(rgb(axis['c:spPr']['a:ln'])['@_val']).toBe(deck.style.textColor)
      }
      expect(rgb(plot(root)['c:valAx']['c:majorGridlines']['c:spPr']['a:ln'])['@_val']).toBe(
        deck.style.textColor,
      )
    }
    await expect(verifyCompiledPresentationStructure(zip, deck)).resolves.toBeUndefined()
  },
)
const mutations: { name: string; type: ChartType; change(root: Xml): void }[] = [
  {
    name: 'unsupported trendline',
    type: 'bar',
    change: (r) =>
      (series(r, 'bar')[0]!['c:trendline'] = {
        'c:trendlineType': { '@_val': 'linear' },
        'c:spPr': { 'a:ln': { 'a:solidFill': { 'a:srgbClr': { '@_val': 'DEADBE' } } } },
      }),
  },
  {
    name: 'unsupported error bars',
    type: 'line',
    change: (r) =>
      (series(r, 'line')[0]!['c:errBars'] = {
        'c:errDir': { '@_val': 'y' },
        'c:errBarType': { '@_val': 'both' },
        'c:errValType': { '@_val': 'fixedVal' },
        'c:val': { '@_val': '30' },
      }),
  },
  {
    name: 'inverted default negative color',
    type: 'bar',
    change: (r) => (series(r, 'bar')[0]!['c:invertIfNegative']['@_val'] = '1'),
  },
  {
    name: 'varying default series colors',
    type: 'bar',
    change: (r) => (plot(r)['c:barChart']['c:varyColors']['@_val'] = '1'),
  },
  {
    name: 'smoothed default line',
    type: 'line',
    change: (r) => (series(r, 'line')[0]!['c:smooth']['@_val'] = '1'),
  },
  {
    name: 'plot background',
    type: 'bar',
    change: (r) => (rgb(plot(r)['c:spPr'])['@_val'] = 'FFFFFF'),
  },
  { name: 'chart background', type: 'line', change: (r) => (rgb(r['c:spPr'])['@_val'] = 'FFFFFF') },
  {
    name: 'background opacity',
    type: 'pie',
    change: (r) => (rgb(r['c:spPr'])['a:alpha'] = { '@_val': '0' }),
  },
  {
    name: 'global data label color',
    type: 'bar',
    change: (r) => (rgb(label(plot(r)['c:barChart']['c:dLbls']))['@_val'] = '000000'),
  },
  {
    name: 'second series label font',
    type: 'line',
    change: (r) => (label(series(r, 'line')[1]!['c:dLbls'])['a:latin']['@_typeface'] = 'Arial'),
  },
  {
    name: 'pie group label font',
    type: 'pie',
    change: (r) => (label(series(r, 'pie')[0]!['c:dLbls'])['a:latin']['@_typeface'] = 'Arial'),
  },
  {
    name: 'pie point label color',
    type: 'pie',
    change: (r) =>
      (rgb(label(items(series(r, 'pie')[0]!['c:dLbls']['c:dLbl'])[1]!))['@_val'] = '000000'),
  },
  {
    name: 'label alpha transform',
    type: 'bar',
    change: (r) => (rgb(label(plot(r)['c:barChart']['c:dLbls']))['a:alpha'] = { '@_val': '0' }),
  },
  {
    name: 'label theme inheritance',
    type: 'bar',
    change: (r) =>
      (label(plot(r)['c:barChart']['c:dLbls'])['a:solidFill'] = {
        'a:schemeClr': { '@_val': 'tx1' },
      }),
  },
  {
    name: 'label local run override',
    type: 'line',
    change: (r) =>
      (series(r, 'line')[0]!['c:dLbls']['c:txPr']['a:p']['a:r'] = {
        'a:rPr': { 'a:latin': { '@_typeface': 'Arial' } },
        'a:t': 'Override',
      }),
  },
  {
    name: 'legend color',
    type: 'bar',
    change: (r) => (rgb(label(r['c:chart']['c:legend']))['@_val'] = '000000'),
  },
  {
    name: 'legend effects',
    type: 'line',
    change: (r) =>
      (r['c:chart']['c:legend']['c:spPr'] = { 'a:effectLst': { 'a:glow': { '@_rad': '10000' } } }),
  },
  {
    name: 'category axis font',
    type: 'bar',
    change: (r) => (label(plot(r)['c:catAx'])['a:latin']['@_typeface'] = 'Arial'),
  },
  {
    name: 'value axis label color',
    type: 'line',
    change: (r) => (rgb(label(plot(r)['c:valAx']))['@_val'] = '000000'),
  },
  {
    name: 'axis line color',
    type: 'bar',
    change: (r) => (rgb(plot(r)['c:catAx']['c:spPr']['a:ln'])['@_val'] = '888888'),
  },
  {
    name: 'gridline color',
    type: 'line',
    change: (r) =>
      (rgb(plot(r)['c:valAx']['c:majorGridlines']['c:spPr']['a:ln'])['@_val'] = '888888'),
  },
  {
    name: 'gridline alpha',
    type: 'bar',
    change: (r) =>
      (rgb(plot(r)['c:valAx']['c:majorGridlines']['c:spPr']['a:ln'])['a:alpha'] = { '@_val': '0' }),
  },
  {
    name: 'series fill',
    type: 'bar',
    change: (r) => (rgb(series(r, 'bar')[1]!['c:spPr'])['@_val'] = 'FF0000'),
  },
  {
    name: 'series line',
    type: 'line',
    change: (r) => (rgb(series(r, 'line')[0]!['c:spPr']['a:ln'])['@_val'] = 'FF0000'),
  },
  {
    name: 'marker fill',
    type: 'line',
    change: (r) => (rgb(series(r, 'line')[1]!['c:marker']['c:spPr'])['@_val'] = 'FF0000'),
  },
  {
    name: 'marker line opacity',
    type: 'line',
    change: (r) =>
      (rgb(series(r, 'line')[0]!['c:marker']['c:spPr']['a:ln'])['a:alpha'] = { '@_val': '0' }),
  },
  {
    name: 'pie parent outline',
    type: 'pie',
    change: (r) => (rgb(series(r, 'pie')[0]!['c:spPr']['a:ln'])['@_val'] = 'F9F9F9'),
  },
  {
    name: 'pie point outline',
    type: 'pie',
    change: (r) =>
      (rgb(items(series(r, 'pie')[0]!['c:dPt'])[1]!['c:spPr']['a:ln'])['@_val'] = 'F9F9F9'),
  },
  {
    name: 'pie point tint',
    type: 'pie',
    change: (r) =>
      (rgb(items(series(r, 'pie')[0]!['c:dPt'])[0]!['c:spPr'])['a:tint'] = { '@_val': '50000' }),
  },
  {
    name: '3D wall effect',
    type: 'bar',
    change: (r) =>
      (r['c:chart']['c:backWall'] = {
        'c:spPr': { 'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } } },
      }),
  },
]
it.each(mutations)(
  'rejects $type $name changes with the same data/workbook/geometry',
  async ({ type, change }) => {
    const { deck, zip, path, xml, document, root } = await compiled(type)
    const workbookPath = Object.keys(zip.files).find((p) => /^ppt\/embeddings\/.+\.xlsx$/.test(p))!,
      workbook = await zip.file(workbookPath)!.async('uint8array'),
      slide = await zip.file('ppt/slides/slide1.xml')!.async('string')
    change(root)
    const changedXml = new XMLBuilder({ ignoreAttributes: false }).build(document)
    const before = parseChartXml(xml),
      after = parseChartXml(changedXml)
    if (!before || !after) throw Error('invalid chart fixture')
    expect(after.categories).toEqual(before.categories)
    expect(after.series.map(({ name, values }) => ({ name, values }))).toEqual(
      before.series.map(({ name, values }) => ({ name, values })),
    )
    zip.file(path, changedXml)
    const reopened = await JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }))
    expect(await reopened.file(workbookPath)!.async('uint8array')).toEqual(workbook)
    expect(await reopened.file('ppt/slides/slide1.xml')!.async('string')).toBe(slide)
    await expect(verifyCompiledPresentationStructure(reopened, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
  },
)
it.each(['bar', 'line', 'pie'] as const)(
  'accepts mixed-case %s colors and a single-series chart without an unused legend',
  async (type) => {
    const deck = fixture(type, false)
    deck.style.textColor = 'aB12cD'
    deck.style.background = 'aBcDeF'
    deck.style.accentColor = '45aBcD'
    const result = await compilePresentationDeck(deck),
      zip = await JSZip.loadAsync(result.bytes)
    await expect(verifyCompiledPresentationStructure(zip, deck)).resolves.toBeUndefined()
  },
)
it('does not scan unused theme palette entries as visible chart style', async () => {
  const { zip, deck } = await compiled('pie')
  const theme = await zip.file('ppt/theme/theme1.xml')!.async('string')
  zip.file('ppt/theme/theme1.xml', theme.replace(/(<a:accent6><a:srgbClr val=")[^"]+/, '$1DEADBE'))
  await expect(verifyCompiledPresentationStructure(zip, deck)).resolves.toBeUndefined()
})
it('rejects empty category labels at the existing SlideIR boundary', async () => {
  const deck = fixture('bar'),
    chart = deck.slides[0]!.elements.find((e) => e.kind === 'chart')!
  if (chart.kind !== 'chart') throw Error('fixture')
  chart.categories[0] = ''
  await expect(compilePresentationDeck(deck)).rejects.toThrow('presentation_invalid')
})
it.each(['bar', 'line'] as const)(
  'supports eight %s series with the maximum category count',
  async (type) => {
    const deck = fixture(type, false),
      chart = deck.slides[0]!.elements.find((e) => e.kind === 'chart')!
    if (chart.kind !== 'chart') throw Error('fixture')
    chart.categories = Array.from({ length: 50 }, (_, i) => `Category ${i + 1}`)
    chart.series = Array.from({ length: 8 }, (_, i) => ({
      name: `Series ${i + 1}`,
      values: chart.categories.map((_, j) => (i + 1) * (j + 1)),
    }))
    const result = await compilePresentationDeck(deck)
    await expect(
      verifyCompiledPresentationStructure(await JSZip.loadAsync(result.bytes), deck),
    ).resolves.toBeUndefined()
  },
)
it.each([32, 50])(
  'supports %i pie categories with complete explicit point and label styling',
  async (count) => {
    const deck = fixture('pie'),
      chart = deck.slides[0]!.elements.find((e) => e.kind === 'chart')!
    if (chart.kind !== 'chart') throw Error('fixture')
    chart.categories = Array.from({ length: count }, (_, i) => `Slice ${i + 1}`)
    chart.series[0]!.values = chart.categories.map((_, i) => i + 1)
    const result = await compilePresentationDeck(deck)
    await expect(
      verifyCompiledPresentationStructure(await JSZip.loadAsync(result.bytes), deck),
    ).resolves.toBeUndefined()
  },
)
it.each(['line', 'pie'] as const)(
  'uses the resolved fallback font in every %s label',
  async (type) => {
    const deck = fixture(type),
      requested = 'Missing Brand Font'
    deck.style.fontFace = requested
    deck.style.fontFallbacks = ['Courier New']
    const result = await compilePresentationDeck(deck, {
      fontAvailable: (font) => font === 'Courier New',
    })
    expect(result.report.fontResolution).toEqual({
      requested,
      used: 'Courier New',
      substituted: true,
    })
    const zip = await JSZip.loadAsync(result.bytes),
      path = Object.keys(zip.files).find((p) => /^ppt\/charts\/chart\d+\.xml$/.test(p))!,
      xml = await zip.file(path)!.async('string')
    expect(xml).not.toContain(requested)
    expect(xml).toContain('typeface="Courier New"')
    await expect(
      verifyCompiledPresentationStructure(zip, {
        ...deck,
        style: { ...deck.style, fontFace: 'Courier New' },
      }),
    ).resolves.toBeUndefined()
  },
)
it('preserves the pre-existing rejection of multiple pie series', async () => {
  await expect(compilePresentationDeck(fixture('pie', true))).rejects.toThrow(
    'presentation_invalid:pie_values',
  )
})
