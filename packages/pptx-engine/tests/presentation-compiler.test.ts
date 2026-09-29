import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'
import { resolveTarget } from '../src/zip'
import { openPptx } from '../src/index'
import { parsePresentationDeck, inspectPresentationGeometry } from '../src/presentation'
import {
  compilePresentationDeck,
  verifyCompiledPresentationStructure,
} from '../src/presentation-compiler'
import { benchmarkDeck } from './fixtures/presentation-benchmark'

describe('presentation contract and compiler', () => {
  it('selects an installed fallback font and reports the substitution', async () => {
    const deck = benchmarkDeck()
    deck.style.fontFace = 'Unavailable Benchmark Font'
    deck.style.fontFallbacks = ['Noto Sans CJK SC', 'Arial']
    const result = await compilePresentationDeck(deck, {
      fontAvailable: (family) => family === 'Noto Sans CJK SC',
    })
    expect(result.report.fontResolution).toEqual({
      requested: 'Unavailable Benchmark Font',
      used: 'Noto Sans CJK SC',
      substituted: true,
    })
    const zip = await JSZip.loadAsync(result.bytes)
    expect(await zip.file('ppt/theme/theme1.xml')!.async('string')).toContain('Noto Sans CJK SC')
    expect(await zip.file('ppt/slides/slide1.xml')!.async('string')).not.toContain(
      'Unavailable Benchmark Font',
    )
  })

  it('rejects an unavailable font when no declared fallback is installed', async () => {
    const deck = benchmarkDeck()
    deck.style.fontFace = 'Unavailable Benchmark Font'
    deck.style.fontFallbacks = ['Also Unavailable']
    await expect(compilePresentationDeck(deck, { fontAvailable: () => false })).rejects.toThrow(
      'font_unavailable',
    )
  })
  it('rejects source attribution that cannot fit the visible footer without truncation', async () => {
    const deck = benchmarkDeck()
    deck.claims[0]!.source = '来源'.repeat(245)
    await expect(compilePresentationDeck(deck)).rejects.toThrow(
      'presentation_invalid:source_footer_overflow',
    )
  })
  it.each(['bar', 'line', 'pie'] as const)(
    'postflights native %s chart caches from the shared SlideIR',
    async (chartType) => {
      const deck = benchmarkDeck()
      const chart = deck.slides[6]!.elements[1]!
      if (chart.kind !== 'chart') throw new Error('invalid fixture')
      chart.chartType = chartType
      await expect(compilePresentationDeck(deck)).resolves.toMatchObject({
        report: { checks: { structure: 'passed' } },
      })
    },
  )
  it('starts native bar chart value axes at zero so differences are not visually exaggerated', async () => {
    const deck = benchmarkDeck()
    const chart = deck.slides[6]!.elements[1]!
    if (chart.kind !== 'chart') throw new Error('invalid fixture')
    chart.chartType = 'bar'
    chart.categories = ['downside', 'base', 'upside']
    chart.series[0]!.values = [371483, 402766, 422318]
    const { bytes } = await compilePresentationDeck(deck)
    const zip = await JSZip.loadAsync(bytes)
    const chartPath = Object.keys(zip.files).find((path) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(path),
    )
    expect(chartPath).toBeDefined()
    const xml = await zip.file(chartPath!)!.async('string')
    expect(xml).toMatch(/<c:valAx>[\s\S]*?<c:scaling>[\s\S]*?<c:min val="0"\/>/)
  })
  it('rejects compiled packages with missing or changed native page objects', async () => {
    const deck = benchmarkDeck()
    const { bytes } = await compilePresentationDeck(deck)
    const mutations: Array<{ slide: number; change: (xml: string) => string }> = [
      { slide: 1, change: (xml) => xml.replace('<a:t>科研汇报</a:t>', '<a:t>标题被修改</a:t>') },
      { slide: 6, change: (xml) => xml.replace('<a:t>120</a:t>', '<a:t>121</a:t>') },
      { slide: 7, change: (xml) => xml.replace('<c:chart ', '<c:broken ') },
      { slide: 1, change: (xml) => xml.replaceAll('x="914400"', 'x="1900000"') },
      {
        slide: 3,
        change: (xml) =>
          xml.replace(/(<p:pic[\s\S]*?<a:off x=")\d+/, (_, prefix: string) => `${prefix}1900000`),
      },
      { slide: 4, change: (xml) => xml.replace('prst="roundRect"', 'prst="ellipse"') },
    ]
    for (const { slide, change } of mutations) {
      const zip = await JSZip.loadAsync(bytes)
      const path = `ppt/slides/slide${slide}.xml`
      const original = await zip.file(path)!.async('string')
      const changed = change(original)
      expect(changed).not.toBe(original)
      zip.file(path, changed)
      await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
        'presentation_compile:structure_mismatch',
      )
    }
    const missing = await JSZip.loadAsync(bytes)
    missing.remove('ppt/slides/slide3.xml')
    await expect(verifyCompiledPresentationStructure(missing, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
    for (const [slide, kind] of [
      [3, 'image'],
      [7, 'chart'],
    ] as const) {
      const zip = await JSZip.loadAsync(bytes)
      const path = `ppt/slides/_rels/slide${slide}.xml.rels`
      const original = await zip.file(path)!.async('string')
      const rels = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' }).parse(
        original,
      ).Relationships.Relationship
      const relation = (Array.isArray(rels) ? rels : [rels]).find((entry) =>
        entry['@_Type']?.endsWith(`/${kind}`),
      )
      expect(relation).toBeDefined()
      const changed = original.replace(
        `Target="${relation['@_Target']}"`,
        'Target="../missing.bin"',
      )
      expect(changed).not.toBe(original)
      zip.file(path, changed)
      await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
        'presentation_compile:structure_mismatch',
      )
    }
    const wrongChart = await JSZip.loadAsync(bytes)
    const chartPath = Object.keys(wrongChart.files).find((name) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(name),
    )!
    const chartXml = await wrongChart.file(chartPath)!.async('string')
    const alteredChart = chartXml.replace('<c:v>120</c:v>', '<c:v>121</c:v>')
    expect(alteredChart).not.toBe(chartXml)
    wrongChart.file(chartPath, alteredChart)
    await expect(verifyCompiledPresentationStructure(wrongChart, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
    const wrongImage = await JSZip.loadAsync(bytes)
    const imageRelsXml = await wrongImage.file('ppt/slides/_rels/slide3.xml.rels')!.async('string')
    const imageRelations = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
    }).parse(imageRelsXml).Relationships.Relationship
    const imageRelation = (Array.isArray(imageRelations) ? imageRelations : [imageRelations]).find(
      (entry) => entry['@_Type']?.endsWith('/image'),
    )
    const imagePath = resolveTarget('ppt/slides/slide3.xml', imageRelation['@_Target'])
    wrongImage.file(imagePath, new Uint8Array([1, 2, 3]))
    await expect(verifyCompiledPresentationStructure(wrongImage, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
  })
  it('checks category caches for every chart series', async () => {
    const deck = benchmarkDeck()
    const chart = deck.slides[6]!.elements[1]!
    if (chart.kind !== 'chart') throw new Error('invalid fixture')
    chart.series.push({ name: '第二组', values: [100, 80] })
    const { bytes } = await compilePresentationDeck(deck)
    const zip = await JSZip.loadAsync(bytes)
    const path = Object.keys(zip.files).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))!
    const xml = await zip.file(path)!.async('string')
    let cacheNumber = 0
    const changed = xml.replace(/<c:multiLvlStrCache>[\s\S]*?<\/c:multiLvlStrCache>/g, (cache) =>
      ++cacheNumber === 2 ? cache.replace('<c:v>甲</c:v>', '<c:v>错误分类</c:v>') : cache,
    )
    expect(cacheNumber).toBe(2)
    expect(changed).not.toBe(xml)
    zip.file(path, changed)
    await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
  })
  it('rejects embedded chart data that disagrees with visible chart caches', async () => {
    const deck = benchmarkDeck()
    const chart = deck.slides[6]!.elements[1]!
    if (chart.kind !== 'chart') throw new Error('invalid fixture')
    chart.series.push({ name: '第二组', values: [100, 80] })
    const { bytes } = await compilePresentationDeck(deck)
    const original = await JSZip.loadAsync(bytes)
    const embedding = Object.keys(original.files).find((name) =>
      /^ppt\/embeddings\/[^/]+\.xlsx$/.test(name),
    )!
    const missing = await JSZip.loadAsync(bytes)
    missing.remove(embedding)
    await expect(verifyCompiledPresentationStructure(missing, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
    for (const [part, from, to] of [
      ['xl/worksheets/sheet1.xml', '<v>120</v>', '<v>121</v>'],
      ['xl/sharedStrings.xml', '<t>第二组</t>', '<t>错误组</t>'],
    ]) {
      const zip = await JSZip.loadAsync(bytes)
      const workbook = await JSZip.loadAsync(await zip.file(embedding)!.async('uint8array'))
      const xml = await workbook.file(part)!.async('string')
      expect(xml).toContain(from)
      workbook.file(part, xml.replace(from, to))
      zip.file(embedding, await workbook.generateAsync({ type: 'uint8array' }))
      await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
        'presentation_compile:structure_mismatch',
      )
    }
    const wrongReference = await JSZip.loadAsync(bytes)
    const chartPath = Object.keys(wrongReference.files).find((name) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(name),
    )!
    const chartXml = await wrongReference.file(chartPath)!.async('string')
    expect(chartXml).toContain('Sheet1!$B$2:$B$3')
    wrongReference.file(chartPath, chartXml.replace('Sheet1!$B$2:$B$3', 'Sheet1!$C$2:$C$3'))
    await expect(verifyCompiledPresentationStructure(wrongReference, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
  })
  it('does not treat an empty workbook value as an editable zero', async () => {
    const deck = benchmarkDeck()
    const chart = deck.slides[6]!.elements[1]!
    if (chart.kind !== 'chart') throw new Error('invalid fixture')
    chart.series[0]!.values[0] = 0
    const { bytes } = await compilePresentationDeck(deck)
    const zip = await JSZip.loadAsync(bytes)
    const embedding = Object.keys(zip.files).find((name) =>
      /^ppt\/embeddings\/[^/]+\.xlsx$/.test(name),
    )!
    const workbook = await JSZip.loadAsync(await zip.file(embedding)!.async('uint8array'))
    const path = 'xl/worksheets/sheet1.xml'
    const xml = await workbook.file(path)!.async('string')
    expect(xml).toContain('<c r="B2"><v>0</v></c>')
    workbook.file(path, xml.replace('<c r="B2"><v>0</v></c>', '<c r="B2"><v></v></c>'))
    zip.file(embedding, await workbook.generateAsync({ type: 'uint8array' }))
    await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
  })
  it('repairs zero values in every series across two native charts', async () => {
    const deck = benchmarkDeck()
    const first = deck.slides[6]!.elements[1]!
    if (first.kind !== 'chart') throw new Error('invalid fixture')
    first.series.push({ name: '第二组', values: [100, 0] })
    deck.slides[7]!.elements[1] = {
      ...structuredClone(first),
      id: 'second-chart',
      series: [{ name: '第三组', values: [0, 80] }],
    }
    const { bytes } = await compilePresentationDeck(deck)
    const zip = await JSZip.loadAsync(bytes)
    const workbooks = Object.keys(zip.files).filter((name) =>
      /^ppt\/embeddings\/[^/]+\.xlsx$/.test(name),
    )
    expect(workbooks).toHaveLength(2)
    for (const path of workbooks) {
      const workbook = await JSZip.loadAsync(await zip.file(path)!.async('uint8array'))
      const sheet = await workbook.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(sheet).not.toContain('<v></v>')
      expect(sheet).toContain('<v>0</v>')
    }
    await expect(verifyCompiledPresentationStructure(zip, deck)).resolves.toBeUndefined()
  })
  it('rejects charts with missing visible values or an unexpected legend', async () => {
    const deck = benchmarkDeck()
    const { bytes } = await compilePresentationDeck(deck)
    const zip = await JSZip.loadAsync(bytes)
    const path = Object.keys(zip.files).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))!
    const original = await zip.file(path)!.async('string')
    expect(original).toContain('<c:showVal val="1"')
    zip.file(path, original.replace('<c:showVal val="1"', '<c:showVal val="0"'))
    await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
    expect(original).not.toContain('<c:legend>')
    zip.file(path, original.replace('</c:chart>', '<c:legend/></c:chart>'))
    await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
    const chart = deck.slides[6]!.elements[1]!
    if (chart.kind !== 'chart') throw new Error('invalid fixture')
    chart.series.push({ name: '第二组', values: [100, 80] })
    const multi = await compilePresentationDeck(deck)
    const multiZip = await JSZip.loadAsync(multi.bytes)
    const multiPath = Object.keys(multiZip.files).find((name) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(name),
    )!
    const multiXml = await multiZip.file(multiPath)!.async('string')
    expect(multiXml).toContain('<c:legend>')
    multiZip.file(multiPath, multiXml.replace(/<c:legend>[\s\S]*?<\/c:legend>/, ''))
    await expect(verifyCompiledPresentationStructure(multiZip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
    chart.series.pop()
    chart.chartType = 'pie'
    const pie = await compilePresentationDeck(deck)
    const pieZip = await JSZip.loadAsync(pie.bytes)
    const piePath = Object.keys(pieZip.files).find((name) =>
      /^ppt\/charts\/chart\d+\.xml$/.test(name),
    )!
    const pieXml = await pieZip.file(piePath)!.async('string')
    const alteredPie = pieXml.replace(
      /(<c:dLbl>[\s\S]*?<c:showVal val=")1"/,
      (_, prefix: string) => `${prefix}0"`,
    )
    expect(alteredPie).not.toBe(pieXml)
    pieZip.file(piePath, alteredPie)
    await expect(verifyCompiledPresentationStructure(pieZip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
  })
  it('rejects changed image alt text and removed cover crop', async () => {
    const deck = benchmarkDeck()
    const image = deck.slides[2]!.elements.find((element) => element.kind === 'image')
    if (!image || image.kind !== 'image') throw new Error('invalid fixture')
    image.altText = '实验装置照片'
    image.fit = 'cover'
    image.w = 4
    image.h = 2
    const { bytes } = await compilePresentationDeck(deck)
    const zip = await JSZip.loadAsync(bytes)
    const path = 'ppt/slides/slide3.xml'
    const original = await zip.file(path)!.async('string')
    expect(original).toContain('descr="实验装置照片"')
    expect(original).toContain('<a:srcRect')
    zip.file(path, original.replace('descr="实验装置照片"', 'descr="错误说明"'))
    await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
    zip.file(path, original.replace(/<a:srcRect[^>]*\/>/, ''))
    await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
    const alteredCrop = original.replace(
      /(<a:srcRect[^>]*\bt=")\d+/,
      (_, prefix: string) => `${prefix}5000`,
    )
    expect(alteredCrop).not.toBe(original)
    zip.file(path, alteredCrop)
    await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
    image.fit = 'contain'
    const contained = await compilePresentationDeck(deck)
    const containZip = await JSZip.loadAsync(contained.bytes)
    const containXml = await containZip.file(path)!.async('string')
    const alteredContain = containXml.replace('<p:blipFill>', '<p:blipFill><a:srcRect t="25000"/>')
    expect(alteredContain).not.toBe(containXml)
    containZip.file(path, alteredContain)
    await expect(verifyCompiledPresentationStructure(containZip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
  })
  it('rejects changed native background, text, shape and table styles', async () => {
    const deck = benchmarkDeck()
    const { bytes } = await compilePresentationDeck(deck)
    const mutations: Array<{ slide: number; before: string; after: string }> = [
      {
        slide: 1,
        before: '<p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"',
        after: '<p:bgPr><a:solidFill><a:srgbClr val="000000"',
      },
      { slide: 1, before: 'typeface="Microsoft YaHei"', after: 'typeface="Arial"' },
      { slide: 1, before: 'sz="3200"', after: 'sz="2800"' },
      { slide: 1, before: 'val="172033"', after: 'val="FF0000"' },
      { slide: 4, before: 'val="2255AA"', after: 'val="FF0000"' },
      { slide: 6, before: 'sz="1600"', after: 'sz="1200"' },
    ]
    for (const { slide, before, after } of mutations) {
      const zip = await JSZip.loadAsync(bytes)
      const path = `ppt/slides/slide${slide}.xml`
      const original = await zip.file(path)!.async('string')
      expect(original).toContain(before)
      zip.file(path, original.replace(before, after))
      await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
        'presentation_compile:structure_mismatch',
      )
    }
  })
  it('accepts a valid empty table cell with paragraph-level font styling', async () => {
    const deck = benchmarkDeck()
    const table = deck.slides[5]!.elements[1]!
    if (table.kind !== 'table') throw new Error('invalid fixture')
    table.rows[1]![1] = ''
    await expect(compilePresentationDeck(deck)).resolves.toMatchObject({
      report: { checks: { structure: 'passed' } },
    })
  })
  it('accepts explicitly styled native text and shapes', async () => {
    const deck = benchmarkDeck()
    const title = deck.slides[0]!.elements[0]!
    const shape = deck.slides[3]!.elements[1]!
    const table = deck.slides[5]!.elements[1]!
    if (title.kind !== 'text' || shape.kind !== 'shape' || table.kind !== 'table')
      throw new Error('invalid fixture')
    Object.assign(title, { color: 'AA1122', fontSize: 28, bold: true, align: 'center' })
    Object.assign(shape, { fill: '00AA22', lineColor: '1122AA' })
    table.fontSize = 12
    await expect(compilePresentationDeck(deck)).resolves.toMatchObject({
      report: { checks: { structure: 'passed' } },
    })
  })
  it('accepts contract-valid lowercase hex colors', async () => {
    const deck = benchmarkDeck()
    deck.style.background = 'ffffff'
    deck.style.textColor = '1720aa'
    deck.style.accentColor = '2255aa'
    await expect(compilePresentationDeck(deck)).resolves.toMatchObject({
      report: { checks: { structure: 'passed' } },
    })
  })
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
    const chart = Object.keys(zip.files).find((name) => /^ppt\/charts\/chart\d+\.xml$/.test(name))
    expect(chart).toBeDefined()
    expect(await zip.file(chart!)!.async('string')).toContain('120')
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
    deck.assets[0]!.licenseEvidence = `attachment:${'a'.repeat(64)}`
    const image = deck.slides[2]!.elements[1]!
    if (image.kind !== 'image') throw new Error('invalid fixture')
    image.altText = '显微镜下的细胞图像'
    const { bytes, report } = await compilePresentationDeck(deck, { trustedAssetEvidence: true })
    expect(report.assetWarnings).toEqual({ missingSource: 0, unknownLicense: 0, missingAltText: 0 })
    const zip = await JSZip.loadAsync(bytes)
    expect(await zip.file('ppt/slides/slide3.xml')!.async('string')).toContain('显微镜下的细胞图像')
    expect(await zip.file('ppt/notesSlides/notesSlide3.xml')!.async('string')).toContain(
      'license: licensed',
    )
    expect(await zip.file('ppt/notesSlides/notesSlide3.xml')!.async('string')).toContain(
      'asserted, not verified',
    )
  })

  it('rejects a forged inline evidence reference by default', async () => {
    const deck = benchmarkDeck()
    deck.assets[0]!.license = 'licensed'
    deck.assets[0]!.licenseEvidence = `attachment:${'a'.repeat(64)}`
    expect(() => parsePresentationDeck(deck)).toThrow(/untrusted_asset_license_evidence/)
    await expect(compilePresentationDeck(deck)).rejects.toThrow(/untrusted_asset_license_evidence/)
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
        d.assets[0].licenseEvidence = 'attachment:unknown'
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

  it('rejects a PNG with valid dimensions but corrupt pixel data', async () => {
    const deck = benchmarkDeck()
    const bytes = Buffer.from(deck.assets[0]!.base64, 'base64')
    bytes[48] = bytes[48]! ^ 0xff
    deck.assets[0]!.base64 = bytes.toString('base64')
    await expect(compilePresentationDeck(deck)).rejects.toThrow(/image_dimensions_or_data/)
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

it('rejects native table internal geometry, merge and border drift while preserving text and outer frame', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[5]!]
  const { bytes } = await compilePresentationDeck(deck)
  const originalZip = await JSZip.loadAsync(bytes)
  const original = await originalZip.file('ppt/slides/slide1.xml')!.async('string')
  const changes = [
    (xml: string) => xml.replace(/(<a:gridCol w=")\d+/, '$199999'),
    (xml: string) => xml.replace(/(<a:tr h=")\d+/, '$199999'),
    (xml: string) => xml.replace('<a:tc>', '<a:tc gridSpan="2">'),
    (xml: string) => xml.replace('<a:tc>', '<a:tc rowSpan="2">'),
    (xml: string) => xml.replace('<a:tc>', '<a:tc hMerge="1">'),
    (xml: string) => xml.replace('<a:tc>', '<a:tc vMerge="1">'),
    (xml: string) => xml.replace(/(<a:lnL w=")\d+/, '$199999'),
    (xml: string) => xml.replace(/(<a:lnR[^>]*><a:solidFill><a:srgbClr val=")[^"]+/, '$1112233'),
    (xml: string) => xml.replace(/(<a:lnT[^]*?<a:prstDash val=")solid/, '$1sysDash'),
    (xml: string) => xml.replace(/<a:lnB[^]*?<\/a:lnB>/, ''),
    ...['alpha', 'tint', 'shade'].map(
      (transform) => (xml: string) =>
        xml.replace(
          /(<a:lnL[^>]*><a:solidFill><a:srgbClr val="[^"]+")\/>/,
          `$1><a:${transform} val="0"/></a:srgbClr>`,
        ),
    ),
    (xml: string) => xml.replace(/(<a:lnL[^>]*cmpd=")sng/, '$1dbl'),
    (xml: string) => xml.replace(/(<a:lnL[^>]*cap=")flat/, '$1rnd'),
    (xml: string) => xml.replace(/(<a:lnL[^>]*algn=")ctr/, '$1in'),
    (xml: string) => xml.replace(/(<a:lnL[^>]*>)/, '$1<a:gradFill/>'),
  ]
  for (const change of changes) {
    const altered = change(original)
    expect(altered).not.toBe(original)
    // Table outer p:xfrm and every text node stay exactly as generated.
    expect(altered.match(/<p:xfrm>[^]*?<\/p:xfrm>/g)).toEqual(
      original.match(/<p:xfrm>[^]*?<\/p:xfrm>/g),
    )
    expect(altered.match(/<a:t>[^]*?<\/a:t>/g)).toEqual(original.match(/<a:t>[^]*?<\/a:t>/g))
    const zip = await JSZip.loadAsync(bytes)
    zip.file('ppt/slides/slide1.xml', altered)
    await expect(verifyCompiledPresentationStructure(zip, deck)).rejects.toThrow(
      'presentation_compile:structure_mismatch',
    )
  }
})
it('accepts generated fractional equal table dimensions with unchanged blank cells and no merges', async () => {
  const deck = benchmarkDeck()
  deck.style.accentColor = 'aBcDeF'
  deck.slides = [deck.slides[5]!]
  const table = deck.slides[0]!.elements.find((e) => e.kind === 'table')!
  if (table.kind !== 'table') throw Error('fixture')
  table.w = 9.1234567
  table.h = 2.2345678
  table.rows = [
    ['a', '', 'c'],
    ['', 'b', ''],
    ['x', 'y', 'z'],
  ]
  const { bytes } = await compilePresentationDeck(deck)
  const zip = await JSZip.loadAsync(bytes)
  await expect(verifyCompiledPresentationStructure(zip, deck)).resolves.toBeUndefined()
})
