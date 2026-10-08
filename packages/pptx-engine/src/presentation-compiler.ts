import PptxGenJS from 'pptxgenjs'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { XMLBuilder, XMLParser, XMLValidator } from 'fast-xml-parser'
import { relsPathFor, resolveTarget } from './zip'
import { parseChartXml } from './chart'
import {
  inspectPresentationGeometry,
  parsePresentationDeck,
  PRESENTATION_HEIGHT,
  PRESENTATION_WIDTH,
  presentationSlideSourceLabels,
  type PresentationInlineAsset,
  type PresentationCompileReport,
  type PresentationDeck,
} from './presentation'

type XmlNode = Record<string, any>
const xmlItems = (value: unknown): XmlNode[] =>
  value === undefined ? [] : Array.isArray(value) ? value : [value as XmlNode]

// Match only the bounded XML shape emitted for a generated solid table border.
function generatedXmlMatches(value: unknown, expected: unknown): boolean {
  if (!expected || typeof expected !== 'object') return value === expected
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  return (
    Object.keys(value).length === Object.keys(expected).length &&
    Object.entries(expected).every(([key, child]) =>
      generatedXmlMatches((value as XmlNode)[key], child),
    )
  )
}

function generatedChartFill(color: string): XmlNode {
  return { 'a:solidFill': { 'a:srgbClr': { '@_val': color.toUpperCase() } } }
}
function generatedChartLine(color: string, width: string): XmlNode {
  return {
    '@_w': width,
    '@_cap': 'flat',
    ...generatedChartFill(color),
    'a:prstDash': { '@_val': 'solid' },
    'a:round': '',
  }
}
function generatedChartText(
  fontFace: string,
  color: string,
  role: 'label' | 'axis' | 'legend',
  size = '1200',
): XmlNode {
  return {
    'a:bodyPr': '',
    'a:lstStyle': '',
    'a:p': {
      'a:pPr': {
        'a:defRPr': {
          ...(role === 'legend'
            ? {}
            : { '@_b': '0', '@_i': '0', '@_strike': 'noStrike', '@_sz': size, '@_u': 'none' }),
          ...generatedChartFill(color),
          'a:latin': { '@_typeface': fontFace },
          ...(role === 'legend' ? { 'a:cs': { '@_typeface': fontFace } } : {}),
        },
      },
      ...(role === 'label' ? {} : { 'a:endParaRPr': { '@_lang': 'en-US' } }),
    },
  }
}

// PptxGenJS hardcodes these pie defaults despite its dataLabel/dataBorder options.
// Repair only fresh generated parts with this exact known serializer shape.
async function normalizeGeneratedPieStyle(zip: JSZip, deck: PresentationDeck): Promise<boolean> {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    parseTagValue: false,
  })
  let changed = false
  for (const path of Object.keys(zip.files).filter((path) =>
    /^ppt\/charts\/chart\d+\.xml$/.test(path),
  )) {
    const xml = await zip.file(path)!.async('string')
    if (!xml.includes('<c:pieChart>')) continue
    const root = parser.parse(xml),
      pie = root['c:chartSpace']?.['c:chart']?.['c:plotArea']?.['c:pieChart']
    const series = xmlItems(pie?.['c:ser'])
    if (series.length !== 1) throw new Error('presentation_compile:structure_mismatch')
    const item = series[0]!,
      label = item['c:dLbls']
    if (
      !generatedXmlMatches(item['c:spPr'], {
        'a:solidFill': { 'a:schemeClr': { '@_val': 'accent1' } },
        'a:ln': generatedChartLine('F9F9F9', '9525'),
        'a:effectLst': '',
      }) ||
      !generatedXmlMatches(
        label?.['c:txPr'],
        generatedChartText('Arial', '000000', 'label', '1800'),
      )
    )
      throw new Error('presentation_compile:structure_mismatch')
    item['c:spPr'] = {
      ...generatedChartFill(deck.style.accentColor),
      'a:ln': generatedChartLine(deck.style.accentColor, '9525'),
      'a:effectLst': '',
    }
    label['c:txPr'] = generatedChartText(deck.style.fontFace, deck.style.textColor, 'label', '1800')
    zip.file(
      path,
      new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_' }).build(root),
    )
    changed = true
  }
  return changed
}

// Inspect only visible roles produced by our supported chart options. In particular,
// pie points explicitly override the serializer's unused parent theme fill/border.
function verifyGeneratedChartStyle(
  root: XmlNode,
  deck: PresentationDeck,
  type: 'bar' | 'line' | 'pie',
  categoryCount: number,
): void {
  const reject = () => {
    throw new Error('presentation_compile:structure_mismatch')
  }
  const normalize = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(normalize)
    if (!value || typeof value !== 'object') return value
    return Object.fromEntries(
      Object.entries(value)
        .filter(([key, child]) => !(key === '#text' && typeof child === 'string' && !child.trim()))
        .map(([key, child]) => [
          key,
          key === 'a:srgbClr' && child && typeof child === 'object'
            ? {
                ...child,
                '@_val':
                  typeof (child as XmlNode)['@_val'] === 'string'
                    ? (child as XmlNode)['@_val'].toUpperCase()
                    : (child as XmlNode)['@_val'],
              }
            : normalize(child),
        ]),
    )
  }
  const match = (value: unknown, expected: unknown) => {
    if (!generatedXmlMatches(normalize(value), expected)) reject()
  }
  const only = (value: XmlNode, keys: string[]) => {
    if (
      !value ||
      typeof value !== 'object' ||
      Object.keys(normalize(value) as XmlNode).some((key) => !keys.includes(key))
    )
      reject()
  }
  const fill = generatedChartFill,
    line = generatedChartLine
  const space = root['c:chartSpace'],
    chart = space?.['c:chart'],
    plot = chart?.['c:plotArea']
  const area = { ...fill(deck.style.background), 'a:ln': { 'a:noFill': '' }, 'a:effectLst': '' }
  match(space?.['c:spPr'], area)
  match(plot?.['c:spPr'], area)
  if (
    ['c:title', 'c:view3D', 'c:backWall', 'c:sideWall', 'c:floor'].some(
      (key) => chart?.[key] !== undefined,
    )
  )
    reject()
  const text = (role: 'label' | 'axis' | 'legend', size?: string) =>
    generatedChartText(deck.style.fontFace, deck.style.textColor, role, size)
  const legend = chart?.['c:legend']
  if (legend !== undefined)
    match(legend, {
      'c:legendPos': { '@_val': 'r' },
      'c:overlay': { '@_val': '0' },
      'c:txPr': text('legend'),
    })
  const plots = xmlItems(plot?.[`c:${type}Chart`]),
    series = plots.flatMap((p) => xmlItems(p['c:ser']))
  only(plot, ['c:layout', `c:${type}Chart`, 'c:catAx', 'c:valAx', 'c:spPr'])
  if (plots.length !== 1) reject()
  for (const value of plots) {
    only(
      value,
      type === 'bar'
        ? [
            'c:barDir',
            'c:grouping',
            'c:varyColors',
            'c:ser',
            'c:dLbls',
            'c:gapWidth',
            'c:overlap',
            'c:axId',
          ]
        : type === 'line'
          ? ['c:varyColors', 'c:ser', 'c:dLbls', 'c:marker', 'c:axId']
          : ['c:varyColors', 'c:ser', 'c:firstSliceAng'],
    )
    match(value['c:varyColors'], { '@_val': type === 'pie' ? '1' : '0' })
    if (type === 'bar') {
      match(value['c:barDir'], { '@_val': 'col' })
      match(value['c:grouping'], { '@_val': 'clustered' })
      match(value['c:gapWidth'], { '@_val': '150' })
      match(value['c:overlap'], { '@_val': '0' })
    } else if (type === 'line') match(value['c:marker'], { '@_val': '1' })
    else match(value['c:firstSliceAng'], { '@_val': '0' })
  }
  const labels = [plots[0]?.['c:dLbls'], ...series.map((s) => s['c:dLbls'])].filter(
    (v) => v !== undefined,
  )
  for (const label of labels) {
    if (label['c:spPr'] !== undefined || label['c:tx'] !== undefined) reject()
    match(label['c:numFmt'], { '@_formatCode': '#,##0.########', '@_sourceLinked': '0' })
    if (type === 'pie') {
      match(label['c:txPr'], text('label', '1800'))
      const points = xmlItems(label['c:dLbl'])
      if (points.length !== categoryCount) reject()
      points.forEach((point, index) => {
        if (point['c:idx']?.['@_val'] !== String(index) || point['c:tx'] !== undefined) reject()
        match(point['c:spPr'], '')
        match(point['c:txPr'], text('label'))
      })
    } else {
      if (label['c:dLbl'] !== undefined) reject()
      match(label['c:txPr'], text('label'))
    }
  }
  for (const [seriesIndex, item] of series.entries()) {
    const seriesColor = seriesIndex % 2 === 0 ? deck.style.accentColor : deck.style.textColor
    only(item, [
      'c:idx',
      'c:order',
      'c:tx',
      'c:spPr',
      'c:dLbls',
      'c:cat',
      'c:val',
      ...(type === 'pie' ? ['c:dPt'] : ['c:invertIfNegative']),
      ...(type === 'line' ? ['c:marker', 'c:smooth'] : []),
    ])
    if (type !== 'pie') match(item['c:invertIfNegative'], { '@_val': '0' })
    if (type === 'line') match(item['c:smooth'], { '@_val': '0' })
    if (type === 'pie') {
      match(item['c:spPr'], {
        ...fill(deck.style.accentColor),
        'a:ln': line(deck.style.accentColor, '9525'),
        'a:effectLst': '',
      })
      const points = xmlItems(item['c:dPt'])
      if (points.length !== categoryCount) reject()
      points.forEach((point, index) => {
        if (point['c:idx']?.['@_val'] !== String(index)) reject()
        match(point['c:spPr'], {
          ...fill(deck.style.accentColor),
          'a:ln': line(deck.style.accentColor, '9525'),
          'a:effectLst': '',
        })
      })
    } else {
      if (item['c:dPt'] !== undefined) reject()
      match(item['c:spPr'], {
        ...fill(seriesColor),
        ...(type === 'line' ? { 'a:ln': line(seriesColor, '25400') } : {}),
        'a:effectLst': '',
      })
      if (type === 'line')
        match(item['c:marker'], {
          'c:symbol': { '@_val': 'circle' },
          'c:size': { '@_val': '6' },
          'c:spPr': {
            ...fill(seriesColor),
            'a:ln': line(seriesColor, '9525'),
            'a:effectLst': '',
          },
        })
      else if (item['c:marker'] !== undefined) reject()
    }
  }
  const categories = xmlItems(plot?.['c:catAx']),
    values = xmlItems(plot?.['c:valAx'])
  if (categories.length !== (type === 'pie' ? 0 : 1) || values.length !== (type === 'pie' ? 0 : 1))
    reject()
  for (const [axis, valueAxis] of [
    ...categories.map((axis) => [axis, false] as const),
    ...values.map((axis) => [axis, true] as const),
  ]) {
    if (axis['c:title'] !== undefined || axis['c:minorGridlines'] !== undefined) reject()
    match(axis['c:txPr'], text('axis'))
    match(axis['c:spPr'], { 'a:ln': line(deck.style.textColor, '12700') })
    if (valueAxis)
      match(axis['c:majorGridlines'], { 'c:spPr': { 'a:ln': line(deck.style.textColor, '12700') } })
    else if (axis['c:majorGridlines'] !== undefined) reject()
  }
}

function shapeIds(root: unknown): string[] {
  if (Array.isArray(root)) return root.flatMap(shapeIds)
  if (!root || typeof root !== 'object') return []
  return Object.entries(root).flatMap(([key, value]) =>
    key === 'p:cNvPr' ? xmlItems(value).map((node) => node['@_id']) : shapeIds(value),
  )
}
function unsupportedShapeReferences(root: unknown): boolean {
  if (Array.isArray(root)) return root.some(unsupportedShapeReferences)
  if (!root || typeof root !== 'object') return false
  return Object.entries(root).some(
    ([key, value]) =>
      ['p:timing', 'p:bldLst', 'p:cxnSp', 'a:stCxn', 'a:endCxn'].includes(key) ||
      /^@_(?:spid|spId|shapeId)$/.test(key) ||
      unsupportedShapeReferences(value),
  )
}
/** Repair only freshly generated slides; no supported IR element carries a shape-ID reference. */
async function normalizeGeneratedShapeIds(zip: JSZip, slideCount: number): Promise<boolean> {
  const parser = new XMLParser({ ignoreAttributes: false, parseAttributeValue: false })
  let changed = false
  for (let index = 0; index < slideCount; index++) {
    const path = `ppt/slides/slide${index + 1}.xml`,
      file = zip.file(path)
    if (!file) throw Error('presentation_compile:structure_mismatch')
    const xml = await file.async('string')
    if (XMLValidator.validate(xml) !== true || unsupportedShapeReferences(parser.parse(xml)))
      throw Error('presentation_compile:structure_mismatch')
    let next = 0
    const normalized = xml.replace(/<p:cNvPr\b[^>]*>/g, (tag) => {
      next++
      if (next > 0xffffffff || !/\bid="[^"]*"/.test(tag))
        throw Error('presentation_compile:structure_mismatch')
      return tag.replace(/\bid="[^"]*"/, `id="${next}"`)
    })
    if (!next) throw Error('presentation_compile:structure_mismatch')
    if (normalized !== xml) {
      zip.file(path, normalized)
      changed = true
    }
  }
  return changed
}

function xmlText(value: unknown): string {
  if (typeof value === 'string') return value
  if (Array.isArray(value)) return value.map(xmlText).join('')
  if (!value || typeof value !== 'object') return ''
  if ('a:p' in value)
    return xmlItems((value as XmlNode)['a:p'])
      .map(xmlText)
      .join('\n')
  return Object.entries(value)
    .map(([key, child]) =>
      key === 'a:t' ? xmlText(child) : key.startsWith('@_') ? '' : xmlText(child),
    )
    .join('')
}

const colorEquals = (node: XmlNode | undefined, expected: string): boolean => {
  const value = node?.['a:srgbClr']?.['@_val']
  return typeof value === 'string' && value.toUpperCase() === expected.toUpperCase()
}

function verifyTextStyle(
  body: XmlNode | undefined,
  fontFace: string,
  fontSize: number,
  color: string,
  bold?: boolean,
  align?: 'left' | 'center' | 'right',
): void {
  const paragraphs = xmlItems(body?.['a:p'])
  const runs = paragraphs.flatMap((paragraph) => xmlItems(paragraph['a:r']))
  if (!runs.length) {
    if (
      paragraphs.length !== 1 ||
      xmlText(body) !== '' ||
      paragraphs.some((paragraph) => {
        const style = paragraph['a:endParaRPr'] as XmlNode | undefined
        return (
          style?.['a:latin']?.['@_typeface'] !== fontFace ||
          style?.['a:ea']?.['@_typeface'] !== fontFace ||
          Number(style?.['@_sz']) !== fontSize * 100
        )
      })
    )
      throw new Error('presentation_compile:structure_mismatch')
    return
  }
  if (
    runs.some((run) => {
      const style = run['a:rPr'] as XmlNode | undefined
      return (
        style?.['a:latin']?.['@_typeface'] !== fontFace ||
        style?.['a:ea']?.['@_typeface'] !== fontFace ||
        Number(style?.['@_sz']) !== fontSize * 100 ||
        !colorEquals(style?.['a:solidFill'], color) ||
        Boolean(Number(style?.['@_b'] ?? 0)) !== Boolean(bold)
      )
    })
  )
    throw new Error('presentation_compile:structure_mismatch')
  if (
    align &&
    paragraphs.some(
      (paragraph) =>
        paragraph['a:pPr']?.['@_algn'] !== { left: 'l', center: 'ctr', right: 'r' }[align],
    )
  )
    throw new Error('presentation_compile:structure_mismatch')
}

/** The visible chart cache and the workbook opened by PowerPoint's Edit Data must agree. */
async function chartWorkbookPath(
  zip: JSZip,
  chartPath: string,
  chartRoot: XmlNode,
  parser: XMLParser,
): Promise<string> {
  const fail = (): never => {
    throw new Error('presentation_compile:structure_mismatch')
  }
  const external = chartRoot['c:chartSpace']?.['c:externalData']
  const id = external?.['@_r:id']
  if (typeof id !== 'string') fail()
  const relsFile = zip.file(relsPathFor(chartPath))
  if (!relsFile) fail()
  const relsXml = await relsFile!.async('string')
  if (relsXml.length > 1024 * 1024 || XMLValidator.validate(relsXml) !== true) fail()
  const relations = xmlItems((parser.parse(relsXml) as XmlNode).Relationships?.Relationship)
  const matches = relations.filter((relation) => relation['@_Id'] === id)
  if (
    matches.length !== 1 ||
    !String(matches[0]?.['@_Type']).endsWith('/package') ||
    matches[0]?.['@_TargetMode'] !== undefined ||
    typeof matches[0]?.['@_Target'] !== 'string'
  )
    fail()
  return resolveTarget(chartPath, matches[0]!['@_Target'])
}

async function verifyChartWorkbook(
  zip: JSZip,
  chartPath: string,
  chartRoot: XmlNode,
  chart: Extract<PresentationDeck['slides'][number]['elements'][number], { kind: 'chart' }>,
  seriesNodes: XmlNode[],
  parser: XMLParser,
): Promise<void> {
  const fail = (): never => {
    throw new Error('presentation_compile:structure_mismatch')
  }
  const workbookPath = await chartWorkbookPath(zip, chartPath, chartRoot, parser)
  const workbookFile = zip.file(workbookPath)
  if (!workbookFile) fail()
  const workbookBytes = await workbookFile!.async('uint8array')
  if (workbookBytes.length > 8 * 1024 * 1024) fail()
  let workbook: JSZip
  try {
    workbook = await JSZip.loadAsync(workbookBytes)
  } catch {
    return fail()
  }
  if (Object.keys(workbook.files).length > 100) fail()
  const readXml = async (path: string): Promise<XmlNode> => {
    const file = workbook.file(path)
    if (!file) fail()
    const xml = await file!.async('string')
    if (xml.length > 2 * 1024 * 1024 || XMLValidator.validate(xml) !== true) fail()
    return parser.parse(xml) as XmlNode
  }
  const strings = xmlItems((await readXml('xl/sharedStrings.xml')).sst?.si).map((item) => {
    if (item.r !== undefined) fail()
    if (typeof item.t === 'string') return item.t
    if (
      item.t &&
      typeof item.t === 'object' &&
      Object.keys(item.t).every((key) => key === '#text' || key === '@_xml:space') &&
      (item.t['#text'] === undefined || typeof item.t['#text'] === 'string')
    )
      return item.t['#text'] ?? ''
    return fail()
  })
  const expected = new Map<string, string | number>([['A1', '']])
  if (seriesNodes.length !== chart.series.length) fail()
  chart.series.forEach((series, index) => {
    const column = String.fromCharCode(66 + index)
    const node = seriesNodes[index]!
    const lastRow = chart.categories.length + 1
    if (
      node['c:tx']?.['c:strRef']?.['c:f'] !== `Sheet1!$${column}$1` ||
      (node['c:cat']?.['c:multiLvlStrRef']?.['c:f'] ??
        node['c:cat']?.['c:strRef']?.['c:f'] ??
        node['c:cat']?.['c:numRef']?.['c:f']) !== `Sheet1!$A$2:$A$${lastRow}` ||
      node['c:val']?.['c:numRef']?.['c:f'] !== `Sheet1!$${column}$2:$${column}$${lastRow}`
    )
      fail()
    expected.set(`${column}1`, series.name)
    series.values.forEach((value, row) => expected.set(`${column}${row + 2}`, value))
  })
  chart.categories.forEach((category, row) => expected.set(`A${row + 2}`, category))
  const rows = xmlItems((await readXml('xl/worksheets/sheet1.xml')).worksheet?.sheetData?.row)
  const cells = rows.flatMap((row) => xmlItems(row.c))
  if (cells.length !== expected.size) fail()
  const seen = new Set<string>()
  for (const cell of cells) {
    const address = cell['@_r']
    const raw = cell.v
    if (typeof address !== 'string' || seen.has(address) || !expected.has(address) || cell.f) fail()
    seen.add(address)
    const want = expected.get(address)
    if (
      typeof raw !== 'string' ||
      (cell['@_t'] === 's'
        ? !/^\d+$/.test(raw)
        : !/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw))
    )
      fail()
    const actual =
      cell['@_t'] === 's'
        ? strings[Number(raw)]
        : cell['@_t'] === undefined || cell['@_t'] === 'n'
          ? Number(raw)
          : undefined
    if (actual !== want) fail()
  }
}

/** PptxGenJS emits an empty worksheet cell for numeric zero; restore the editable value. */
async function restoreGeneratedChartZeros(zip: JSZip, deck: PresentationDeck): Promise<boolean> {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' })
  const chartPaths = Object.keys(zip.files)
    .filter((path) => /^ppt\/charts\/chart\d+\.xml$/.test(path))
    .sort(
      (left, right) =>
        Number(/\d+(?=\.xml$)/.exec(left)![0]) - Number(/\d+(?=\.xml$)/.exec(right)![0]),
    )
  let chartIndex = 0
  let changed = false
  for (const slide of deck.slides)
    for (const element of slide.elements) {
      if (element.kind !== 'chart') continue
      const chartPath = chartPaths[chartIndex++]
      if (!chartPath) throw new Error('presentation_compile:structure_mismatch')
      const zeroCells = element.series.flatMap((series, seriesIndex) =>
        series.values.flatMap((value, rowIndex) =>
          value === 0 ? [`${String.fromCharCode(66 + seriesIndex)}${rowIndex + 2}`] : [],
        ),
      )
      if (!zeroCells.length) continue
      const chartXml = await zip.file(chartPath)!.async('string')
      if (chartXml.length > 2 * 1024 * 1024 || XMLValidator.validate(chartXml) !== true)
        throw new Error('presentation_compile:structure_mismatch')
      const path = await chartWorkbookPath(
        zip,
        chartPath,
        parser.parse(chartXml) as XmlNode,
        parser,
      )
      const file = zip.file(path)
      if (!file) throw new Error('presentation_compile:structure_mismatch')
      const workbook = await JSZip.loadAsync(await file.async('uint8array'))
      const sheetPath = 'xl/worksheets/sheet1.xml'
      const sheet = workbook.file(sheetPath)
      if (!sheet) throw new Error('presentation_compile:structure_mismatch')
      let xml = await sheet.async('string')
      for (const address of zeroCells) {
        const before = `<c r="${address}"><v></v></c>`
        if (xml.includes(before)) {
          xml = xml.replace(before, `<c r="${address}"><v>0</v></c>`)
          changed = true
        }
      }
      workbook.file(sheetPath, xml)
      zip.file(path, await workbook.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }))
    }
  return changed
}

/** Verify the generated OOXML has the promised native objects before reporting structure passed. */
export async function verifyCompiledPresentationStructure(
  zip: JSZip,
  deck: PresentationDeck,
): Promise<void> {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    parseTagValue: false,
    trimValues: false,
  })
  const verifiedImages = new Set<string>()
  for (const [index, ir] of deck.slides.entries()) {
    const slidePath = `ppt/slides/slide${index + 1}.xml`
    const file = zip.file(slidePath)
    if (!file) throw new Error('presentation_compile:structure_mismatch')
    const xml = await file.async('string')
    if (
      new TextEncoder().encode(xml).byteLength > 8 * 1024 * 1024 ||
      XMLValidator.validate(xml) !== true
    )
      throw new Error('presentation_compile:structure_mismatch')
    const root = parser.parse(xml) as XmlNode
    const nativeIds = shapeIds(root)
    if (
      unsupportedShapeReferences(root) ||
      !nativeIds.length ||
      nativeIds.some(
        (id) => typeof id !== 'string' || !/^[1-9]\d*$/.test(id) || Number(id) > 0xffffffff,
      ) ||
      new Set(nativeIds).size !== nativeIds.length
    )
      throw Error('presentation_compile:structure_mismatch')
    if (
      !colorEquals(
        root['p:sld']?.['p:cSld']?.['p:bg']?.['p:bgPr']?.['a:solidFill'],
        deck.style.background,
      )
    )
      throw new Error('presentation_compile:structure_mismatch')
    const tree = root['p:sld']?.['p:cSld']?.['p:spTree'] as XmlNode | undefined
    if (!tree || tree['p:grpSp'] || tree['p:cxnSp'])
      throw new Error('presentation_compile:structure_mismatch')
    const expected = new Map(ir.elements.map((element) => [element.id, element]))
    const sourceLabels = presentationSlideSourceLabels(ir, deck.claims)
    if (sourceLabels.length)
      expected.set('source-attribution', {
        kind: 'text',
        id: 'source-attribution',
        x: 0.5,
        y: 7.05,
        w: 12.3,
        h: 0.3,
        text: sourceLabels.join('；'),
      })
    const seen = new Set<string>()
    let relationships:
      Map<string, { type: string; target: string; targetMode?: string }> | undefined
    const linkedPart = async (id: string, kind: 'image' | 'chart') => {
      if (!relationships) {
        const rels = zip.file(relsPathFor(slidePath))
        if (!rels) throw new Error('presentation_compile:structure_mismatch')
        const relsXml = await rels.async('string')
        if (relsXml.length > 1024 * 1024 || XMLValidator.validate(relsXml) !== true)
          throw new Error('presentation_compile:structure_mismatch')
        const nodes = xmlItems((parser.parse(relsXml) as XmlNode).Relationships?.Relationship)
        relationships = new Map(
          nodes.map((node) => [
            node['@_Id'],
            {
              type: node['@_Type'],
              target: node['@_Target'],
              targetMode: node['@_TargetMode'],
            },
          ]),
        )
        if (relationships.size !== nodes.length)
          throw new Error('presentation_compile:structure_mismatch')
      }
      const relation = relationships.get(id)
      if (
        !relation ||
        typeof relation.type !== 'string' ||
        !relation.type.endsWith(`/${kind}`) ||
        typeof relation.target !== 'string' ||
        relation.targetMode !== undefined
      )
        throw new Error('presentation_compile:structure_mismatch')
      const target = resolveTarget(slidePath, relation.target)
      if (!zip.file(target)) throw new Error('presentation_compile:structure_mismatch')
      return target
    }
    for (const [tag, nv] of [
      ['p:sp', 'p:nvSpPr'],
      ['p:pic', 'p:nvPicPr'],
      ['p:graphicFrame', 'p:nvGraphicFramePr'],
    ] as const)
      for (const object of xmlItems(tree[tag])) {
        const name = object[nv]?.['p:cNvPr']?.['@_name']
        const element = expected.get(name)
        if (!element || seen.has(name)) throw new Error('presentation_compile:structure_mismatch')
        seen.add(name)
        const actualKind =
          tag === 'p:pic'
            ? 'image'
            : tag === 'p:sp'
              ? element.kind === 'text' && object['p:txBody']
                ? 'text'
                : 'shape'
              : object['a:graphic']?.['a:graphicData']?.['a:tbl']
                ? 'table'
                : object['a:graphic']?.['a:graphicData']?.['c:chart']
                  ? 'chart'
                  : 'unknown'
        if (actualKind !== element.kind) throw new Error('presentation_compile:structure_mismatch')
        {
          const transform =
            tag === 'p:graphicFrame' ? object['p:xfrm'] : object['p:spPr']?.['a:xfrm']
          const actual = [
            transform?.['a:off']?.['@_x'],
            transform?.['a:off']?.['@_y'],
            transform?.['a:ext']?.['@_cx'],
            transform?.['a:ext']?.['@_cy'],
          ].map(Number)
          let box = [element.x, element.y, element.w, element.h]
          if (element.kind === 'image' && element.fit !== 'cover') {
            const asset = deck.assets.find((item) => item.id === element.assetId)
            if (!asset || !('base64' in asset))
              throw new Error('presentation_compile:structure_mismatch')
            const scale = Math.min(element.w / asset.width, element.h / asset.height)
            const w = asset.width * scale
            const h = asset.height * scale
            box = [element.x + (element.w - w) / 2, element.y + (element.h - h) / 2, w, h]
          }
          const expectedBox = box.map((value) => value * 914400)
          if (
            actual.some(
              (value, coordinate) =>
                !Number.isFinite(value) || Math.abs(value - expectedBox[coordinate]!) > 19050,
            )
          )
            throw new Error('presentation_compile:structure_mismatch')
        }
        if (
          element.kind === 'shape' &&
          object['p:spPr']?.['a:prstGeom']?.['@_prst'] !== element.shape
        )
          throw new Error('presentation_compile:structure_mismatch')
        if (element.kind === 'text') {
          if (xmlText(object['p:txBody']) !== element.text)
            throw new Error('presentation_compile:structure_mismatch')
          verifyTextStyle(
            object['p:txBody'],
            deck.style.fontFace,
            element.id === 'source-attribution' ? 8 : (element.fontSize ?? 20),
            element.color ?? deck.style.textColor,
            element.bold ?? false,
            element.id === 'source-attribution' ? undefined : (element.align ?? 'left'),
          )
        }
        if (element.kind === 'shape') {
          const properties = object['p:spPr'] as XmlNode | undefined
          if (
            !colorEquals(properties?.['a:solidFill'], element.fill ?? deck.style.accentColor) ||
            !colorEquals(
              properties?.['a:ln']?.['a:solidFill'],
              element.lineColor ?? element.fill ?? deck.style.accentColor,
            )
          )
            throw new Error('presentation_compile:structure_mismatch')
        }
        if (element.kind === 'table') {
          const table = object['a:graphic']['a:graphicData']['a:tbl'] as XmlNode
          const columns = xmlItems(table['a:tblGrid']?.['a:gridCol'])
          const nativeRows = xmlItems(table['a:tr'])
          // PptxGenJS inch2Emu rounds each supplied column/row independently.
          const columnWidth = Math.round((element.w / element.rows[0]!.length) * 914400)
          const rowHeight = Math.round((element.h / element.rows.length) * 914400)
          if (
            columns.length !== element.rows[0]!.length ||
            columns.some((column) => Number(column['@_w']) !== columnWidth) ||
            nativeRows.length !== element.rows.length ||
            nativeRows.some((row) => Number(row['@_h']) !== rowHeight)
          )
            throw new Error('presentation_compile:structure_mismatch')
          const rows = nativeRows.map((row) =>
            xmlItems(row['a:tc']).map((cell) => {
              // SlideIR has no merge topology: generated cells must remain unmerged.
              if (
                ['gridSpan', 'rowSpan', 'hMerge', 'vMerge'].some((key) =>
                  Object.hasOwn(cell, `@_${key}`),
                ) ||
                ['a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'].some((side) => {
                  const border = cell['a:tcPr']?.[side] as XmlNode | undefined
                  return !generatedXmlMatches(border, {
                    '@_w': '12700',
                    '@_cap': 'flat',
                    '@_cmpd': 'sng',
                    '@_algn': 'ctr',
                    'a:solidFill': {
                      'a:srgbClr': { '@_val': deck.style.accentColor.toUpperCase() },
                    },
                    'a:prstDash': { '@_val': 'solid' },
                    'a:round': '',
                    'a:headEnd': { '@_type': 'none', '@_w': 'med', '@_len': 'med' },
                    'a:tailEnd': { '@_type': 'none', '@_w': 'med', '@_len': 'med' },
                  })
                })
              )
                throw new Error('presentation_compile:structure_mismatch')
              verifyTextStyle(
                cell['a:txBody'],
                deck.style.fontFace,
                element.fontSize ?? 16,
                deck.style.textColor,
              )
              return xmlText(cell['a:txBody'])
            }),
          )
          if (JSON.stringify(rows) !== JSON.stringify(element.rows))
            throw new Error('presentation_compile:structure_mismatch')
        }
        if (element.kind === 'image') {
          if (
            object[nv]?.['p:cNvPr']?.['@_descr'] !==
            (element.altText ?? 'Image description missing')
          )
            throw new Error('presentation_compile:structure_mismatch')
          const asset = deck.assets.find((item) => item.id === element.assetId)
          if (!asset || !('base64' in asset))
            throw new Error('presentation_compile:structure_mismatch')
          const crop = object['p:blipFill']?.['a:srcRect'] as XmlNode | undefined
          if (element.fit === 'cover') {
            if (!crop) throw new Error('presentation_compile:structure_mismatch')
            const sourceAspect = asset.width / asset.height
            const targetAspect = element.w / element.h
            const expectedCrop =
              sourceAspect > targetAspect
                ? {
                    l: (1 - targetAspect / sourceAspect) * 50000,
                    r: (1 - targetAspect / sourceAspect) * 50000,
                    t: 0,
                    b: 0,
                  }
                : {
                    l: 0,
                    r: 0,
                    t: (1 - sourceAspect / targetAspect) * 50000,
                    b: (1 - sourceAspect / targetAspect) * 50000,
                  }
            if (
              Object.entries(expectedCrop).some(([side, expected]) => {
                const actual = Number(crop[`@_${side}`] ?? 0)
                return !Number.isFinite(actual) || Math.abs(actual - expected) > 100
              })
            )
              throw new Error('presentation_compile:structure_mismatch')
          } else if (
            crop &&
            ['l', 'r', 't', 'b'].some((side) => Number(crop[`@_${side}`] ?? 0) !== 0)
          )
            throw new Error('presentation_compile:structure_mismatch')
          const id = object['p:blipFill']?.['a:blip']?.['@_r:embed']
          if (typeof id !== 'string') throw new Error('presentation_compile:structure_mismatch')
          const path = await linkedPart(id, 'image')
          const key = `${asset.id}/${path}`
          if (!verifiedImages.has(key)) {
            if (
              !Buffer.from(await zip.file(path)!.async('uint8array')).equals(
                Buffer.from(asset.base64, 'base64'),
              )
            )
              throw new Error('presentation_compile:structure_mismatch')
            verifiedImages.add(key)
          }
        }
        if (element.kind === 'chart') {
          const id = object['a:graphic']['a:graphicData']['c:chart']?.['@_r:id']
          if (typeof id !== 'string') throw new Error('presentation_compile:structure_mismatch')
          const chartPath = await linkedPart(id, 'chart')
          const chartXml = await zip.file(chartPath)!.async('string')
          if (chartXml.length > 2 * 1024 * 1024 || XMLValidator.validate(chartXml) !== true)
            throw new Error('presentation_compile:structure_mismatch')
          const chart = parseChartXml(chartXml)
          const chartRoot = parser.parse(chartXml) as XmlNode
          const plotArea = chartRoot['c:chartSpace']?.['c:chart']?.['c:plotArea'] as
            XmlNode | undefined
          const plotTag = `${element.chartType === 'bar' ? 'bar' : element.chartType}Chart`
          const plots = xmlItems(plotArea?.[`c:${plotTag}`])
          const seriesNodes = plots.flatMap((plot) => xmlItems(plot['c:ser']))
          const chartBody = chartRoot['c:chartSpace']?.['c:chart'] as XmlNode | undefined
          const labels = [
            plots[0]?.['c:dLbls'],
            ...seriesNodes.map((series) => series['c:dLbls']),
          ].filter((value) => value !== undefined)
          const labelsMatch =
            element.chartType === 'pie'
              ? labels.length === 1 &&
                xmlItems(labels[0]?.['c:dLbl']).length === element.categories.length &&
                xmlItems(labels[0]?.['c:dLbl']).every(
                  (label) => Number(label['c:showVal']?.['@_val']) === 1,
                ) &&
                Number(labels[0]?.['c:showCatName']?.['@_val']) === 1 &&
                Number(labels[0]?.['c:showPercent']?.['@_val']) === 1
              : labels.length > 0 &&
                labels.every((value) => Number(value?.['c:showVal']?.['@_val']) === 1)
          if (
            (chartBody?.['c:legend'] !== undefined) !== element.series.length > 1 ||
            plots.length !== 1 ||
            !labelsMatch
          )
            throw new Error('presentation_compile:structure_mismatch')
          verifyGeneratedChartStyle(chartRoot, deck, element.chartType, element.categories.length)
          await verifyChartWorkbook(zip, chartPath, chartRoot, element, seriesNodes, parser)
          const categoriesMatch =
            seriesNodes.length === element.series.length &&
            seriesNodes.every((series) => {
              const cat = series['c:cat'] as XmlNode | undefined
              const cache =
                cat?.['c:multiLvlStrRef']?.['c:multiLvlStrCache'] ??
                cat?.['c:strRef']?.['c:strCache'] ??
                cat?.['c:numRef']?.['c:numCache']
              const points = xmlItems(cache?.['c:lvl']?.['c:pt'] ?? cache?.['c:pt'])
              const actual = points.map((point) => String(point['c:v'] ?? ''))
              return JSON.stringify(actual) === JSON.stringify(element.categories)
            })
          if (
            chart?.kind !== element.chartType ||
            !categoriesMatch ||
            JSON.stringify(chart.categories) !== JSON.stringify(element.categories) ||
            JSON.stringify(
              chart.series.map((series) => ({ name: series.name, values: series.values })),
            ) !== JSON.stringify(element.series)
          )
            throw new Error('presentation_compile:structure_mismatch')
        }
      }
    if (seen.size !== expected.size) throw new Error('presentation_compile:structure_mismatch')
  }
}

/** Validate encoded raster dimensions before passing bytes to the PPTX writer. No external I/O. */
function imageData(asset: PresentationInlineAsset): string {
  const bytes = Buffer.from(asset.base64, 'base64')
  if (bytes.toString('base64') !== asset.base64)
    throw new Error('presentation_invalid:image_encoding')
  let width = 0,
    height = 0
  if (asset.mime === 'image/png') {
    if (
      bytes.length >= 33 &&
      bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10])) &&
      bytes.toString('ascii', 12, 16) === 'IHDR' &&
      bytes.toString('ascii', bytes.length - 8, bytes.length - 4) === 'IEND'
    ) {
      width = bytes.readUInt32BE(16)
      height = bytes.readUInt32BE(20)
      if (width === asset.width && height === asset.height) {
        try {
          const decoded = PNG.sync.read(bytes)
          if (decoded.width !== width || decoded.height !== height)
            throw new Error('decoded_dimensions_mismatch')
        } catch {
          throw new Error('presentation_invalid:image_dimensions_or_data')
        }
      }
    }
  } else {
    // JPEG marker segments carry dimensions in SOF; never trust caller-supplied dimensions.
    let offset = 2
    while (offset + 4 <= bytes.length && bytes[offset] === 0xff) {
      const marker = bytes[offset + 1]!
      if (marker === 0xda || marker === 0xd9) break
      if (marker === 0xff) {
        offset++
        continue
      }
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        offset += 2
        continue
      }
      const length = bytes.readUInt16BE(offset + 2)
      if (length < 2 || offset + 2 + length > bytes.length) break
      if ([0xc0, 0xc1, 0xc2].includes(marker) && length >= 8) {
        height = bytes.readUInt16BE(offset + 5)
        width = bytes.readUInt16BE(offset + 7)
        break
      }
      offset += length + 2
    }
  }
  if (width !== asset.width || height !== asset.height || !width || !height)
    throw new Error('presentation_invalid:image_dimensions_or_data')
  return `data:${asset.mime};base64,${asset.base64}`
}

/** Deterministic mapping from validated IR to editable OOXML; this is not a rendered visual review. */
export async function compilePresentationDeck(
  input: unknown,
  options: { trustedAssetEvidence?: boolean; fontAvailable?: (family: string) => boolean } = {},
): Promise<{ bytes: Uint8Array; report: PresentationCompileReport; sourceSlideIds?: string[] }> {
  const requestedDeck = parsePresentationDeck(input, options)
  let fontResolution: PresentationCompileReport['fontResolution']
  let usedFont = requestedDeck.style.fontFace
  if (requestedDeck.style.fontFallbacks && options.fontAvailable) {
    usedFont = [usedFont, ...requestedDeck.style.fontFallbacks].find(options.fontAvailable) ?? ''
    if (!usedFont) throw new Error('font_unavailable')
    fontResolution = {
      requested: requestedDeck.style.fontFace,
      used: usedFont,
      substituted: usedFont !== requestedDeck.style.fontFace,
    }
  }
  const deck =
    usedFont === requestedDeck.style.fontFace
      ? requestedDeck
      : { ...requestedDeck, style: { ...requestedDeck.style, fontFace: usedFont } }
  const geometry = inspectPresentationGeometry(deck)
  if (geometry.some((issue) => issue.kind === 'out_of_bounds'))
    throw new Error('presentation_geometry:out_of_bounds')
  const assets = new Map(
    deck.assets.map((asset) => {
      if ('attachmentId' in asset) throw new Error('presentation_invalid:unresolved_asset')
      return [asset.id, { ...asset, data: imageData(asset) }] as const
    }),
  )
  const assetWarnings = { missingSource: 0, unknownLicense: 0, missingAltText: 0 }
  const pptx = new PptxGenJS()
  pptx.defineLayout({
    name: 'WISWORK_16_9',
    width: PRESENTATION_WIDTH,
    height: PRESENTATION_HEIGHT,
  })
  pptx.layout = 'WISWORK_16_9'
  pptx.author = 'WisWork'
  pptx.subject = deck.id
  pptx.title = deck.title
  pptx.theme = { headFontFace: deck.style.fontFace, bodyFontFace: deck.style.fontFace }
  for (const ir of deck.slides) {
    const sources = presentationSlideSourceLabels(ir, deck.claims)
    const sourceText = sources.join('；')
    if (sourceText.length > 500) throw new Error('presentation_invalid:source_footer_overflow')
    const slide = pptx.addSlide()
    slide.background = { color: deck.style.background }
    for (const el of ir.elements) {
      const box = { x: el.x, y: el.y, w: el.w, h: el.h, objectName: el.id }
      if (el.kind === 'text')
        slide.addText(el.text, {
          ...box,
          fontFace: deck.style.fontFace,
          fontSize: el.fontSize ?? 20,
          color: el.color ?? deck.style.textColor,
          bold: el.bold ?? false,
          align: el.align ?? 'left',
          margin: 0,
          breakLine: false,
          valign: 'top',
        })
      else if (el.kind === 'shape')
        slide.addShape(pptx.ShapeType[el.shape], {
          ...box,
          fill: { color: el.fill ?? deck.style.accentColor },
          line: { color: el.lineColor ?? el.fill ?? deck.style.accentColor },
        })
      else if (el.kind === 'image') {
        const asset = assets.get(el.assetId)!
        if (!asset.source && !asset.sources?.length) assetWarnings.missingSource++
        if (!asset.license || asset.license === 'unknown') assetWarnings.unknownLicense++
        if (!el.altText) assetWarnings.missingAltText++
        // Contain uses verified dimensions. Cover uses native image crop, never rasterizes text.
        if (el.fit === 'cover') {
          // PptxGenJS derives cover's source aspect ratio from w/h, not encoded data.
          // Supply header-verified intrinsic dimensions; sizing sets the final target box.
          slide.addImage({
            ...box,
            data: asset.data,
            altText: el.altText ?? 'Image description missing',
            w: asset.width / 96,
            h: asset.height / 96,
            sizing: { type: 'cover', w: el.w, h: el.h },
          })
        } else {
          const scale = Math.min(el.w / asset.width, el.h / asset.height)
          const w = asset.width * scale,
            h = asset.height * scale
          slide.addImage({
            ...box,
            data: asset.data,
            altText: el.altText ?? 'Image description missing',
            x: el.x + (el.w - w) / 2,
            y: el.y + (el.h - h) / 2,
            w,
            h,
          })
        }
      } else if (el.kind === 'table')
        slide.addTable(
          el.rows.map((row) => row.map((text) => ({ text }))),
          {
            ...box,
            fontFace: deck.style.fontFace,
            fontSize: el.fontSize ?? 16,
            color: deck.style.textColor,
            border: { type: 'solid', color: deck.style.accentColor, pt: 1 },
            margin: 0.04,
            rowH: el.h / el.rows.length,
            colW: Array(el.rows[0]!.length).fill(el.w / el.rows[0]!.length),
            autoPage: false,
          },
        )
      else
        slide.addChart(
          pptx.ChartType[el.chartType],
          el.series.map((series) => ({
            name: series.name,
            labels: el.categories,
            values: series.values,
          })),
          {
            ...box,
            showLegend: el.series.length > 1,
            showTitle: false,
            chartColors:
              el.series.length > 1 && el.chartType !== 'pie'
                ? el.series.map((_, index) =>
                    index % 2 === 0 ? deck.style.accentColor : deck.style.textColor,
                  )
                : [deck.style.accentColor],
            ...(el.chartType === 'pie'
              ? { dataBorder: { color: deck.style.accentColor, pt: 0.75 } }
              : {}),
            showValue: true,
            dataLabelFormatCode: '#,##0.########',
            dataLabelFontFace: deck.style.fontFace,
            dataLabelColor: deck.style.textColor,
            legendColor: deck.style.textColor,
            catAxisLabelColor: deck.style.textColor,
            valAxisLabelColor: deck.style.textColor,
            catAxisLineColor: deck.style.textColor,
            valAxisLineColor: deck.style.textColor,
            catGridLine: { color: deck.style.textColor, style: 'none' },
            valGridLine: { color: deck.style.textColor, size: 1, style: 'solid', cap: 'flat' },
            chartArea: { fill: { color: deck.style.background } },
            plotArea: { fill: { color: deck.style.background } },
            ...(el.chartType === 'bar' ? { valAxisMinVal: 0 } : {}),
            catAxisLabelFontFace: deck.style.fontFace,
            valAxisLabelFontFace: deck.style.fontFace,
            legendFontFace: deck.style.fontFace,
          },
        )
    }
    const claims = (ir.claimIds ?? []).map((id) => deck.claims.find((claim) => claim.id === id)!)
    if (sourceText)
      slide.addText(sourceText, {
        x: 0.5,
        y: 7.05,
        w: 12.3,
        h: 0.3,
        fontFace: deck.style.fontFace,
        fontSize: 8,
        color: deck.style.textColor,
        margin: 0,
        valign: 'top',
        align: 'left',
        objectName: 'source-attribution',
      })
    const assetSources = ir.elements
      .filter((el) => el.kind === 'image')
      .map((el) => {
        const asset = assets.get(el.assetId)!
        const sources = asset.sources?.length
          ? asset.sources.join('; ')
          : (asset.source ?? 'source not supplied')
        return `Image [${asset.id}]: ${sources}; license: ${asset.license ?? 'unknown'}${asset.licenseEvidence ? ` (asserted, not verified; evidence: ${asset.licenseEvidence})` : ''}; alt text: ${el.altText ?? 'missing'}`
      })
    slide.addNotes(
      [
        ir.title,
        ir.notes ?? '',
        ...claims.map((claim, i) => `${sources[i]}\n${claim.text}`),
        ...assetSources,
        'Attribution supplied by the input; source accuracy and license rights have not been independently verified.',
      ]
        .filter(Boolean)
        .join('\n\n'),
    )
  }
  const output = await pptx.write({ outputType: 'uint8array', compression: true })
  if (!(output instanceof Uint8Array)) throw new Error('presentation_compile:unexpected_output')
  const zip = await JSZip.loadAsync(output)
  const normalizedIds = await normalizeGeneratedShapeIds(zip, deck.slides.length)
  const presentation = zip.file('ppt/presentation.xml')
  if (!presentation) throw new Error('presentation_compile:missing_presentation')
  const parsed = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    parseAttributeValue: false,
    isArray: (name) => name === 'p:sldId',
  }).parse(await presentation.async('string')) as {
    'p:presentation'?: { 'p:sldIdLst'?: { 'p:sldId'?: Array<{ '@_id'?: unknown }> } }
  }
  const slideIds = parsed['p:presentation']?.['p:sldIdLst']?.['p:sldId']
  if (!Array.isArray(slideIds) || slideIds.length !== deck.slides.length)
    throw new Error('presentation_compile:invalid_slide_ids')
  const sourceSlideIds = slideIds.map((slide) => {
    const id = slide['@_id']
    if (
      typeof id !== 'string' ||
      !/^[1-9]\d*$/.test(id) ||
      !Number.isSafeInteger(Number(id)) ||
      Number(id) < 256 ||
      Number(id) > 0xffffffff
    )
      throw new Error('presentation_compile:invalid_slide_ids')
    return `${id}#`
  })
  if (new Set(sourceSlideIds).size !== sourceSlideIds.length)
    throw new Error('presentation_compile:invalid_slide_ids')
  const normalizedPie = await normalizeGeneratedPieStyle(zip, deck)
  const repairedZeros = await restoreGeneratedChartZeros(zip, deck)
  await verifyCompiledPresentationStructure(zip, deck)
  return {
    bytes:
      repairedZeros || normalizedIds || normalizedPie
        ? await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
        : output,
    sourceSlideIds,
    report: {
      deckId: deck.id,
      slideCount: deck.slides.length,
      elementCount: deck.slides.reduce((n, slide) => n + slide.elements.length, 0),
      geometry,
      assetWarnings,
      ...(fontResolution ? { fontResolution } : {}),
      checks: {
        structure: 'passed',
        geometry: geometry.length ? 'warning' : 'passed',
        render: 'not_run',
        sources: 'not_verified',
        roundTrip: 'not_run',
      },
    },
  }
}
