import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { loadBoundedZip, MAX_PPTX_XML_BYTES } from './powerpoint-package.js'

type Node = Record<string, unknown>
const parser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  processEntities: false,
})

function elements(nodes: Node[], tag: string): Node[][] {
  const found: Node[][] = []
  for (const node of nodes) {
    for (const [key, value] of Object.entries(node)) {
      if (!Array.isArray(value)) continue
      if (key === tag) found.push(value as Node[])
      found.push(...elements(value as Node[], tag))
    }
  }
  return found
}

function value(nodes: Node[]): string {
  return elements(nodes, 'a:t')
    .map((parts) => parts.map((part) => part['#text'] ?? '').join(''))
    .join('')
}

function xmlNodes(xml: string): Node[] {
  if (
    new TextEncoder().encode(xml).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/i.test(xml) ||
    XMLValidator.validate(xml) !== true
  )
    throw new Error('office_api_unsupported')
  return parser.parse(xml) as Node[]
}

export interface ComplexPagePackageSummary {
  tables: Array<{
    shapeId: string
    rows: string[][]
    simpleCells: boolean[][]
    truncated?: boolean
  }>
  charts: Array<{
    shapeId: string
    plotTypes: string[]
    visualOptions: {
      barDirections: string[]
      groupings: string[]
      legendPositions: string[]
      valueLabels: string[]
    }
    series: Array<{ name?: string; categories: string[]; values: string[] }>
    cacheOnly: true
    /** Known explicit fields only. A missing field is unverified, never proof of equality. */
    visibleStyle?: Record<string, string>
    truncated?: boolean
  }>
  truncated: boolean
}

const direct = (nodes: Node[], tag: string): Node[] =>
  nodes.filter((node) => Object.hasOwn(node, tag))
function child(nodes: Node[], tag: string): Node[] | undefined {
  const found = direct(nodes, tag)
  return found.length === 1 && Array.isArray(found[0]![tag])
    ? (found[0]![tag] as Node[])
    : undefined
}
const attrs = (node: Node | undefined): Node => (node?.[':@'] ?? {}) as Node
const onlyChildren = (nodes: Node[], allowed: readonly string[]) =>
  nodes.every((node) =>
    Object.keys(node).every(
      (key) =>
        key === ':@' || (key === '#text' && !String(node[key]).trim()) || allowed.includes(key),
    ),
  )
function fill(nodes: Node[]): string | undefined {
  const colors = direct(nodes, 'a:solidFill'),
    none = direct(nodes, 'a:noFill')
  if (
    none.length === 1 &&
    !colors.length &&
    !Object.keys(attrs(none[0])).length &&
    !(none[0]!['a:noFill'] as Node[]).length
  )
    return 'none'
  const solid = child(nodes, 'a:solidFill')
  if (!solid || none.length || !onlyChildren(solid, ['a:srgbClr'])) return undefined
  const rgb = direct(solid, 'a:srgbClr')
  const color = attrs(rgb[0])['@_val']
  if (
    rgb.length !== 1 ||
    Object.keys(attrs(rgb[0])).join(',') !== '@_val' ||
    (rgb[0]!['a:srgbClr'] as Node[]).length ||
    typeof color !== 'string' ||
    !/^[A-Fa-f0-9]{6}$/.test(color)
  )
    return undefined
  return color.toUpperCase()
}
function shapeStyle(nodes: Node[] | undefined): Record<string, string> {
  if (!nodes) return {}
  const effect = child(nodes, 'a:effectLst')
  if (!onlyChildren(nodes, ['a:solidFill', 'a:noFill', 'a:ln', 'a:effectLst']) || effect?.length)
    return {}
  const output: Record<string, string> = {},
    color = fill(nodes)
  if (color !== undefined) output.fill = color
  const line = child(nodes, 'a:ln'),
    lineNode = direct(nodes, 'a:ln')[0]
  if (
    line &&
    onlyChildren(line, ['a:solidFill', 'a:noFill', 'a:prstDash', 'a:round', 'a:bevel', 'a:miter'])
  ) {
    const lineAttrs = attrs(lineNode),
      lineColor = fill(line)
    if (
      Object.keys(lineAttrs).some((key) => !['@_w', '@_cap', '@_cmpd', '@_algn'].includes(key)) ||
      (lineAttrs['@_cmpd'] !== undefined && lineAttrs['@_cmpd'] !== 'sng') ||
      (lineAttrs['@_algn'] !== undefined && lineAttrs['@_algn'] !== 'ctr')
    )
      return output
    if (lineColor !== undefined) {
      output.line = lineColor
      if (lineColor !== 'none') {
        const width = lineAttrs['@_w']
        if (typeof width === 'string' && /^\d+$/.test(width) && Number(width) <= 12_700_000)
          output.lineWidth = String(Number(width))
        const cap = lineAttrs['@_cap'] ?? 'flat'
        if (['flat', 'rnd', 'sq'].includes(String(cap))) output.lineCap = String(cap)
        const dashNodes = direct(line, 'a:prstDash'),
          dash = attrs(dashNodes[0])['@_val'] ?? 'solid'
        if (
          dashNodes.length <= 1 &&
          [
            'solid',
            'dot',
            'dash',
            'lgDash',
            'dashDot',
            'lgDashDot',
            'lgDashDotDot',
            'sysDash',
            'sysDot',
            'sysDashDot',
            'sysDashDotDot',
          ].includes(String(dash))
        )
          output.lineDash = String(dash)
      }
    }
  }
  return output
}
function textStyle(nodes: Node[] | undefined): Record<string, string> {
  if (!nodes || !onlyChildren(nodes, ['a:bodyPr', 'a:lstStyle', 'a:p'])) return {}
  const body = direct(nodes, 'a:bodyPr')[0],
    bodyAttrs = attrs(body)
  if (
    Object.keys(bodyAttrs).some(
      (key) => !['@_rot', '@_vert', '@_anchor', '@_wrap'].includes(key),
    ) ||
    (bodyAttrs['@_rot'] !== undefined && Number(bodyAttrs['@_rot']) !== 0) ||
    (bodyAttrs['@_vert'] !== undefined && bodyAttrs['@_vert'] !== 'horz') ||
    child(nodes, 'a:lstStyle')?.length
  )
    return {}
  const paragraphs = direct(nodes, 'a:p')
  if (paragraphs.length !== 1) return {}
  const paragraph = paragraphs[0]!['a:p'] as Node[],
    properties = child(paragraph, 'a:pPr'),
    defaults = properties && child(properties, 'a:defRPr')
  if (
    !defaults ||
    !onlyChildren(paragraph, ['a:pPr', 'a:endParaRPr']) ||
    !onlyChildren(defaults, ['a:solidFill', 'a:latin', 'a:ea', 'a:cs'])
  )
    return {}
  const output: Record<string, string> = {},
    color = fill(defaults)
  if (color !== undefined && color !== 'none') output.color = color
  for (const [tag, key] of [
    ['a:latin', 'font'],
    ['a:ea', 'eastAsianFont'],
    ['a:cs', 'complexFont'],
  ]) {
    const fontNodes = direct(defaults, tag!),
      face = attrs(fontNodes[0])['@_typeface']
    if (
      fontNodes.length === 1 &&
      typeof face === 'string' &&
      face.length > 0 &&
      face.length <= 128 &&
      !face.startsWith('+') &&
      !(fontNodes[0]![tag!] as Node[]).length
    )
      output[key!] = face.toLowerCase()
  }
  const run = attrs(direct(properties!, 'a:defRPr')[0])
  if (
    typeof run['@_sz'] === 'string' &&
    /^\d+$/.test(run['@_sz']) &&
    Number(run['@_sz']) > 0 &&
    Number(run['@_sz']) <= 40000
  )
    output.size = String(Number(run['@_sz']))
  for (const key of ['b', 'i']) {
    const val = run[`@_${key}`] ?? '0'
    if (['0', '1', 'false', 'true'].includes(String(val)))
      output[key] = val === '1' || val === 'true' ? '1' : '0'
  }
  for (const [key, fallback] of [
    ['u', 'none'],
    ['strike', 'noStrike'],
  ])
    if (typeof (run[`@_${key}`] ?? fallback) === 'string')
      output[key!] = String(run[`@_${key}`] ?? fallback)
  return output
}
/** Conservative semantic projection: ordinary explicit RGB styles, never inherited theme proof. */
function chartVisibleStyle(nodes: Node[], plotTypes: string[]): Record<string, string> | undefined {
  if (plotTypes.length !== 1 || !['barChart', 'lineChart', 'pieChart'].includes(plotTypes[0]!))
    return undefined
  const space = child(nodes, 'c:chartSpace'),
    chart = space && child(space, 'c:chart'),
    area = chart && child(chart, 'c:plotArea')
  if (!space || !chart || !area) return undefined
  const output: Record<string, string> = {}
  const record = (role: string, values: Record<string, string>) => {
    for (const [key, val] of Object.entries(values)) output[`${role}.${key}`] = val
  }
  const common = (role: string, container: Node[]) => {
    record(role, shapeStyle(child(container, 'c:spPr')))
    record(`${role}.text`, textStyle(child(container, 'c:txPr')))
  }
  const inheritedCommon = (role: string, container: Node[], parent: string) => {
    const inherited: Record<string, string> = {}
    for (const [key, value] of Object.entries(output)) {
      if (!key.startsWith(`${parent}.`)) continue
      const suffix = key.slice(parent.length + 1)
      if (suffix.includes('point.') || suffix === '$explicit') continue
      if (
        suffix.startsWith('text.')
          ? !direct(container, 'c:txPr').length
          : !direct(container, 'c:spPr').length
      )
        inherited[suffix] = value
    }
    record(role, inherited)
    output[`${role}.$explicit`] = '1'
    common(role, container)
  }
  common('chart', space)
  common('plot', area)
  const legend = child(chart, 'c:legend')
  if (legend) common('legend', legend)
  const plot = child(area, `c:${plotTypes[0]}`)
  if (!plot) return undefined
  const labels = child(plot, 'c:dLbls')
  if (labels) common('labels', labels)
  const series = direct(plot, 'c:ser')
  if (series.length > 10) return undefined
  for (let index = 0; index < series.length; index++) {
    const ser = series[index]!['c:ser'] as Node[],
      role = `series.${index}`
    common(role, ser)
    const label = child(ser, 'c:dLbls')
    if (label) {
      inheritedCommon(`${role}.labels`, label, 'labels')
      const pointLabels = direct(label, 'c:dLbl')
      if (pointLabels.length > 50) return undefined
      const labelIndices = new Set<number>()
      for (const pointLabel of pointLabels) {
        const parts = pointLabel['c:dLbl'] as Node[]
        const labelIndex = attrs(direct(parts, 'c:idx')[0])['@_val']
        if (
          typeof labelIndex !== 'string' ||
          !/^\d+$/.test(labelIndex) ||
          Number(labelIndex) > 49 ||
          labelIndices.has(Number(labelIndex))
        )
          return undefined
        labelIndices.add(Number(labelIndex))
        inheritedCommon(`${role}.labels.point.${Number(labelIndex)}`, parts, `${role}.labels`)
      }
    }
    const marker = child(ser, 'c:marker')
    if (marker) common(`${role}.marker`, marker)
    const points = direct(ser, 'c:dPt')
    if (points.length > 50) return undefined
    const seen = new Set<string>()
    for (const point of points) {
      const parts = point['c:dPt'] as Node[],
        pointIndex = attrs(direct(parts, 'c:idx')[0])['@_val']
      if (
        typeof pointIndex !== 'string' ||
        !/^\d+$/.test(pointIndex) ||
        Number(pointIndex) > 49 ||
        seen.has(String(Number(pointIndex)))
      )
        return undefined
      seen.add(String(Number(pointIndex)))
      const parentProperties = child(ser, 'c:spPr'),
        pointProperties = child(parts, 'c:spPr')
      const inherited = shapeStyle(parentProperties),
        local = shapeStyle(pointProperties)
      // An omitted point line inherits the series border. Explicit but unsupported overrides
      // do not inherit known parent values and must remain unverified.
      if (pointProperties && direct(pointProperties, 'a:ln').length) {
        for (const key of Object.keys(inherited)) if (key.startsWith('line')) delete inherited[key]
      }
      if (
        pointProperties &&
        pointProperties.some((node) =>
          Object.keys(node).some((key) =>
            [
              'a:solidFill',
              'a:noFill',
              'a:gradFill',
              'a:blipFill',
              'a:pattFill',
              'a:grpFill',
            ].includes(key),
          ),
        )
      )
        delete inherited.fill
      output[`${role}.point.${Number(pointIndex)}.$explicit`] = '1'
      record(`${role}.point.${Number(pointIndex)}`, { ...inherited, ...local })
      const pointMarker = child(parts, 'c:marker')
      if (pointMarker) {
        inheritedCommon(`${role}.point.${Number(pointIndex)}.marker`, pointMarker, `${role}.marker`)
      }
    }
  }
  for (const axis of ['catAx', 'valAx']) {
    const axes = direct(area, `c:${axis}`)
    if (axes.length > 1) return undefined
    if (!axes.length) continue
    const parts = axes[0]![`c:${axis}`] as Node[]
    common(axis, parts)
    for (const tag of ['majorGridlines', 'minorGridlines']) {
      const grid = child(parts, `c:${tag}`)
      if (grid) common(`${axis}.${tag}`, grid)
      else output[`${axis}.${tag}.line`] = 'none'
    }
  }
  const ordered = Object.fromEntries(Object.entries(output).sort(([a], [b]) => a.localeCompare(b)))
  return new TextEncoder().encode(JSON.stringify(ordered)).length <= 64 * 1024 ? ordered : undefined
}

/** Read only the exported page and embedded chart caches; workbook links are never followed. */
export async function inspectPowerPointComplexPagePackage(
  base64: string,
  signal?: AbortSignal,
  options: { slideIndex?: number; maxBytes?: number; includeVisibleStyle?: boolean } = {},
): Promise<ComplexPagePackageSummary> {
  const zip = await loadBoundedZip(base64, signal, true, options.maxBytes)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  const slidePath =
    options.slideIndex === undefined
      ? slides.length === 1
        ? slides[0]
        : undefined
      : slides.includes(`ppt/slides/slide${options.slideIndex + 1}.xml`)
        ? `ppt/slides/slide${options.slideIndex + 1}.xml`
        : undefined
  if (!slidePath) throw new Error('office_api_unsupported')
  const slide = xmlNodes(await zip.file(slidePath)!.async('string'))
  const frames = elements(slide, 'p:graphicFrame')
  const output: ComplexPagePackageSummary = { tables: [], charts: [], truncated: false }
  // Fixed text budget also covers JSON escaping and the bounded array structure below.
  let remainingText = 8 * 1024
  const clip = (text: string, limit = 128): string => {
    const length = Math.min(text.length, limit, remainingText)
    if (text.length > length) output.truncated = true
    remainingText -= length
    return text.slice(0, length)
  }
  const relPath = slidePath.replace('/slides/', '/slides/_rels/') + '.rels'
  let relationships: Node[] | undefined
  const findChart = async (id: string): Promise<Node[]> => {
    if (!relationships) {
      const file = zip.file(relPath)
      if (!file) throw new Error('office_api_unsupported')
      relationships = xmlNodes(await file.async('string'))
    }
    const relNodes = relationshipNodes(relationships).filter(
      (node) => (node[':@'] as Node | undefined)?.['@_Id'] === id,
    )
    if (relNodes.length !== 1) throw new Error('office_api_unsupported')
    const attrs = relNodes[0]![':@'] as Node
    const target = attrs['@_Target']
    const type = attrs['@_Type']
    const chartPath =
      typeof target === 'string' && /^\/ppt\/charts\/chart\d+\.xml$/.test(target)
        ? target.slice(1)
        : typeof target === 'string' && /^\.\.\/charts\/chart\d+\.xml$/.test(target)
          ? `ppt/charts/${target.slice('../charts/'.length)}`
          : undefined
    if (
      attrs['@_TargetMode'] === 'External' ||
      !chartPath ||
      typeof type !== 'string' ||
      !type.endsWith('/chart')
    )
      throw new Error('office_api_unsupported')
    const file = zip.file(chartPath)
    if (!file) throw new Error('office_api_unsupported')
    return xmlNodes(await file.async('string'))
  }
  for (const frame of frames) {
    if (signal?.aborted) throw new Error('cancelled')
    const rawShapeId = frameId(frame)
    const shapeId = rawShapeId ? clip(rawShapeId, 64) : undefined
    if (!shapeId) continue
    const tables = elements(frame, 'a:tbl')
    if (tables.length) {
      if (output.tables.length >= 16) {
        output.truncated = true
        continue
      }
      const hasMergedCells = tagNodes(tables[0]!, 'a:tc').some((node) => {
        const attrs = node[':@'] as Node | undefined
        return (
          attrs &&
          ['@_gridSpan', '@_rowSpan', '@_hMerge', '@_vMerge'].some((name) =>
            Object.hasOwn(attrs, name),
          )
        )
      })
      const allRows = elements(tables[0]!, 'a:tr')
      const simpleCells: boolean[][] = []
      const rows = allRows.slice(0, 20).map((row) => {
        const cells = elements(row, 'a:tc')
        if (cells.length > 12) output.truncated = true
        const selected = cells.slice(0, 12)
        simpleCells.push(
          selected.map(
            (cell) =>
              !hasMergedCells &&
              elements(cell, 'a:p').length <= 1 &&
              elements(cell, 'a:r').length <= 1 &&
              elements(cell, 'a:br').length === 0 &&
              elements(cell, 'a:fld').length === 0 &&
              elements(cell, 'a:rPr').length === 0 &&
              elements(cell, 'a:pPr').length === 0 &&
              elements(cell, 'a:endParaRPr').length === 0,
          ),
        )
        return selected.map((cell) => clip(value(cell)))
      })
      if (allRows.length > 20) output.truncated = true
      output.tables.push({
        shapeId,
        rows,
        simpleCells,
        ...(output.truncated ? { truncated: true } : {}),
      })
      continue
    }
    const chartTags = chartNodes(frame)
    if (!chartTags.length) continue
    if (output.charts.length >= 16) {
      output.truncated = true
      continue
    }
    const id = (chartTags[0]![':@'] as Node | undefined)?.['@_r:id']
    if (typeof id !== 'string') throw new Error('office_api_unsupported')
    const chart = await findChart(id)
    const plotTypes = [
      'barChart',
      'lineChart',
      'areaChart',
      'pieChart',
      'doughnutChart',
      'scatterChart',
      'bubbleChart',
      'radarChart',
      'stockChart',
      'surfaceChart',
      'ofPieChart',
    ].filter((type) => elements(chart, `c:${type}`).length > 0)
    if (!plotTypes.length) output.truncated = true
    const allSeries = elements(chart, 'c:ser')
    const visualValues = (tag: string): string[] => {
      const nodes = tagNodes(chart, tag)
      if (nodes.length > 8) output.truncated = true
      return nodes.slice(0, 8).map((node) => {
        const val = (node[':@'] as Node | undefined)?.['@_val']
        if (typeof val !== 'string' || val.length > 32) {
          output.truncated = true
          return ''
        }
        return val
      })
    }
    const visualOptions = {
      barDirections: visualValues('c:barDir'),
      groupings: visualValues('c:grouping'),
      legendPositions: visualValues('c:legendPos'),
      valueLabels: visualValues('c:showVal'),
    }
    const series = allSeries.slice(0, 8).map((ser) => {
      const cache = (container: string): string[] => {
        const groups = elements(ser, container)
        if (!groups.length) return []
        const multi = elements(groups[0]!, 'c:multiLvlStrCache')[0]
        const levels = multi ? elements(multi, 'c:lvl') : []
        if (multi && levels.length !== 1) output.truncated = true
        const cached = multi
          ? levels.length === 1
            ? levels[0]
            : undefined
          : (elements(groups[0]!, 'c:strCache')[0] ?? elements(groups[0]!, 'c:numCache')[0])
        if (!cached) return []
        const points = tagNodes(cached, 'c:pt')
        if (points.length > 32) output.truncated = true
        return points.slice(0, 32).map((point, index) => {
          if ((point[':@'] as Node | undefined)?.['@_idx'] !== String(index))
            output.truncated = true
          return clip(
            elements(point['c:pt'] as Node[], 'c:v')[0]
              ?.map((item) => item['#text'] ?? '')
              .join('') ?? '',
          )
        })
      }
      const tx = elements(ser, 'c:tx')[0] ?? []
      const name = elements(tx, 'c:v')[0]
        ?.map((item) => item['#text'] ?? '')
        .join('')
      return {
        ...(name ? { name: clip(name) } : {}),
        categories: cache('c:cat'),
        values: cache('c:val'),
      }
    })
    if (allSeries.length > 8) output.truncated = true
    const visibleStyle = options.includeVisibleStyle
      ? chartVisibleStyle(chart, plotTypes)
      : undefined
    output.charts.push({
      shapeId,
      plotTypes,
      visualOptions,
      series,
      cacheOnly: true,
      ...(visibleStyle ? { visibleStyle } : {}),
      ...(output.truncated ? { truncated: true } : {}),
    })
  }
  return output
}

function relationshipNodes(nodes: Node[]): Node[] {
  return tagNodes(nodes, 'Relationship')
}

function chartNodes(nodes: Node[]): Node[] {
  return tagNodes(nodes, 'c:chart')
}

function tagNodes(nodes: Node[], tag: string): Node[] {
  return nodes.flatMap((node) => [
    ...(Object.hasOwn(node, tag) ? [node] : []),
    ...Object.values(node)
      .filter(Array.isArray)
      .flatMap((parts) => tagNodes(parts as Node[], tag)),
  ])
}

function frameId(nodes: Node[]): string | undefined {
  for (const node of nodes) {
    if (Object.hasOwn(node, 'p:cNvPr')) {
      const id = (node[':@'] as Node | undefined)?.['@_id']
      return typeof id === 'string' ? id : undefined
    }
    for (const parts of Object.values(node).filter(Array.isArray)) {
      const id = frameId(parts as Node[])
      if (id) return id
    }
  }
  return undefined
}

/** Exact-cell evidence for reversible text edits. The digest excludes only text payloads. */
export async function inspectPowerPointTableCellPackage(
  base64: string,
  shapeId: string,
  rowIndex: number,
  columnIndex: number,
  signal?: AbortSignal,
): Promise<{ text: string; structureDigest: string }> {
  const result = await inspectPowerPointTableCellsPackage(
    base64,
    shapeId,
    [{ rowIndex, columnIndex }],
    signal,
  )
  return { text: result.cells[0]!.text, structureDigest: result.structureDigest }
}

/** Whole-table evidence with only the selected cell text payloads excluded. */
export async function inspectPowerPointTableCellsPackage(
  base64: string,
  shapeId: string,
  cells: Array<{ rowIndex: number; columnIndex: number }>,
  signal?: AbortSignal,
): Promise<{
  cells: Array<{ rowIndex: number; columnIndex: number; text: string }>
  structureDigest: string
}> {
  if (signal?.aborted) throw new Error('cancelled')
  if (
    !Array.isArray(cells) ||
    cells.length < 1 ||
    cells.length > 8 ||
    cells.some(
      (cell) =>
        !cell ||
        !Number.isSafeInteger(cell.rowIndex) ||
        cell.rowIndex < 0 ||
        !Number.isSafeInteger(cell.columnIndex) ||
        cell.columnIndex < 0,
    ) ||
    new Set(cells.map(({ rowIndex, columnIndex }) => `${rowIndex}:${columnIndex}`)).size !==
      cells.length
  )
    throw new Error('invalid_tool_input')
  const zip = await loadBoundedZip(base64, signal)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) throw new Error('office_api_unsupported')
  const slide = xmlNodes(await zip.file(slides[0]!)!.async('string'))
  const frames = elements(slide, 'p:graphicFrame').filter((item) => frameId(item) === shapeId)
  if (frames.length !== 1) throw new Error('presentation_existing_target_unsupported')
  const tables = elements(frames[0]!, 'a:tbl')
  if (!tables || tables.length !== 1) throw new Error('presentation_existing_target_unsupported')
  if (
    tagNodes(tables[0]!, 'a:tc').some((node) => {
      const attrs = node[':@'] as Node | undefined
      return (
        attrs &&
        ['@_gridSpan', '@_rowSpan', '@_hMerge', '@_vMerge'].some((name) =>
          Object.hasOwn(attrs, name),
        )
      )
    })
  )
    throw new Error('presentation_existing_target_unsupported')
  const rows = elements(tables[0]!, 'a:tr')
  const selected = cells.map(({ rowIndex, columnIndex }) => {
    const cell = rows[rowIndex] && elements(rows[rowIndex]!, 'a:tc')[columnIndex]
    if (
      !cell ||
      elements(cell, 'a:p').length > 1 ||
      elements(cell, 'a:r').length > 1 ||
      ['a:br', 'a:fld', 'a:rPr', 'a:pPr', 'a:endParaRPr'].some((tag) => elements(cell, tag).length)
    )
      throw new Error('presentation_existing_target_unsupported')
    const text = value(cell)
    if (text.length > 128) throw new Error('presentation_existing_target_unsupported')
    return { rowIndex, columnIndex, text }
  })
  const tableForDigest = structuredClone(tables[0]!)
  const blankTargetText = (nodes: Node[]) => {
    for (const node of nodes)
      for (const [key, entry] of Object.entries(node)) {
        if (key === 'a:t') node[key] = []
        else if (Array.isArray(entry)) blankTargetText(entry as Node[])
      }
  }
  const digestRows = elements(tableForDigest, 'a:tr')
  for (const { rowIndex, columnIndex } of cells) {
    const digestRow = digestRows[rowIndex]
    const digestCell = digestRow && elements(digestRow, 'a:tc')[columnIndex]
    if (!digestCell) throw new Error('presentation_existing_target_unsupported')
    blankTargetText(digestCell)
  }
  const hash = await crypto.subtle.digest(
    'SHA-256',
    new TextEncoder().encode(JSON.stringify(tableForDigest)),
  )
  if (signal?.aborted) throw new Error('cancelled')
  return {
    cells: selected,
    structureDigest: Array.from(new Uint8Array(hash), (byte) =>
      byte.toString(16).padStart(2, '0'),
    ).join(''),
  }
}
