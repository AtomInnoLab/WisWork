import PptxGenJS from 'pptxgenjs'
import JSZip from 'jszip'
import { XMLParser, XMLValidator } from 'fast-xml-parser'
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

function xmlText(value: unknown): string {
  if (typeof value === 'string') return value
  if (Array.isArray(value)) return value.map(xmlText).join('')
  if (!value || typeof value !== 'object') return ''
  return Object.entries(value)
    .map(([key, child]) =>
      key === 'a:t' ? xmlText(child) : key.startsWith('@_') ? '' : xmlText(child),
    )
    .join('')
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
        text: sourceLabels.join('；').slice(0, 500),
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
        if (element.kind === 'text' && xmlText(object['p:txBody']) !== element.text)
          throw new Error('presentation_compile:structure_mismatch')
        if (element.kind === 'table') {
          const rows = xmlItems(object['a:graphic']['a:graphicData']['a:tbl']['a:tr']).map((row) =>
            xmlItems(row['a:tc']).map((cell) => xmlText(cell['a:txBody'])),
          )
          if (JSON.stringify(rows) !== JSON.stringify(element.rows))
            throw new Error('presentation_compile:structure_mismatch')
        }
        if (element.kind === 'image') {
          const id = object['p:blipFill']?.['a:blip']?.['@_r:embed']
          if (typeof id !== 'string') throw new Error('presentation_compile:structure_mismatch')
          const path = await linkedPart(id, 'image')
          const asset = deck.assets.find((item) => item.id === element.assetId)
          if (!asset || !('base64' in asset))
            throw new Error('presentation_compile:structure_mismatch')
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
          const chartXml = await zip.file(await linkedPart(id, 'chart'))!.async('string')
          if (chartXml.length > 2 * 1024 * 1024 || XMLValidator.validate(chartXml) !== true)
            throw new Error('presentation_compile:structure_mismatch')
          const chart = parseChartXml(chartXml)
          const chartRoot = parser.parse(chartXml) as XmlNode
          const plotArea = chartRoot['c:chartSpace']?.['c:chart']?.['c:plotArea'] as
            XmlNode | undefined
          const plotTag = `${element.chartType === 'bar' ? 'bar' : element.chartType}Chart`
          const plots = xmlItems(plotArea?.[`c:${plotTag}`])
          const seriesNodes = plots.flatMap((plot) => xmlItems(plot['c:ser']))
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
  options: { trustedAssetEvidence?: boolean } = {},
): Promise<{ bytes: Uint8Array; report: PresentationCompileReport; sourceSlideIds?: string[] }> {
  const deck = parsePresentationDeck(input, options)
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
            chartColors: [deck.style.accentColor],
            showValue: true,
            catAxisLabelFontFace: deck.style.fontFace,
            valAxisLabelFontFace: deck.style.fontFace,
            legendFontFace: deck.style.fontFace,
          },
        )
    }
    const claims = (ir.claimIds ?? []).map((id) => deck.claims.find((claim) => claim.id === id)!)
    const sources = presentationSlideSourceLabels(ir, deck.claims)
    if (sources.length)
      slide.addText(sources.join('；').slice(0, 500), {
        x: 0.5,
        y: 7.05,
        w: 12.3,
        h: 0.3,
        fontFace: deck.style.fontFace,
        fontSize: 8,
        color: deck.style.textColor,
        margin: 0,
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
  await verifyCompiledPresentationStructure(zip, deck)
  return {
    bytes: output,
    sourceSlideIds,
    report: {
      deckId: deck.id,
      slideCount: deck.slides.length,
      elementCount: deck.slides.reduce((n, slide) => n + slide.elements.length, 0),
      geometry,
      assetWarnings,
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
