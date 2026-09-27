import PptxGenJS from 'pptxgenjs'
import JSZip from 'jszip'
import { XMLParser } from 'fast-xml-parser'
import {
  inspectPresentationGeometry,
  parsePresentationDeck,
  PRESENTATION_HEIGHT,
  PRESENTATION_WIDTH,
  presentationSlideSourceLabels,
  type PresentationInlineAsset,
  type PresentationCompileReport,
} from './presentation'

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
