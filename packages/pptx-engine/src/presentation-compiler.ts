import PptxGenJS from 'pptxgenjs'
import {
  inspectPresentationGeometry,
  parsePresentationDeck,
  PRESENTATION_HEIGHT,
  PRESENTATION_WIDTH,
  type PresentationAsset,
  type PresentationCompileReport,
} from './presentation'

/** Validate encoded raster dimensions before passing bytes to the PPTX writer. No external I/O. */
function imageData(asset: PresentationAsset): string {
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
): Promise<{ bytes: Uint8Array; report: PresentationCompileReport }> {
  const deck = parsePresentationDeck(input)
  const geometry = inspectPresentationGeometry(deck)
  if (geometry.some((issue) => issue.kind === 'out_of_bounds'))
    throw new Error('presentation_geometry:out_of_bounds')
  const assets = new Map(
    deck.assets.map((asset) => [asset.id, { ...asset, data: imageData(asset) }]),
  )
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
        // Contain uses verified dimensions. Cover uses native image crop, never rasterizes text.
        if (el.fit === 'cover') {
          slide.addImage({ ...box, data: asset.data, sizing: { type: 'cover', w: el.w, h: el.h } })
        } else {
          const scale = Math.min(el.w / asset.width, el.h / asset.height)
          const w = asset.width * scale,
            h = asset.height * scale
          slide.addImage({
            ...box,
            data: asset.data,
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
    const sources = claims.map(
      (claim) => `[${claim.id}] ${claim.source}${claim.locator ? ` · ${claim.locator}` : ''}`,
    )
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
        return asset.source
          ? `Image [${asset.id}]: ${asset.source}`
          : `Image [${asset.id}]: source not supplied`
      })
    slide.addNotes(
      [
        ir.title,
        ir.notes ?? '',
        ...claims.map((claim, i) => `${sources[i]}\n${claim.text}`),
        ...assetSources,
        'Attribution supplied by the input; source accuracy has not been verified.',
      ]
        .filter(Boolean)
        .join('\n\n'),
    )
  }
  const output = await pptx.write({ outputType: 'uint8array', compression: true })
  if (!(output instanceof Uint8Array)) throw new Error('presentation_compile:unexpected_output')
  return {
    bytes: output,
    report: {
      deckId: deck.id,
      slideCount: deck.slides.length,
      elementCount: deck.slides.reduce((n, slide) => n + slide.elements.length, 0),
      geometry,
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
