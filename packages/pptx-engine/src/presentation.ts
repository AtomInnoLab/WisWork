import {
  type Schema,
  text,
  number,
  choice,
  array,
  object,
  id,
  color,
  valid,
} from './presentation-schema'

/** Browser-safe, fixed 16:9 presentation contract. Coordinates are inches, text sizes points. */
export const PRESENTATION_WIDTH = 13.333333
export const PRESENTATION_HEIGHT = 7.5
export const PRESENTATION_TEXT_BUDGET = 250_000
export interface PresentationStyle {
  fontFace: string
  background: string
  textColor: string
  accentColor: string
}
export type PresentationAsset = PresentationInlineAsset | { id: string; attachmentId: string }
export interface PresentationInlineAsset {
  id: string
  mime: 'image/png' | 'image/jpeg'
  base64: string
  width: number
  height: number
  source?: string
  license?: 'owned' | 'licensed' | 'public_domain' | 'unknown'
}
export interface PresentationClaim {
  id: string
  text: string
  source: string
  locator?: string
}
export interface ElementGeometry {
  id: string
  x: number
  y: number
  w: number
  h: number
  role?: 'content' | 'background' | 'decoration'
  allowOverlap?: boolean
}
export type SlideIRElement = ElementGeometry &
  (
    | {
        kind: 'text'
        text: string
        fontSize?: number
        color?: string
        bold?: boolean
        align?: 'left' | 'center' | 'right'
      }
    | { kind: 'shape'; shape: 'rect' | 'ellipse' | 'roundRect'; fill?: string; lineColor?: string }
    | { kind: 'image'; assetId: string; fit?: 'contain' | 'cover'; altText?: string }
    | { kind: 'table'; rows: string[][]; fontSize?: number }
    | {
        kind: 'chart'
        chartType: 'bar' | 'line' | 'pie'
        categories: string[]
        series: { name: string; values: number[] }[]
      }
  )
export interface SlideIR {
  id: string
  title: string
  notes?: string
  claimIds?: string[]
  elements: SlideIRElement[]
}
export interface PresentationDeck {
  version: 1
  id: string
  title: string
  style: PresentationStyle
  assets: PresentationAsset[]
  claims: PresentationClaim[]
  slides: SlideIR[]
}
export interface GeometryIssue {
  kind: 'out_of_bounds' | 'overlap'
  slideId: string
  elementIds: string[]
}
export interface PresentationCompileReport {
  deckId: string
  slideCount: number
  elementCount: number
  geometry: GeometryIssue[]
  assetWarnings?: { missingSource: number; unknownLicense: number; missingAltText: number }
  checks: {
    structure: 'passed'
    geometry: 'passed' | 'warning'
    render: 'not_run'
    sources: 'not_verified'
    roundTrip: 'not_run'
  }
}

const geometry = {
  id,
  x: number(0, PRESENTATION_WIDTH),
  y: number(0, PRESENTATION_HEIGHT),
  w: number(0.01, PRESENTATION_WIDTH),
  h: number(0.01, PRESENTATION_HEIGHT),
  role: choice('content', 'background', 'decoration'),
  allowOverlap: { type: 'boolean' },
}
const element = (properties: Record<string, Schema>, required: string[]): Schema =>
  object({ ...geometry, ...properties }, ['id', 'x', 'y', 'w', 'h', ...required])
/** Single schema shared by tool discovery and runtime validation; arbitrary code/options are rejected. */
export const PRESENTATION_DECK_SCHEMA: Schema = object({
  version: { type: 'number', enum: [1] },
  id,
  title: text(300, 1),
  style: object({
    fontFace: { ...text(80, 1), pattern: '^[^<>\\r\\n]+$' },
    background: color,
    textColor: color,
    accentColor: color,
  }),
  assets: array(
    {
      anyOf: [
        object({ id, attachmentId: { ...text(64, 64), pattern: '^[a-f0-9]{64}$' } }),
        object(
          {
            id,
            mime: choice('image/png', 'image/jpeg'),
            base64: { ...text(5_600_000, 4), pattern: '^[A-Za-z0-9+/]*={0,2}$' },
            width: number(1, 16384),
            height: number(1, 16384),
            source: text(2000, 1),
            license: choice('owned', 'licensed', 'public_domain', 'unknown'),
          },
          ['id', 'mime', 'base64', 'width', 'height'],
        ),
      ],
    },
    undefined,
  ),
  claims: array(
    object({ id, text: text(12000, 1), source: text(2000, 1), locator: text(1000, 1) }, [
      'id',
      'text',
      'source',
    ]),
    256,
  ),
  slides: array(
    object(
      {
        id,
        title: text(300, 1),
        notes: text(12000),
        claimIds: array(id, 32),
        elements: array(
          {
            anyOf: [
              element(
                {
                  kind: choice('text'),
                  text: text(12000, 1),
                  fontSize: number(6, 96),
                  color,
                  bold: { type: 'boolean' },
                  align: choice('left', 'center', 'right'),
                },
                ['kind', 'text'],
              ),
              element(
                {
                  kind: choice('shape'),
                  shape: choice('rect', 'ellipse', 'roundRect'),
                  fill: color,
                  lineColor: color,
                },
                ['kind', 'shape'],
              ),
              element({ kind: choice('image'), assetId: id, fit: choice('contain', 'cover'), altText: text(500, 1) }, [
                'kind',
                'assetId',
              ]),
              element(
                {
                  kind: choice('table'),
                  rows: array(array(text(2000), 16, 1), 50, 1),
                  fontSize: number(6, 48),
                },
                ['kind', 'rows'],
              ),
              element(
                {
                  kind: choice('chart'),
                  chartType: choice('bar', 'line', 'pie'),
                  categories: array(text(200, 1), 50, 1),
                  series: array(
                    object({ name: text(200, 1), values: array(number(-1e15, 1e15), 50, 1) }),
                    10,
                    1,
                  ),
                },
                ['kind', 'chartType', 'categories', 'series'],
              ),
            ],
          },
          128,
          1,
        ),
      },
      ['id', 'title', 'elements'],
    ),
    32,
    1,
  ),
})

function reject(reason: string): never {
  throw new Error(`presentation_invalid:${reason}`)
}
function unique(items: { id: string }[], label: string): void {
  if (new Set(items.map((item) => item.id)).size !== items.length) reject(`duplicate_${label}`)
}

export function parsePresentationDeck(input: unknown): PresentationDeck {
  if (!valid(input, PRESENTATION_DECK_SCHEMA)) reject('schema')
  const deck = input as PresentationDeck
  unique(deck.slides, 'slide')
  unique(deck.assets, 'asset')
  unique(deck.claims, 'claim')
  if (
    deck.assets.reduce((sum, asset) => sum + ('base64' in asset ? asset.base64.length : 0), 0) >
    28_000_000
  )
    reject('asset_budget')
  // Attachment references must be resolved by the document-bound PC service before compilation.
  for (const asset of deck.assets) {
    if ('attachmentId' in asset) continue
    if (asset.base64.length % 4 !== 0) reject('image_encoding')
    if (
      !Number.isInteger(asset.width) ||
      !Number.isInteger(asset.height) ||
      asset.width * asset.height > 40_000_000
    )
      reject('image_dimensions')
    if (
      !(asset.mime === 'image/png'
        ? asset.base64.startsWith('iVBORw0KGgo')
        : asset.base64.startsWith('/9j/'))
    )
      reject('image_signature')
  }
  const assetIds = new Set(deck.assets.map((asset) => asset.id))
  const claimIds = new Set(deck.claims.map((claim) => claim.id))
  let textBudget = deck.claims.reduce(
    (sum, claim) => sum + claim.text.length + claim.source.length + (claim.locator?.length ?? 0),
    0,
  )
  for (const slide of deck.slides) {
    unique(slide.elements, 'element')
    if (
      slide.claimIds?.some((claim) => !claimIds.has(claim)) ||
      new Set(slide.claimIds).size !== (slide.claimIds?.length ?? 0)
    )
      reject('claim_reference')
    textBudget += slide.title.length + (slide.notes?.length ?? 0)
    for (const el of slide.elements) {
      if (el.kind === 'image' && !assetIds.has(el.assetId)) reject('asset_reference')
      if (el.kind === 'text') textBudget += el.text.length
      if (el.kind === 'table') {
        if (el.rows.some((row) => row.length !== el.rows[0]!.length)) reject('ragged_table')
        textBudget += el.rows.flat().join('').length
      }
      if (el.kind === 'chart') {
        if (el.series.some((series) => series.values.length !== el.categories.length))
          reject('chart_lengths')
        if (
          el.chartType === 'pie' &&
          (el.series.length !== 1 ||
            el.series[0]!.values.some((value) => value < 0) ||
            !el.series[0]!.values.some((value) => value > 0))
        )
          reject('pie_values')
        textBudget +=
          el.categories.join('').length + el.series.map((series) => series.name).join('').length
      }
    }
  }
  if (textBudget > PRESENTATION_TEXT_BUDGET) reject('text_budget')
  return structuredClone(deck)
}

export function inspectPresentationGeometry(deck: PresentationDeck): GeometryIssue[] {
  const issues: GeometryIssue[] = []
  for (const slide of deck.slides) {
    for (const el of slide.elements) {
      const bottom =
        slide.claimIds?.length && el.role !== 'background' && el.role !== 'decoration'
          ? 6.95
          : PRESENTATION_HEIGHT
      if (
        el.x < 0 ||
        el.y < 0 ||
        el.x + el.w > PRESENTATION_WIDTH + 1e-6 ||
        el.y + el.h > bottom + 1e-6
      )
        issues.push({ kind: 'out_of_bounds', slideId: slide.id, elementIds: [el.id] })
    }
    const content = slide.elements.filter(
      (el) => (!el.role || el.role === 'content') && !el.allowOverlap,
    )
    for (let i = 0; i < content.length; i++)
      for (let j = i + 1; j < content.length; j++) {
        const a = content[i]!,
          b = content[j]!
        if (
          Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) > 0.001 &&
          Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) > 0.001
        )
          issues.push({ kind: 'overlap', slideId: slide.id, elementIds: [a.id, b.id] })
      }
  }
  return issues
}
