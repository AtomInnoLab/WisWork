import type { PowerPointShape } from './browser-powerpoint-adapter.js'

export interface PresentationBaselineContext {
  slideIds: string[]
  selectedSlideIds: string[]
  selectedShapeIds: string[]
  slideWidth?: number
  slideHeight?: number
}
export interface PresentationBaselinePage {
  slideId: string
  shapes: Array<
    PowerPointShape & {
      text?: string
      font?: { name: string | null; size: number | null; color: string | null; bold?: boolean | null; italic?: boolean | null; underline?: string | null }
    }
  >
  masterId?: string
  layoutId?: string
}
/** Only plain target text with determinate aggregate font attributes uses whole-range replacement. */
export function nativePlainTextEditable(shape: PresentationBaselinePage['shapes'][number]): boolean {
  return (shape.type === 'TextBox' || shape.type === 'GeometricShape') &&
    shape.text !== undefined &&
    (shape.text.length === 0 ||
      (shape.font !== undefined && shape.font.name !== null && shape.font.size !== null && shape.font.color !== null &&
        typeof shape.font.bold === 'boolean' && typeof shape.font.italic === 'boolean' &&
        typeof shape.font.underline === 'string'))
}
export interface PresentationBaselineAdapter {
  readContext(signal?: AbortSignal): Promise<PresentationBaselineContext>
  readPage(slideId: string, signal?: AbortSignal): Promise<PresentationBaselinePage>
}
const MAX_SLIDES = 500
const MAX_SHAPES = 100
const MAX_TEXT = 12_000
const MAX_PAGE_TEXT = 120_000
const MAX_BYTES = 256 * 1024

function cancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}
function invalid(): never {
  throw new Error('office_read_failed')
}
function id(value: unknown): string {
  if (typeof value !== 'string' || !value.length || value.length > 256) invalid()
  return value
}
function ids(items: Array<{ id: string }>, maximum: number): string[] {
  if (!Array.isArray(items)) invalid()
  if (items.length > maximum) throw new Error('presentation_baseline_limit_exceeded')
  const result = items.map((item) => id(item?.id))
  if (new Set(result).size !== result.length) invalid()
  return result
}
function number(value: unknown, minimum = -Infinity): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < minimum) invalid()
  return value
}
function boundedString(value: unknown, maximum: number): string {
  if (typeof value !== 'string') invalid()
  if (value.length > maximum) throw new Error('presentation_baseline_limit_exceeded')
  return value
}
function budget<T>(value: T): T {
  if (new TextEncoder().encode(JSON.stringify(value)).byteLength > MAX_BYTES)
    throw new Error('presentation_baseline_limit_exceeded')
  return value
}
function supported(version: string): boolean {
  return (
    typeof Office !== 'undefined' &&
    typeof Office.context?.requirements?.isSetSupported === 'function' &&
    Office.context.requirements.isSetSupported('PowerPointApi', version)
  )
}
function runtime(): void {
  if (
    typeof PowerPoint === 'undefined' ||
    typeof PowerPoint.run !== 'function' ||
    typeof Office === 'undefined' ||
    String(Office.context?.host) !== 'PowerPoint' ||
    !supported('1.5')
  )
    throw new Error('office_api_unsupported')
}
async function sync(context: PowerPoint.RequestContext, signal?: AbortSignal): Promise<void> {
  cancelled(signal)
  await context.sync()
  cancelled(signal)
}

/** Read-only native context. No generated-artifact or document-settings dependency. */
export class BrowserPresentationBaselineAdapter implements PresentationBaselineAdapter {
  async readContext(signal?: AbortSignal): Promise<PresentationBaselineContext> {
    cancelled(signal)
    runtime()
    return PowerPoint.run(async (context) => {
      const presentation = context.presentation
      const slides = presentation.slides
      const selectedSlides = presentation.getSelectedSlides()
      const selectedShapes = presentation.getSelectedShapes()
      slides.load({ $top: MAX_SLIDES + 1, id: true })
      selectedSlides.load({ $top: MAX_SLIDES + 1, id: true })
      selectedShapes.load({ $top: MAX_SHAPES + 1, id: true })
      const dimensions = supported('1.10') ? presentation.pageSetup : undefined
      dimensions?.load('slideWidth,slideHeight')
      await sync(context, signal)
      const result: PresentationBaselineContext = {
        slideIds: ids(slides.items, MAX_SLIDES),
        selectedSlideIds: ids(selectedSlides.items, MAX_SLIDES),
        selectedShapeIds: ids(selectedShapes.items, MAX_SHAPES),
      }
      if (
        result.selectedSlideIds.some((value) => !result.slideIds.includes(value)) ||
        (!result.selectedSlideIds.length && result.selectedShapeIds.length)
      )
        invalid()
      if (dimensions) {
        result.slideWidth = number(dimensions.slideWidth, Number.MIN_VALUE)
        result.slideHeight = number(dimensions.slideHeight, Number.MIN_VALUE)
      }
      return budget(result)
    })
  }

  async readPage(slideId: string, signal?: AbortSignal): Promise<PresentationBaselinePage> {
    cancelled(signal)
    id(slideId)
    runtime()
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItem(slideId)
      slide.load('id')
      slide.shapes.load({
        $top: MAX_SHAPES + 1,
        id: true,
        name: true,
        type: true,
        left: true,
        top: true,
        width: true,
        height: true,
      })
      slide.slideMaster.load('id')
      slide.layout.load('id')
      await sync(context, signal)
      if (slide.id !== slideId) invalid()
      const shapes = slide.shapes.items
      ids(shapes, MAX_SHAPES)
      const result: PresentationBaselinePage = {
        slideId,
        masterId: id(slide.slideMaster.id),
        layoutId: id(slide.layout.id),
        shapes: shapes.map((shape) => ({
          id: id(shape.id),
          name: boundedString(shape.name, 1024),
          type: boundedString(shape.type, 128),
          left: number(shape.left),
          top: number(shape.top),
          width: number(shape.width, 0),
          height: number(shape.height, 0),
        })),
      }
      // Placeholder can contain text or a picture. Never assume its textFrame exists.
      const placeholders = shapes.filter((shape) => shape.type === 'Placeholder')
      if (placeholders.length && !supported('1.10')) throw new Error('office_api_unsupported')
      const placeholderFrames = placeholders.map((shape) => ({
        shape,
        frame: shape.getTextFrameOrNullObject(),
      }))
      if (placeholderFrames.length) await sync(context, signal)
      // Table, group, chart, picture and unknown shapes retain geometry only.
      const textShapes = shapes
        .filter((shape) => shape.type === 'TextBox' || shape.type === 'GeometricShape')
        .map((shape) => ({ shape, frame: shape.textFrame }))
      for (const entry of placeholderFrames) {
        if (typeof entry.frame.isNullObject !== 'boolean') invalid()
        if (!entry.frame.isNullObject) textShapes.push(entry)
      }
      for (const { frame } of textShapes) {
        frame.textRange.load('text')
        frame.textRange.font.load('name,size,color,bold,italic,underline')
      }
      if (textShapes.length) await sync(context, signal)
      let textLength = 0
      for (const { shape, frame } of textShapes) {
        const target = result.shapes.find((item) => item.id === shape.id)!
        const range = frame.textRange
        target.text = boundedString(range.text, MAX_TEXT)
        textLength += target.text.length
        if (textLength > MAX_PAGE_TEXT) throw new Error('presentation_baseline_limit_exceeded')
        target.font = {
          name: range.font.name === null ? null : boundedString(range.font.name, 256),
          size: range.font.size === null ? null : number(range.font.size, 0),
          color: range.font.color === null ? null : boundedString(range.font.color, 256),
          bold: range.font.bold === null ? null : typeof range.font.bold === 'boolean' ? range.font.bold : invalid(),
          italic: range.font.italic === null ? null : typeof range.font.italic === 'boolean' ? range.font.italic : invalid(),
          underline: range.font.underline === null ? null : boundedString(range.font.underline, 64),
        }
      }
      return budget(result)
    })
  }
}
