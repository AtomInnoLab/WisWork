export const MAX_POWERPOINT_SLIDES = 500
export const MAX_POWERPOINT_SHAPES = 1_000
export const MAX_POWERPOINT_TEXT = 12_000
export const MAX_POWERPOINT_RESULT_BYTES = 256 * 1024

export interface PowerPointShape {
  id: string
  name: string
  type: string
  left: number
  top: number
  width: number
  height: number
}

export interface SlideShapesResult {
  slideId: string
  slideIndex: number
  shapes: PowerPointShape[]
}

export interface SlideTextResult {
  slideId: string
  shapeId: string
  text: string
  paragraphs: string[]
}

export interface SlideVerification {
  slideId: string
  slideIndex: number
  shapes: PowerPointShape[]
  overflows: Array<{
    shapeId: string
    edge: 'left' | 'top' | 'right' | 'bottom'
    overflowBy: number
  }>
  overlaps: Array<{ shapeAId: string; shapeBId: string; overlapX: number; overlapY: number }>
}

export interface VerifySlidesResult {
  slideWidth: number
  slideHeight: number
  slides: SlideVerification[]
  truncated?: boolean
}

export interface PowerPointAdapter {
  screenshotSlide(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ base64: string; mime: 'image/png' }>
  listSlideShapes(slideIndex: number, signal?: AbortSignal): Promise<SlideShapesResult>
  readSlideText(slideIndex: number, shapeId: string, signal?: AbortSignal): Promise<SlideTextResult>
  verifySlides(signal?: AbortSignal): Promise<VerifySlidesResult>
  snapshotSlide(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; fingerprint: string }>
  editSlideText(
    slideIndex: number,
    shapeId: string,
    text: string,
    signal?: AbortSignal,
  ): Promise<void>
  duplicateSlide(slideIndex: number, signal?: AbortSignal): Promise<{ slideId: string }>
}

type RuntimeRecord = Record<string, unknown>

function cancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}

function runtime(): RuntimeRecord {
  const root = globalThis as unknown as RuntimeRecord
  const office = root.Office as RuntimeRecord | undefined
  const powerPoint = root.PowerPoint as RuntimeRecord | undefined
  const context = office?.context as RuntimeRecord | undefined
  const requirements = context?.requirements as RuntimeRecord | undefined
  const supports = requirements?.isSetSupported
  if (
    !office ||
    !powerPoint ||
    context?.host !== 'PowerPoint' ||
    typeof supports !== 'function' ||
    !(supports as (name: string, version: string) => boolean).call(
      requirements,
      'PowerPointApi',
      '1.2',
    ) ||
    typeof powerPoint.run !== 'function'
  )
    throw new Error('office_api_unsupported')
  return powerPoint
}

async function sync(context: RuntimeRecord, signal?: AbortSignal): Promise<void> {
  cancelled(signal)
  if (typeof context.sync !== 'function') throw new Error('office_api_unsupported')
  await (context.sync as () => Promise<void>)()
  cancelled(signal)
}

function finite(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : 0
}

function string(value: unknown, maximum = 256): string {
  return typeof value === 'string' ? value.slice(0, maximum) : ''
}

function shapeInfo(value: RuntimeRecord): PowerPointShape {
  return {
    id: string(value.id),
    name: string(value.name),
    type: string(value.type, 64),
    left: finite(value.left),
    top: finite(value.top),
    width: Math.max(0, finite(value.width)),
    height: Math.max(0, finite(value.height)),
  }
}

function assertSlide(slides: RuntimeRecord, index: number): RuntimeRecord {
  const items = slides.items as RuntimeRecord[] | undefined
  if (!items || index < 0 || index >= items.length) throw new Error('invalid_tool_input')
  return items[index]
}

function hash(value: string): string {
  let result = 0x811c9dc5
  for (let index = 0; index < value.length; index += 1) {
    result ^= value.charCodeAt(index)
    result = Math.imul(result, 0x01000193)
  }
  return `${value.length}:${(result >>> 0).toString(16).padStart(8, '0')}`
}

export class BrowserPowerPointAdapter implements PowerPointAdapter {
  private run<T>(callback: (context: RuntimeRecord) => Promise<T>): Promise<T> {
    const powerPoint = runtime()
    return (powerPoint.run as (callback: (context: RuntimeRecord) => Promise<T>) => Promise<T>)(
      callback,
    )
  }

  async listSlideShapes(slideIndex: number, signal?: AbortSignal): Promise<SlideShapesResult> {
    cancelled(signal)
    return this.run(async (context) => {
      const presentation = context.presentation as RuntimeRecord
      const slides = presentation.slides as RuntimeRecord
      ;(slides.load as (properties: string) => void)('items/id')
      await sync(context, signal)
      const slide = assertSlide(slides, slideIndex)
      const shapes = slide.shapes as RuntimeRecord
      ;(shapes.load as (properties: string) => void)(
        'items/id,items/name,items/type,items/left,items/top,items/width,items/height',
      )
      await sync(context, signal)
      const items = shapes.items as RuntimeRecord[]
      if (items.length > MAX_POWERPOINT_SHAPES) throw new Error('office_read_failed')
      return { slideId: string(slide.id), slideIndex, shapes: items.map(shapeInfo) }
    })
  }

  async screenshotSlide(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ base64: string; mime: 'image/png' }> {
    cancelled(signal)
    return this.run(async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      ;(slides.load as (properties: string) => void)('items/id')
      await sync(context, signal)
      const slide = assertSlide(slides, slideIndex)
      if (typeof slide.getImageAsBase64 !== 'function') throw new Error('office_api_unsupported')
      const image = (slide.getImageAsBase64 as (options: { width: number }) => RuntimeRecord)({
        width: 960,
      })
      await sync(context, signal)
      if (typeof image.value !== 'string') throw new Error('office_read_failed')
      return { base64: image.value, mime: 'image/png' }
    })
  }

  async readSlideText(
    slideIndex: number,
    shapeId: string,
    signal?: AbortSignal,
  ): Promise<SlideTextResult> {
    cancelled(signal)
    return this.run(async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      ;(slides.load as (properties: string) => void)('items/id')
      await sync(context, signal)
      const slide = assertSlide(slides, slideIndex)
      const shapes = slide.shapes as RuntimeRecord
      if (typeof shapes.getItem !== 'function') throw new Error('office_api_unsupported')
      const shape = (shapes.getItem as (id: string) => RuntimeRecord)(shapeId)
      const textFrame = shape.textFrame as RuntimeRecord | undefined
      const textRange = textFrame?.textRange as RuntimeRecord | undefined
      if (!textRange || typeof textRange.load !== 'function')
        throw new Error('office_api_unsupported')
      ;(textRange.load as (properties: string) => void)('text')
      await sync(context, signal)
      const value = string(textRange.text, MAX_POWERPOINT_TEXT)
      return { slideId: string(slide.id), shapeId, text: value, paragraphs: value.split(/\r?\n/) }
    })
  }

  async verifySlides(signal?: AbortSignal): Promise<VerifySlidesResult> {
    cancelled(signal)
    return this.run(async (context) => {
      const presentation = context.presentation as RuntimeRecord
      const slides = presentation.slides as RuntimeRecord
      const pageSetup = presentation.pageSetup as RuntimeRecord
      ;(slides.load as (properties: string) => void)('items/id')
      ;(pageSetup.load as (properties: string[]) => void)(['slideWidth', 'slideHeight'])
      await sync(context, signal)
      const slideItems = slides.items as RuntimeRecord[]
      const boundedSlides = slideItems.slice(0, MAX_POWERPOINT_SLIDES)
      for (const slide of boundedSlides) {
        const shapes = slide.shapes as RuntimeRecord
        ;(shapes.load as (properties: string) => void)(
          'items/id,items/name,items/type,items/left,items/top,items/width,items/height',
        )
      }
      await sync(context, signal)
      const slideWidth = finite(pageSetup.slideWidth)
      const slideHeight = finite(pageSetup.slideHeight)
      const results = boundedSlides.map((slide, slideIndex): SlideVerification => {
        const raw = ((slide.shapes as RuntimeRecord).items as RuntimeRecord[]) ?? []
        if (raw.length > MAX_POWERPOINT_SHAPES) throw new Error('office_read_failed')
        const shapes = raw.map(shapeInfo)
        const overflows: SlideVerification['overflows'] = []
        for (const shape of shapes) {
          if (shape.left < 0)
            overflows.push({ shapeId: shape.id, edge: 'left', overflowBy: -shape.left })
          if (shape.top < 0)
            overflows.push({ shapeId: shape.id, edge: 'top', overflowBy: -shape.top })
          if (shape.left + shape.width > slideWidth)
            overflows.push({
              shapeId: shape.id,
              edge: 'right',
              overflowBy: shape.left + shape.width - slideWidth,
            })
          if (shape.top + shape.height > slideHeight)
            overflows.push({
              shapeId: shape.id,
              edge: 'bottom',
              overflowBy: shape.top + shape.height - slideHeight,
            })
        }
        const overlaps: SlideVerification['overlaps'] = []
        for (let first = 0; first < shapes.length; first += 1)
          for (let second = first + 1; second < shapes.length; second += 1) {
            const a = shapes[first]
            const b = shapes[second]
            const overlapX = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)
            const overlapY = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top)
            if (overlapX > 0 && overlapY > 0)
              overlaps.push({ shapeAId: a.id, shapeBId: b.id, overlapX, overlapY })
          }
        return { slideId: string(slide.id), slideIndex, shapes, overflows, overlaps }
      })
      return {
        slideWidth,
        slideHeight,
        slides: results,
        truncated: slideItems.length > MAX_POWERPOINT_SLIDES,
      }
    })
  }

  async snapshotSlide(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; fingerprint: string }> {
    const result = await this.listSlideShapes(slideIndex, signal)
    return {
      slideId: result.slideId,
      fingerprint: `${result.slideId}:${hash(JSON.stringify(result.shapes))}`,
    }
  }

  async editSlideText(
    slideIndex: number,
    shapeId: string,
    value: string,
    signal?: AbortSignal,
  ): Promise<void> {
    cancelled(signal)
    await this.run(async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      ;(slides.load as (properties: string) => void)('items/id')
      await sync(context, signal)
      const slide = assertSlide(slides, slideIndex)
      const shapes = slide.shapes as RuntimeRecord
      if (typeof shapes.getItem !== 'function') throw new Error('office_api_unsupported')
      const shape = (shapes.getItem as (id: string) => RuntimeRecord)(shapeId)
      const textRange = (shape.textFrame as RuntimeRecord | undefined)?.textRange as
        RuntimeRecord | undefined
      if (!textRange) throw new Error('office_api_unsupported')
      cancelled(signal)
      textRange.text = value
      await sync(context, signal)
    })
  }

  async duplicateSlide(slideIndex: number, signal?: AbortSignal): Promise<{ slideId: string }> {
    cancelled(signal)
    return this.run(async (context) => {
      const presentation = context.presentation as RuntimeRecord
      const slides = presentation.slides as RuntimeRecord
      ;(slides.load as (properties: string) => void)('items/id')
      await sync(context, signal)
      const slide = assertSlide(slides, slideIndex)
      if (
        typeof slide.exportAsBase64 !== 'function' ||
        typeof presentation.insertSlidesFromBase64 !== 'function'
      )
        throw new Error('office_api_unsupported')
      const exported = (slide.exportAsBase64 as () => RuntimeRecord)()
      await sync(context, signal)
      if (typeof exported.value !== 'string') throw new Error('office_write_failed')
      cancelled(signal)
      ;(
        presentation.insertSlidesFromBase64 as (
          value: string,
          options: { targetSlideId: string },
        ) => void
      )(exported.value, { targetSlideId: string(slide.id) })
      await sync(context, signal)
      return { slideId: string(slide.id) }
    })
  }
}
