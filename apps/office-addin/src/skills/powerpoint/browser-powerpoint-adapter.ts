import {
  capturePowerPointPackage,
  verifyImportedPowerPointPackage,
  verifyPowerPointPackage,
  type PackageEditResult,
} from './powerpoint-package.js'
import { readUntilConverged } from '../shared/office-write-transaction.js'

export const MAX_POWERPOINT_SHAPES = 1_000
export const MAX_POWERPOINT_TEXT = 12_000
export const MAX_POWERPOINT_RESULT_BYTES = 256 * 1024
export const MAX_POWERPOINT_SNAPSHOT_BASE64 = 8 * 1024 * 1024
export const MAX_POWERPOINT_VERIFY_OVERLAPS = 1_000
export const MAX_POWERPOINT_VERIFY_SLIDES = 20
export const MAX_POWERPOINT_VERIFY_SHAPES = 100
export const MAX_POWERPOINT_VERIFY_OVERFLOWS = 2_000

function uncertainPowerPointState(errorLocation: string, cause?: unknown): Error {
  return Object.assign(
    cause === undefined
      ? new Error('office_state_uncertain')
      : new Error('office_state_uncertain', { cause }),
    { debugInfo: { errorLocation } },
  )
}

function isPowerPointMac(): boolean {
  try {
    return String(Office.context.platform).toLowerCase() === 'mac'
  } catch {
    return false
  }
}

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
  shapesTruncated: boolean
  overflows: Array<{
    shapeId: string
    edge: 'left' | 'top' | 'right' | 'bottom'
    overflowBy: number
  }>
  overlaps: Array<{ shapeAId: string; shapeBId: string; overlapX: number; overlapY: number }>
  overlapsTruncated: boolean
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
  exportSlidePackage(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; base64: string; fingerprint: string }>
  replaceSlidePackage(
    slideIndex: number,
    base64: string,
    applyMaster?: boolean,
    expected?: PackageEditResult,
    signal?: AbortSignal,
  ): Promise<{ slideId: string }>
  executeDeclarative(
    operations: PowerPointDeclarativeOperation[],
    signal?: AbortSignal,
  ): Promise<{ createdShapeIds: string[]; insertedSlideId?: string }>
}

export type PowerPointDeclarativeOperation =
  | { op: 'set_shape_text'; slide_index: number; shape_id: string; text: string }
  | { op: 'duplicate_slide'; slide_index: number }
  | {
      op: 'set_shape_geometry'
      slide_index: number
      shape_id: string
      left: number
      top: number
      width: number
      height: number
    }
  | {
      op: 'add_text_box'
      slide_index: number
      name: string
      text: string
      left: number
      top: number
      width: number
      height: number
    }
  | { op: 'delete_shape'; slide_index: number; shape_id: string }

type RuntimeRecord = Record<string, unknown>

function cancelled(signal?: AbortSignal): void {
  if (signal?.aborted) throw new Error('cancelled')
}

function runtime(minimumVersion: '1.4' | '1.8' | '1.10'): RuntimeRecord {
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
      minimumVersion,
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

function loadSlides(slides: RuntimeRecord): void {
  ;(slides.load as (properties: unknown) => void)({
    $top: MAX_POWERPOINT_VERIFY_SLIDES + 1,
    id: true,
  })
}

function loadShapes(shapes: RuntimeRecord, limit = MAX_POWERPOINT_VERIFY_SHAPES): void {
  ;(shapes.load as (properties: unknown) => void)({
    $top: limit + 1,
    id: true,
    name: true,
    type: true,
    left: true,
    top: true,
    width: true,
    height: true,
  })
}

async function getSlide(
  context: RuntimeRecord,
  slides: RuntimeRecord,
  index: number,
  signal?: AbortSignal,
): Promise<RuntimeRecord> {
  if (typeof slides.getCount !== 'function' || typeof slides.getItemAt !== 'function')
    throw new Error('office_api_unsupported')
  const count = (slides.getCount as () => RuntimeRecord)()
  await sync(context, signal)
  if (!Number.isSafeInteger(count.value) || index < 0 || index >= (count.value as number))
    throw new Error('invalid_tool_input')
  const slide = (slides.getItemAt as (position: number) => RuntimeRecord)(index)
  if (typeof slide.load !== 'function') throw new Error('office_api_unsupported')
  ;(slide.load as (properties: string) => void)('id')
  await sync(context, signal)
  return slide
}

async function getSlideCount(
  context: RuntimeRecord,
  slides: RuntimeRecord,
  signal?: AbortSignal,
): Promise<number> {
  if (typeof slides.getCount !== 'function') throw new Error('office_api_unsupported')
  const count = (slides.getCount as () => RuntimeRecord)()
  await sync(context, signal)
  if (!Number.isSafeInteger(count.value) || (count.value as number) < 0)
    throw new Error('office_read_failed')
  return count.value as number
}

function hash(value: string): string {
  let result = 0x811c9dc5
  for (let index = 0; index < value.length; index += 1) {
    result ^= value.charCodeAt(index)
    result = Math.imul(result, 0x01000193)
  }
  return `${value.length}:${(result >>> 0).toString(16).padStart(8, '0')}`
}

function slideSemanticFingerprint(value: string): string {
  const separator = value.indexOf(':')
  return separator < 0 ? value : value.slice(separator + 1)
}

export class BrowserPowerPointAdapter implements PowerPointAdapter {
  private run<T>(
    minimumVersion: '1.4' | '1.8' | '1.10',
    callback: (context: RuntimeRecord) => Promise<T>,
  ): Promise<T> {
    const powerPoint = runtime(minimumVersion)
    return (powerPoint.run as (callback: (context: RuntimeRecord) => Promise<T>) => Promise<T>)(
      callback,
    )
  }

  async listSlideShapes(slideIndex: number, signal?: AbortSignal): Promise<SlideShapesResult> {
    cancelled(signal)
    return this.run('1.4', async (context) => {
      const presentation = context.presentation as RuntimeRecord
      const slides = presentation.slides as RuntimeRecord
      const slide = await getSlide(context, slides, slideIndex, signal)
      const shapes = slide.shapes as RuntimeRecord
      loadShapes(shapes, MAX_POWERPOINT_SHAPES)
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
    return this.run('1.8', async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      const slide = await getSlide(context, slides, slideIndex, signal)
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
    return this.run('1.4', async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      const slide = await getSlide(context, slides, slideIndex, signal)
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
    return this.run('1.10', async (context) => {
      const presentation = context.presentation as RuntimeRecord
      const slides = presentation.slides as RuntimeRecord
      const pageSetup = presentation.pageSetup as RuntimeRecord
      loadSlides(slides)
      ;(pageSetup.load as (properties: string[]) => void)(['slideWidth', 'slideHeight'])
      await sync(context, signal)
      const slideItems = slides.items as RuntimeRecord[]
      const boundedSlides = slideItems.slice(0, MAX_POWERPOINT_VERIFY_SLIDES)
      for (const slide of boundedSlides) {
        const shapes = slide.shapes as RuntimeRecord
        loadShapes(shapes)
      }
      await sync(context, signal)
      const slideWidth = finite(pageSetup.slideWidth)
      const slideHeight = finite(pageSetup.slideHeight)
      let remainingOverlaps = MAX_POWERPOINT_VERIFY_OVERLAPS
      let remainingOverflows = MAX_POWERPOINT_VERIFY_OVERFLOWS
      const results = boundedSlides.map((slide, slideIndex): SlideVerification => {
        const raw = ((slide.shapes as RuntimeRecord).items as RuntimeRecord[]) ?? []
        const shapesTruncated = raw.length > MAX_POWERPOINT_VERIFY_SHAPES
        const shapes = raw.slice(0, MAX_POWERPOINT_VERIFY_SHAPES).map(shapeInfo)
        const overflows: SlideVerification['overflows'] = []
        for (const shape of shapes) {
          if (shape.left < 0 && remainingOverflows-- > 0)
            overflows.push({ shapeId: shape.id, edge: 'left', overflowBy: -shape.left })
          if (shape.top < 0 && remainingOverflows-- > 0)
            overflows.push({ shapeId: shape.id, edge: 'top', overflowBy: -shape.top })
          if (shape.left + shape.width > slideWidth && remainingOverflows-- > 0)
            overflows.push({
              shapeId: shape.id,
              edge: 'right',
              overflowBy: shape.left + shape.width - slideWidth,
            })
          if (shape.top + shape.height > slideHeight && remainingOverflows-- > 0)
            overflows.push({
              shapeId: shape.id,
              edge: 'bottom',
              overflowBy: shape.top + shape.height - slideHeight,
            })
        }
        const overlaps: SlideVerification['overlaps'] = []
        let overlapsTruncated = remainingOverlaps <= 0
        for (let first = 0; first < shapes.length; first += 1)
          for (let second = first + 1; second < shapes.length; second += 1) {
            if (remainingOverlaps <= 0) {
              overlapsTruncated = true
              break
            }
            const a = shapes[first]
            const b = shapes[second]
            const overlapX = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)
            const overlapY = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top)
            if (overlapX > 0 && overlapY > 0) {
              overlaps.push({ shapeAId: a.id, shapeBId: b.id, overlapX, overlapY })
              remainingOverlaps -= 1
            }
          }
        return {
          slideId: string(slide.id),
          slideIndex,
          shapes,
          shapesTruncated,
          overflows,
          overlaps,
          overlapsTruncated,
        }
      })
      return {
        slideWidth,
        slideHeight,
        slides: results,
        truncated:
          slideItems.length > MAX_POWERPOINT_VERIFY_SLIDES ||
          remainingOverlaps <= 0 ||
          remainingOverflows <= 0,
      }
    })
  }

  async snapshotSlide(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; fingerprint: string }> {
    cancelled(signal)
    return this.run('1.4', async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      const slide = await getSlide(context, slides, slideIndex, signal)
      const shapes = slide.shapes as RuntimeRecord
      if (!shapes || typeof shapes.load !== 'function') throw new Error('office_api_unsupported')
      loadShapes(shapes, MAX_POWERPOINT_SHAPES)
      await sync(context, signal)
      const items = shapes.items as RuntimeRecord[]
      if (!Array.isArray(items) || items.length > MAX_POWERPOINT_SHAPES)
        throw new Error('office_read_failed')
      for (const shape of items) {
        const textFrame = shape.textFrame as RuntimeRecord | undefined
        if (textFrame && typeof textFrame.load === 'function')
          (textFrame.load as (properties: string) => void)('hasText')
      }
      await sync(context, signal)
      for (const shape of items) {
        const textFrame = shape.textFrame as RuntimeRecord | undefined
        const textRange = textFrame?.textRange as RuntimeRecord | undefined
        if (textFrame?.hasText === true && textRange && typeof textRange.load === 'function')
          (textRange.load as (properties: string) => void)('text')
      }
      await sync(context, signal)
      const slideId = string(slide.id)
      const semanticShapes = items
        .map((shape) => {
          const textFrame = shape.textFrame as RuntimeRecord | undefined
          const textRange = textFrame?.textRange as RuntimeRecord | undefined
          return {
            ...shapeInfo(shape),
            text: textFrame?.hasText === true ? string(textRange?.text, MAX_POWERPOINT_TEXT) : '',
          }
        })
        .sort((first, second) => first.id.localeCompare(second.id))
      return { slideId, fingerprint: `${slideId}:${hash(JSON.stringify(semanticShapes))}` }
    })
  }

  async exportSlidePackage(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; base64: string; fingerprint: string }> {
    cancelled(signal)
    return this.run('1.8', async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      const slide = await getSlide(context, slides, slideIndex, signal)
      if (typeof slide.exportAsBase64 !== 'function') throw new Error('office_api_unsupported')
      const exported = (slide.exportAsBase64 as () => RuntimeRecord)()
      await sync(context, signal)
      if (
        typeof exported.value !== 'string' ||
        !exported.value ||
        exported.value.length > MAX_POWERPOINT_SNAPSHOT_BASE64
      )
        throw new Error('office_read_failed')
      const slideId = string(slide.id)
      return { slideId, base64: exported.value, fingerprint: `${slideId}:${hash(exported.value)}` }
    })
  }

  async replaceSlidePackage(
    slideIndex: number,
    base64: string,
    applyMaster = false,
    expected?: PackageEditResult,
    signal?: AbortSignal,
  ): Promise<{ slideId: string }> {
    cancelled(signal)
    if (applyMaster && isPowerPointMac()) throw new Error('office_api_unsupported')
    if (!base64 || base64.length > MAX_POWERPOINT_SNAPSHOT_BASE64)
      throw new Error('office_write_failed')
    return this.run('1.8', async (context) => {
      const presentation = context.presentation as RuntimeRecord
      const slides = presentation.slides as RuntimeRecord
      const beforeCount = await getSlideCount(context, slides, signal)
      const slide = await getSlide(context, slides, slideIndex, signal)
      const previous =
        slideIndex > 0 ? await getSlide(context, slides, slideIndex - 1, signal) : undefined
      if (
        typeof presentation.insertSlidesFromBase64 !== 'function' ||
        typeof slide.delete !== 'function'
      )
        throw new Error('office_api_unsupported')
      const masters = applyMaster
        ? (presentation.slideMasters as RuntimeRecord | undefined)
        : undefined
      if (applyMaster && (!masters || !(slide.layout as RuntimeRecord | undefined)))
        throw new Error('office_api_unsupported')
      const affectedLayouts = new Map<string, number>()
      const originalLayouts = new Map<string, RuntimeRecord>()
      const originalLayoutIds = new Map<string, string>()
      if (applyMaster && masters) {
        const sourceLayout = slide.layout as RuntimeRecord
        ;(masters.load as (properties: string) => void)('items')
        ;(slides.load as (properties: string) => void)('items/id')
        ;(sourceLayout.load as (properties: string) => void)('id')
        await sync(context, signal)
        const masterItems = masters.items as RuntimeRecord[]
        const slideItems = slides.items as RuntimeRecord[]
        for (const master of masterItems) {
          const layouts = master.layouts as RuntimeRecord
          if (typeof layouts.load !== 'function') throw new Error('office_api_unsupported')
          ;(layouts.load as (properties: string) => void)('items/id,items/name')
        }
        for (const item of slideItems) {
          const layout = item.layout as RuntimeRecord
          if (typeof layout?.load !== 'function' || typeof item.applyLayout !== 'function')
            throw new Error('office_api_unsupported')
          ;(layout.load as (properties: string) => void)('id')
        }
        await sync(context, signal)
        const sourceMaster = masterItems.find((master) =>
          (((master.layouts as RuntimeRecord).items as RuntimeRecord[]) ?? []).some(
            (layout) => layout.id === sourceLayout.id,
          ),
        )
        if (!sourceMaster) throw new Error('office_api_unsupported')
        const sourceLayouts =
          ((sourceMaster.layouts as RuntimeRecord).items as RuntimeRecord[]) ?? []
        const sourceOrdinals = new Map(
          sourceLayouts.map((layout, index) => [string(layout.id), index]),
        )
        for (const item of slideItems) {
          const layoutId = string((item.layout as RuntimeRecord).id)
          const ordinal = sourceOrdinals.get(layoutId)
          if (ordinal !== undefined) {
            affectedLayouts.set(string(item.id), ordinal)
            originalLayouts.set(string(item.id), item.layout as RuntimeRecord)
            originalLayoutIds.set(string(item.id), layoutId)
          }
        }
      }
      if (typeof slide.exportAsBase64 !== 'function') throw new Error('office_api_unsupported')
      const original = (slide.exportAsBase64 as () => RuntimeRecord)()
      await sync(context, signal)
      if (
        typeof original.value !== 'string' ||
        !original.value ||
        original.value.length > MAX_POWERPOINT_SNAPSHOT_BASE64
      )
        throw new Error('office_read_failed')
      const originalExpected = await capturePowerPointPackage(original.value, signal)
      const replacementExactProof = await capturePowerPointPackage(base64, signal)
      const originalSlideId = string(slide.id)

      const classifyPackage = async (
        item: RuntimeRecord,
      ): Promise<'original' | 'imported' | 'third'> => {
        if (typeof item.exportAsBase64 !== 'function') return 'third'
        const exported = (item.exportAsBase64 as () => RuntimeRecord)()
        await sync(context)
        if (typeof exported.value !== 'string') return 'third'
        if (await verifyPowerPointPackage(exported.value, originalExpected)) return 'original'
        if (
          expected
            ? await verifyImportedPowerPointPackage(exported.value, expected)
            : await verifyPowerPointPackage(exported.value, replacementExactProof)
        )
          return 'imported'
        return 'third'
      }

      const proveRecovery = async (verifyLayouts: boolean): Promise<void> => {
        if ((await getSlideCount(context, slides)) !== beforeCount)
          throw new Error('office_recovery_failed')
        const restored = await getSlide(context, slides, slideIndex)
        if (typeof restored.exportAsBase64 !== 'function') throw new Error('office_recovery_failed')
        const restoredExport = (restored.exportAsBase64 as () => RuntimeRecord)()
        await sync(context)
        if (
          typeof restoredExport.value !== 'string' ||
          !(await verifyPowerPointPackage(restoredExport.value, originalExpected))
        )
          throw new Error('office_recovery_failed')
        if (!verifyLayouts) return
        if (typeof slides.load !== 'function') throw new Error('office_recovery_failed')
        ;(slides.load as (properties: string) => void)('items/id')
        await sync(context)
        const survivors = ((slides.items as RuntimeRecord[]) ?? []).filter(
          (item) => string(item.id) !== originalSlideId && originalLayoutIds.has(string(item.id)),
        )
        for (const item of survivors) {
          const layout = item.layout as RuntimeRecord
          if (typeof layout?.load !== 'function') throw new Error('office_recovery_failed')
          ;(layout.load as (properties: string) => void)('id')
        }
        await sync(context)
        for (const item of survivors) {
          if (string((item.layout as RuntimeRecord).id) !== originalLayoutIds.get(string(item.id)))
            throw new Error('office_recovery_failed')
        }
      }
      cancelled(signal)
      const insertOptions = {
        formatting: 'KeepSourceFormatting',
        ...(previous ? { targetSlideId: string(previous.id) } : {}),
      }
      ;(
        presentation.insertSlidesFromBase64 as (
          value: string,
          options: { formatting: string; targetSlideId?: string },
        ) => void
      )(base64, insertOptions)
      ;(slide.delete as () => void)()
      try {
        await sync(context, signal)
      } catch (writeError) {
        // A failed Office.js batch may commit a prefix. Recovery is allowed only after proving
        // exact ownership of both the imported candidate and the captured original package.
        const converged = await readUntilConverged({
          read: async () => {
            const currentCount = await getSlideCount(context, slides)
            if (currentCount < 1 || slideIndex >= currentCount)
              return { currentCount, current: undefined, currentKind: 'missing' as const }
            const current = await getSlide(context, slides, slideIndex)
            return { currentCount, current, currentKind: await classifyPackage(current) }
          },
          accept: ({ currentCount, currentKind }) =>
            currentCount !== beforeCount || currentKind !== 'original',
        })
        const { currentCount, current, currentKind } = converged
        if (currentCount < 1 || slideIndex >= currentCount)
          throw new Error('office_state_uncertain', { cause: writeError })
        if (!current) throw new Error('office_state_uncertain', { cause: writeError })
        if (currentKind === 'original') {
          if (currentCount !== beforeCount)
            throw new Error('office_concurrent_change', { cause: writeError })
          await proveRecovery(false)
          throw new Error('office_write_failed', { cause: writeError })
        }
        if (currentKind !== 'imported')
          throw new Error('office_concurrent_change', { cause: writeError })
        if (currentCount === beforeCount + 1) {
          const survivingOriginal = await getSlide(context, slides, slideIndex + 1)
          if ((await classifyPackage(survivingOriginal)) !== 'original')
            throw new Error('office_concurrent_change', { cause: writeError })
          if (
            (await classifyPackage(current)) !== 'imported' ||
            typeof current.delete !== 'function'
          )
            throw new Error('office_concurrent_change', { cause: writeError })
          ;(current.delete as () => void)()
          await sync(context)
        } else if (currentCount === beforeCount) {
          if (
            (await classifyPackage(current)) !== 'imported' ||
            typeof current.delete !== 'function'
          )
            throw new Error('office_concurrent_change', { cause: writeError })
          ;(
            presentation.insertSlidesFromBase64 as (
              value: string,
              options: { formatting: string; targetSlideId?: string },
            ) => void
          )(original.value, insertOptions)
          ;(current.delete as () => void)()
          await sync(context)
        } else {
          throw new Error('office_state_uncertain', { cause: writeError })
        }
        await proveRecovery(false)
        throw new Error('office_write_failed', { cause: writeError })
      }
      if ((await getSlideCount(context, slides, signal)) !== beforeCount)
        throw new Error('office_verify_failed')
      const inserted = await getSlide(context, slides, slideIndex, signal)
      if (!expected || typeof inserted.exportAsBase64 !== 'function')
        throw new Error('office_api_unsupported')
      const imported = (inserted.exportAsBase64 as () => RuntimeRecord)()
      await sync(context, signal)
      if (
        typeof imported.value !== 'string' ||
        !(await verifyImportedPowerPointPackage(imported.value, expected, signal))
      ) {
        // Verification failure is not proof that the current slide is still owned by this
        // transaction. Never overwrite a possible concurrent edit with the captured package.
        throw uncertainPowerPointState('PowerPoint.replaceSlidePackage.packageImportVerify')
      }
      if (applyMaster) {
        try {
          const insertedLayout = inserted.layout as RuntimeRecord | undefined
          if (
            !masters ||
            !insertedLayout ||
            typeof masters.load !== 'function' ||
            typeof insertedLayout.load !== 'function'
          )
            throw new Error('office_api_unsupported')
          ;(masters.load as (properties: string) => void)('items')
          ;(slides.load as (properties: string) => void)('items')
          ;(insertedLayout.load as (properties: string) => void)('id,name')
          await sync(context, signal)
          const masterItems = masters.items as RuntimeRecord[]
          const slideItems = slides.items as RuntimeRecord[]
          for (const master of masterItems) {
            const layouts = master.layouts as RuntimeRecord
            ;(layouts.load as (properties: string) => void)('items/id,items/name')
          }
          for (const item of slideItems) {
            const layout = item.layout as RuntimeRecord
            ;(layout.load as (properties: string) => void)('id,name')
          }
          await sync(context, signal)
          const primary = masterItems.find((master) =>
            (((master.layouts as RuntimeRecord).items as RuntimeRecord[]) ?? []).some(
              (layout) => layout.id === insertedLayout.id,
            ),
          )
          if (!primary) throw new Error('office_verify_failed')
          const primaryLayouts = ((primary.layouts as RuntimeRecord).items as RuntimeRecord[]) ?? []
          cancelled(signal)
          for (const item of slideItems) {
            const intendedOrdinal = affectedLayouts.get(string(item.id))
            const intended =
              intendedOrdinal === undefined ? undefined : primaryLayouts[intendedOrdinal]
            if (intended && typeof item.applyLayout === 'function')
              (item.applyLayout as (target: RuntimeRecord) => void)(intended)
          }
          await sync(context, signal)
          for (const item of slideItems) {
            if (affectedLayouts.has(string(item.id))) {
              ;((item.layout as RuntimeRecord).load as (properties: string) => void)('id,name')
            }
          }
          await sync(context, signal)
          const primaryLayoutIds = new Set(primaryLayouts.map((layout) => string(layout.id)))
          for (const item of slideItems) {
            const expectedOrdinal = affectedLayouts.get(string(item.id))
            if (expectedOrdinal === undefined) continue
            const layout = item.layout as RuntimeRecord
            if (
              string(layout.id) !== string(primaryLayouts[expectedOrdinal]?.id) ||
              !primaryLayoutIds.has(string(layout.id))
            )
              throw new Error('office_verify_failed')
          }
        } catch (error) {
          // Layout and package writes span multiple Office batches. Without an atomic compare-
          // and-set primitive, restoring here could overwrite a concurrent user edit.
          throw uncertainPowerPointState('PowerPoint.replaceSlidePackage.layoutApplyVerify', error)
        }
      }
      return { slideId: string(inserted.id) }
    })
  }

  async executeDeclarative(
    operations: PowerPointDeclarativeOperation[],
    signal?: AbortSignal,
  ): Promise<{ createdShapeIds: string[]; insertedSlideId?: string }> {
    cancelled(signal)
    if (operations.length === 1 && operations[0].op === 'duplicate_slide') {
      const inserted = await this.duplicateSlide(operations[0].slide_index, signal)
      return { createdShapeIds: [], insertedSlideId: inserted.slideId }
    }
    const deleteTargets = new Set(
      operations
        .filter((operation) => operation.op === 'delete_shape')
        .map((operation) => `${operation.slide_index}/${operation.shape_id}`),
    )
    if (
      operations.some(
        (operation) =>
          'shape_id' in operation &&
          operation.op !== 'delete_shape' &&
          deleteTargets.has(`${operation.slide_index}/${operation.shape_id}`),
      )
    )
      throw new Error('invalid_tool_input')
    const repeatedTargets = new Set<string>()
    for (const operation of operations) {
      if (operation.op !== 'set_shape_text' && operation.op !== 'set_shape_geometry') continue
      const key = `${operation.slide_index}/${operation.shape_id}/${operation.op}`
      if (repeatedTargets.has(key)) throw new Error('invalid_tool_input')
      repeatedTargets.add(key)
    }
    return this.run('1.8', async (context) => {
      const presentation = context.presentation as RuntimeRecord
      const slides = presentation.slides as RuntimeRecord
      const queued: Array<() => void> = []
      const createdShapes: RuntimeRecord[] = []
      type TrackedMutation =
        | {
            kind: 'text'
            target: RuntimeRecord
            before?: string
            after: string
          }
        | {
            kind: 'geometry'
            target: RuntimeRecord
            before?: [number, number, number, number]
            after: [number, number, number, number]
          }
      const trackedByTarget = new Map<string, TrackedMutation>()
      let hasUnrecoverableMutation = false
      for (const operation of operations) {
        const slide = await getSlide(context, slides, operation.slide_index, signal)
        if (
          operation.op === 'set_shape_text' ||
          operation.op === 'set_shape_geometry' ||
          operation.op === 'delete_shape'
        ) {
          const shapes = slide.shapes as RuntimeRecord
          if (typeof shapes.getItem !== 'function') throw new Error('office_api_unsupported')
          const shape = (shapes.getItem as (id: string) => RuntimeRecord)(operation.shape_id)
          if (operation.op === 'set_shape_text') {
            const textRange = (shape.textFrame as RuntimeRecord | undefined)?.textRange as
              RuntimeRecord | undefined
            if (!textRange || typeof textRange.load !== 'function')
              throw new Error('office_api_unsupported')
            ;(textRange.load as (properties: string) => void)('text')
            const key = `${operation.slide_index}/${operation.shape_id}/text`
            const existing = trackedByTarget.get(key)
            if (existing?.kind === 'text') existing.after = operation.text
            else
              trackedByTarget.set(key, { kind: 'text', target: textRange, after: operation.text })
            queued.push(() => {
              textRange.text = operation.text
            })
          } else if (operation.op === 'set_shape_geometry') {
            if (typeof shape.load !== 'function') throw new Error('office_api_unsupported')
            ;(shape.load as (properties: string) => void)('left,top,width,height')
            const key = `${operation.slide_index}/${operation.shape_id}/geometry`
            const after: [number, number, number, number] = [
              operation.left,
              operation.top,
              operation.width,
              operation.height,
            ]
            const existing = trackedByTarget.get(key)
            if (existing?.kind === 'geometry') existing.after = after
            else trackedByTarget.set(key, { kind: 'geometry', target: shape, after })
            queued.push(() => {
              shape.left = operation.left
              shape.top = operation.top
              shape.width = operation.width
              shape.height = operation.height
            })
          } else {
            if (typeof shape.delete !== 'function') throw new Error('office_api_unsupported')
            hasUnrecoverableMutation = true
            queued.push(() => (shape.delete as () => void)())
          }
        } else if (operation.op === 'add_text_box') {
          const shapes = slide.shapes as RuntimeRecord
          if (typeof shapes.addTextBox !== 'function') throw new Error('office_api_unsupported')
          queued.push(() => {
            const created = (
              shapes.addTextBox as (text: string, options: Record<string, number>) => RuntimeRecord
            )(operation.text, {
              left: operation.left,
              top: operation.top,
              width: operation.width,
              height: operation.height,
            })
            created.name = operation.name
            if (typeof created.load !== 'function') throw new Error('office_api_unsupported')
            ;(created.load as (properties: string) => void)('id')
            createdShapes.push(created)
          })
          hasUnrecoverableMutation = true
        } else {
          if (
            typeof slide.exportAsBase64 !== 'function' ||
            typeof presentation.insertSlidesFromBase64 !== 'function'
          )
            throw new Error('office_api_unsupported')
          const exported = (slide.exportAsBase64 as () => RuntimeRecord)()
          await sync(context, signal)
          if (
            typeof exported.value !== 'string' ||
            !exported.value ||
            exported.value.length > MAX_POWERPOINT_SNAPSHOT_BASE64
          )
            throw new Error('office_write_failed')
          queued.push(() =>
            (
              presentation.insertSlidesFromBase64 as (
                value: string,
                options: { targetSlideId: string },
              ) => void
            )(exported.value as string, { targetSlideId: string(slide.id) }),
          )
          hasUnrecoverableMutation = true
        }
      }
      await sync(context, signal)
      for (const mutation of trackedByTarget.values()) {
        if (mutation.kind === 'text')
          mutation.before = string(mutation.target.text, MAX_POWERPOINT_TEXT)
        else
          mutation.before = [
            finite(mutation.target.left),
            finite(mutation.target.top),
            finite(mutation.target.width),
            finite(mutation.target.height),
          ]
      }
      cancelled(signal)
      for (const write of queued) write()
      try {
        await sync(context, signal)
        if (trackedByTarget.size > 0) {
          const applied = await readUntilConverged({
            signal,
            read: async () => {
              for (const mutation of trackedByTarget.values()) {
                if (mutation.kind === 'text')
                  (mutation.target.load as (properties: string) => void)('text')
                else (mutation.target.load as (properties: string) => void)('left,top,width,height')
              }
              await sync(context, signal)
              return [...trackedByTarget.values()].every((mutation) => {
                if (mutation.kind === 'text')
                  return string(mutation.target.text, MAX_POWERPOINT_TEXT) === mutation.after
                return [
                  finite(mutation.target.left),
                  finite(mutation.target.top),
                  finite(mutation.target.width),
                  finite(mutation.target.height),
                ].every((value, index) => Math.abs(value - mutation.after[index]) <= 0.01)
              })
            },
            accept: Boolean,
          })
          if (!applied) throw new Error('office_verify_failed')
        }
      } catch {
        const observed = await readUntilConverged({
          read: async (): Promise<'before' | 'attributable' | 'applied' | 'third'> => {
            for (const mutation of trackedByTarget.values()) {
              if (mutation.kind === 'text')
                (mutation.target.load as (properties: string) => void)('text')
              else (mutation.target.load as (properties: string) => void)('left,top,width,height')
            }
            await sync(context)
            let beforeCount = 0
            let afterCount = 0
            for (const mutation of trackedByTarget.values()) {
              if (mutation.kind === 'text') {
                const current = string(mutation.target.text, MAX_POWERPOINT_TEXT)
                if (current === mutation.before) beforeCount += 1
                else if (current === mutation.after) afterCount += 1
                else return 'third'
              } else {
                const current = [
                  finite(mutation.target.left),
                  finite(mutation.target.top),
                  finite(mutation.target.width),
                  finite(mutation.target.height),
                ]
                if (current.every((value, index) => value === mutation.before?.[index]))
                  beforeCount += 1
                else if (
                  current.every((value, index) => Math.abs(value - mutation.after[index]) <= 0.01)
                )
                  afterCount += 1
                else return 'third'
              }
            }
            if (afterCount === trackedByTarget.size) return 'applied'
            if (beforeCount === trackedByTarget.size) return 'before'
            return 'attributable'
          },
          accept: (state) => state === 'applied',
        })
        if (observed === 'third') throw new Error('office_concurrent_change')
        if (observed === 'applied' && !hasUnrecoverableMutation)
          return { createdShapeIds: createdShapes.map((shape) => string(shape.id)) }
        if (observed === 'before' && !hasUnrecoverableMutation)
          throw new Error(signal?.aborted ? 'cancelled' : 'office_write_failed')
        // We cannot safely synthesize an arbitrary deleted shape or prove ownership of an
        // inserted object after a rejected batch. Surface uncertainty instead of overwriting.
        if (hasUnrecoverableMutation) throw new Error('office_state_uncertain')
        // Office.js has no atomic compare-and-set for shapes. An attributable partial state is
        // uncertain; restoring it could overwrite a user edit between classification and sync.
        throw new Error('office_state_uncertain')
      }
      return { createdShapeIds: createdShapes.map((shape) => string(shape.id)) }
    })
  }

  async editSlideText(
    slideIndex: number,
    shapeId: string,
    value: string,
    signal?: AbortSignal,
  ): Promise<void> {
    cancelled(signal)
    await this.run('1.4', async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      const slide = await getSlide(context, slides, slideIndex, signal)
      const shapes = slide.shapes as RuntimeRecord
      if (typeof shapes.getItem !== 'function') throw new Error('office_api_unsupported')
      const shape = (shapes.getItem as (id: string) => RuntimeRecord)(shapeId)
      const textRange = (shape.textFrame as RuntimeRecord | undefined)?.textRange as
        RuntimeRecord | undefined
      if (!textRange || typeof textRange.load !== 'function')
        throw new Error('office_api_unsupported')
      ;(textRange.load as (properties: string) => void)('text')
      await sync(context, signal)
      const before = string(textRange.text, MAX_POWERPOINT_TEXT)
      cancelled(signal)
      textRange.text = value
      try {
        await sync(context, signal)
        const applied = await readUntilConverged({
          signal,
          read: async () => {
            ;(textRange.load as (properties: string) => void)('text')
            await sync(context, signal)
            return string(textRange.text, MAX_POWERPOINT_TEXT)
          },
          accept: (current) => current === value,
        })
        if (applied !== value) throw new Error('office_verify_failed')
      } catch {
        // A rejected Office.js sync may still have committed the assignment. Reconcile the
        // semantic target before deciding whether the write failed or cancellation won.
        const current = await readUntilConverged({
          read: async () => {
            ;(textRange.load as (properties: string) => void)('text')
            await sync(context)
            return string(textRange.text, MAX_POWERPOINT_TEXT)
          },
          accept: (observed) => observed === value,
        })
        if (current === before)
          throw new Error(signal?.aborted ? 'cancelled' : 'office_write_failed')
        if (current !== value) throw new Error('office_concurrent_change')
        return
      }
    })
  }

  async duplicateSlide(slideIndex: number, signal?: AbortSignal): Promise<{ slideId: string }> {
    cancelled(signal)
    const before = await this.snapshotSlide(slideIndex, signal)
    let followingBefore: { slideId: string; fingerprint: string } | undefined
    try {
      followingBefore = await this.snapshotSlide(slideIndex + 1, signal)
    } catch (error) {
      if (!(error instanceof Error) || error.message !== 'invalid_tool_input') throw error
    }
    let sourcePackage: { slideId: string; base64: string; fingerprint: string }
    try {
      sourcePackage = await this.exportSlidePackage(slideIndex, signal)
    } catch (error) {
      if (error instanceof Error && error.message === 'office_read_failed')
        throw new Error('office_write_failed', { cause: error })
      throw error
    }
    if (sourcePackage.slideId !== before.slideId) throw new Error('office_concurrent_change')
    const preWriteSource = await this.snapshotSlide(slideIndex, signal)
    if (preWriteSource.fingerprint !== before.fingerprint)
      throw new Error('office_concurrent_change')
    let writeError: unknown
    try {
      await this.run('1.8', async (context) => {
        const presentation = context.presentation as RuntimeRecord
        const slides = presentation.slides as RuntimeRecord
        const slide = await getSlide(context, slides, slideIndex, signal)
        if (typeof presentation.insertSlidesFromBase64 !== 'function')
          throw new Error('office_api_unsupported')
        cancelled(signal)
        ;(
          presentation.insertSlidesFromBase64 as (
            value: string,
            options: { targetSlideId: string },
          ) => void
        )(sourcePackage.base64, { targetSlideId: string(slide.id) })
        await sync(context, signal)
      })
    } catch (error) {
      writeError = error
    }
    const source = await this.snapshotSlide(slideIndex)
    const sourceChanged = source.fingerprint !== before.fingerprint
    const following = await readUntilConverged({
      read: async () => {
        try {
          return await this.snapshotSlide(slideIndex + 1)
        } catch (readError) {
          if (readError instanceof Error && readError.message === 'invalid_tool_input')
            return undefined
          throw new Error('office_state_uncertain', { cause: readError })
        }
      },
      accept: (candidate) =>
        followingBefore
          ? candidate !== undefined && candidate.fingerprint !== followingBefore.fingerprint
          : candidate !== undefined,
    })
    if (
      (!followingBefore && !following) ||
      (followingBefore && following?.fingerprint === followingBefore.fingerprint)
    )
      throw new Error(
        signal?.aborted
          ? 'cancelled'
          : writeError
            ? 'office_write_failed'
            : 'office_state_uncertain',
        {
          cause: writeError,
        },
      )
    if (
      !following ||
      following.slideId === followingBefore?.slideId ||
      slideSemanticFingerprint(following.fingerprint) !==
        slideSemanticFingerprint(before.fingerprint)
    )
      throw new Error('office_concurrent_change', { cause: writeError })
    if (!writeError && !signal?.aborted && !sourceChanged) return { slideId: following.slideId }
    const sourcePackageProof = await capturePowerPointPackage(sourcePackage.base64)
    const ownsInsertedPackage = await readUntilConverged({
      read: async () => {
        try {
          const exported = await this.exportSlidePackage(slideIndex + 1)
          return (
            exported.slideId === following.slideId &&
            (await verifyPowerPointPackage(exported.base64, sourcePackageProof))
          )
        } catch {
          return false
        }
      },
      accept: Boolean,
    })
    if (!ownsInsertedPackage) throw new Error('office_concurrent_change', { cause: writeError })
    if (sourceChanged) throw new Error('office_concurrent_change', { cause: writeError })
    const stableSource = await this.snapshotSlide(slideIndex)
    if (stableSource.fingerprint !== before.fingerprint)
      throw new Error('office_concurrent_change', { cause: writeError })
    // The exact inserted package is durably present. Report the applied result even when Stop or
    // a rejected sync raced the commit; deleting it would reopen a destructive TOCTOU window.
    return { slideId: following.slideId }
  }
}
