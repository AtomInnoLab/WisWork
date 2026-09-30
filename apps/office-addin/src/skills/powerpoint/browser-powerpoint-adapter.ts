import { masterOperationKey, masterStateValuesFingerprint } from './presentation-master-program.js'
import {
  parsePowerPointStyleDependencies,
  type PowerPointStyleDependencies,
} from './presentation-style-dependencies.js'
import {
  capturePowerPointPackage,
  inspectPowerPointPictureMediaBatch,
  presentationPackageDigest,
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

export interface PowerPointPageInspection {
  slideId: string
  slideWidth: number
  slideHeight: number
  shapes: PowerPointShape[]
  shapesTruncated: boolean
  overflows: SlideVerification['overflows']
  overlaps: SlideVerification['overlaps']
  overlapsTruncated: boolean
  screenshot: { mime: 'image/png'; base64: string; renderer?: 'libreoffice' }
}

export interface VerifySlidesResult {
  slideWidth: number
  slideHeight: number
  slides: SlideVerification[]
  truncated?: boolean
}

export interface PowerPointMasterState {
  masters: Array<{
    id: string
    name: string
    background: {
      type: string
      color?: string
      transparency?: number
      gradientType?: string
      pattern?: string
      foregroundColor?: string
      backgroundColor?: string
      pictureTransparency?: number
    }
    themeColors: Record<string, string>
    layouts: Array<{
      id: string
      name: string
      isMasterBackgroundFollowed: boolean
      areBackgroundGraphicsHidden: boolean
      background: { type: string }
    }>
  }>
}

export type PowerPointMasterOperation =
  | {
      op: 'set_master_background'
      master_id: string
      fill:
        | { type: 'solid'; color: string; transparency: number }
        | { type: 'gradient'; gradient_type: string }
        | { type: 'pattern'; pattern: string; foreground_color: string; background_color: string }
        | { type: 'picture_or_texture'; image_base64: string; transparency: number }
    }
  | { op: 'set_master_theme_color'; master_id: string; theme_color: string; color: string }
  | {
      op: 'set_layout_background_following'
      master_id: string
      layout_id: string
      follow_master: boolean
      show_master_graphics: boolean
    }

export interface PowerPointMasterExecutionPreimage {
  before: PowerPointMasterState
  operations: PowerPointMasterOperation[]
  slideIds: string[]
  dependencies: PowerPointStyleDependencies
}

export interface PresentationPageGeometry {
  left: number
  top: number
  width: number
  height: number
}

export interface PresentationTextRangeSnapshot {
  slideId: string
  shapeId: string
  start: number
  length: number
  fullText: string
  text: string
  font: {
    name: string | null
    size: number | null
    color: string | null
    bold: boolean | null
    italic: boolean | null
    underline: string | null
  }
}

export interface PowerPointAdapter {
  readSlideOrder?(signal?: AbortSignal): Promise<string[]>
  exportPresentationPagePackage?(
    slideId: string,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; slideIds: string[]; base64: string }>
  inspectSlidePictureFingerprints?(
    slideId: string,
    shapeIds: string[],
    signal?: AbortSignal,
  ): Promise<{
    slideId: string
    slideIds: string[]
    fingerprints: Record<string, string>
    mediaDigests: Record<string, string>
  }>
  readPresentationPageGeometry?(
    slideId: string,
    shapeId: string,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; shapeId: string; geometry: PresentationPageGeometry }>
  editPresentationPageGeometry?(
    slideId: string,
    shapeId: string,
    geometry: PresentationPageGeometry,
    expectedGeometry: PresentationPageGeometry,
    signal?: AbortSignal,
  ): Promise<void>
  listPresentationPageShapes?(
    slideId: string,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; shapes: PowerPointShape[]; shapesTruncated: boolean }>
  readPresentationPageText?(
    slideId: string,
    shapeId: string,
    signal?: AbortSignal,
  ): Promise<SlideTextResult>
  editPresentationPageText?(
    slideId: string,
    shapeId: string,
    text: string,
    expectedText: string,
    signal?: AbortSignal,
  ): Promise<void>
  readPresentationPageTextRange?(
    slideId: string,
    shapeId: string,
    start: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<PresentationTextRangeSnapshot>
  editPresentationPageTextRange?(
    expected: PresentationTextRangeSnapshot,
    after: string,
    signal?: AbortSignal,
  ): Promise<void>
  readPresentationTableCell?(
    slideId: string,
    shapeId: string,
    rowIndex: number,
    columnIndex: number,
    signal?: AbortSignal,
  ): Promise<{
    slideId: string
    shapeId: string
    rowIndex: number
    columnIndex: number
    text: string
    rowCount: number
    columnCount: number
  }>
  editPresentationTableCell?(
    slideId: string,
    shapeId: string,
    rowIndex: number,
    columnIndex: number,
    text: string,
    expectedText: string,
    signal?: AbortSignal,
  ): Promise<void>
  inspectPresentationPage?(
    slideId: string,
    signal?: AbortSignal,
    fallbackBase64?: string,
  ): Promise<PowerPointPageInspection>
  inspectStyleDependencies?(signal?: AbortSignal): Promise<PowerPointStyleDependencies>
  inspectSlideMasters(signal?: AbortSignal): Promise<PowerPointMasterState>
  executeMasterOperations(
    operations: PowerPointMasterOperation[],
    signal?: AbortSignal,
    preimage?: PowerPointMasterExecutionPreimage,
    beforeWrite?: () => Promise<void>,
    writeGuard?: () => void,
  ): Promise<void>
  screenshotSlide(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; base64: string; mime: 'image/png' }>
  listSlideShapes(slideIndex: number, signal?: AbortSignal): Promise<SlideShapesResult>
  readSlideText(slideIndex: number, shapeId: string, signal?: AbortSignal): Promise<SlideTextResult>
  readSlideTable(slideIndex: number, shapeId: string, signal?: AbortSignal): Promise<string[][]>
  verifySlides(signal?: AbortSignal): Promise<VerifySlidesResult>
  snapshotSlide(
    slideIndex: number,
    signal?: AbortSignal,
    includeShapes?: boolean,
  ): Promise<{
    slideId: string
    fingerprint: string
    shapes?: Array<
      PowerPointShape & {
        text: string
        tableValues?: string[][]
        rotation?: number
        altTextTitle?: string
        altTextDescription?: string
      }
    >
  }>
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
    preimage?: { slideId: string; packageDigest: string; slideIds?: string[] },
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
      op: 'add_geometric_shape'
      slide_index: number
      name: string
      shape: 'rect' | 'ellipse' | 'roundRect'
      left: number
      top: number
      width: number
      height: number
      fill: string
      lineColor: string
    }
  | {
      op: 'add_native_table'
      slide_index: number
      name: string
      rows: string[][]
      left: number
      top: number
      width: number
      height: number
      fontFace: string
      fontSize: number
      color: string
      borderColor?: string
      cellMargin?: number
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
      fontFace?: string
      fontSize?: number
      color?: string
      bold?: boolean
      align?: 'left' | 'center' | 'right'
      margin?: number
      verticalAlignment?: 'top' | 'middle' | 'bottom'
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

function explicitCanvasBackground(
  shape: PowerPointShape,
  slideWidth: number,
  slideHeight: number,
): boolean {
  const tolerance = 0.01
  return (
    shape.type === 'GeometricShape' &&
    /^(?:background|背景)(?:[\s_-]|$)/i.test(shape.name) &&
    Math.abs(shape.left) <= tolerance &&
    Math.abs(shape.top) <= tolerance &&
    Math.abs(shape.width - slideWidth) <= tolerance &&
    Math.abs(shape.height - slideHeight) <= tolerance
  )
}

function loadSlides(slides: RuntimeRecord): void {
  ;(slides.load as (properties: unknown) => void)({
    $top: MAX_POWERPOINT_VERIFY_SLIDES + 1,
    id: true,
  })
}

function loadShapes(
  shapes: RuntimeRecord,
  limit = MAX_POWERPOINT_VERIFY_SHAPES,
  strong = false,
): void {
  ;(shapes.load as (properties: unknown) => void)({
    $top: limit + 1,
    id: true,
    name: true,
    type: true,
    left: true,
    top: true,
    width: true,
    height: true,
    ...(strong ? { rotation: true, altTextTitle: true, altTextDescription: true } : {}),
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

/** Read every slide ID with its count in one Office batch; never use bounded QA windows. */
async function readCompleteSlideOrder(
  context: RuntimeRecord,
  slides: RuntimeRecord,
  signal?: AbortSignal,
): Promise<string[]> {
  if (typeof slides.load !== 'function' || typeof slides.getCount !== 'function')
    throw new Error('office_api_unsupported')
  const count = (slides.getCount as () => RuntimeRecord)()
  ;(slides.load as (properties: string) => void)('items/id')
  await sync(context, signal)
  if (
    !Number.isSafeInteger(count.value) ||
    (count.value as number) < 1 ||
    !Array.isArray(slides.items) ||
    slides.items.length !== count.value
  )
    throw new Error('office_read_failed')
  const ids = (slides.items as RuntimeRecord[]).map((item) => item?.id)
  if (
    ids.some(
      (id) =>
        typeof id !== 'string' ||
        !id.trim() ||
        id.length > 256 ||
        [...id].some(
          (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
        ),
    ) ||
    new Set(ids).size !== ids.length
  )
    throw new Error('office_read_failed')
  return ids as string[]
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

/** Validate the Office-produced screenshot envelope without claiming a visual QA pass. */
export function validatePowerPointPageScreenshot(value: unknown): string {
  const limit = 2 * 1024 * 1024
  if (
    typeof value !== 'string' ||
    !value.length ||
    value.length > Math.ceil(limit / 3) * 4 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)
  )
    throw new Error('office_read_failed')
  let binary: string
  try {
    binary = atob(value)
  } catch {
    throw new Error('office_read_failed')
  }
  if (
    binary.length > limit ||
    binary.length < 45 ||
    btoa(binary) !== value ||
    binary.slice(0, 8) !== '\x89PNG\r\n\x1a\n' ||
    binary.slice(12, 16) !== 'IHDR' ||
    binary.slice(-8, -4) !== 'IEND'
  )
    throw new Error('office_read_failed')
  const uint32 = (offset: number) =>
    binary.charCodeAt(offset) * 0x1000000 +
    binary.charCodeAt(offset + 1) * 0x10000 +
    binary.charCodeAt(offset + 2) * 0x100 +
    binary.charCodeAt(offset + 3)
  const width = uint32(16),
    height = uint32(20)
  if (
    uint32(8) !== 13 ||
    !width ||
    !height ||
    width > 8192 ||
    height > 8192 ||
    width * height > 16_000_000
  )
    throw new Error('office_read_failed')
  return value
}

function pageId(value: unknown): asserts value is string {
  if (typeof value !== 'string' || !value.length || value.length > 256)
    throw new Error('invalid_tool_input')
}
function boundedPageText(value: unknown): string {
  if (typeof value !== 'string' || value.length > MAX_POWERPOINT_TEXT)
    throw new Error('office_read_failed')
  return value
}
async function getPageById(
  context: RuntimeRecord,
  id: string,
  signal?: AbortSignal,
): Promise<RuntimeRecord> {
  const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
  if (typeof slides?.getItem !== 'function') throw new Error('office_api_unsupported')
  const slide = (slides.getItem as (id: string) => RuntimeRecord)(id)
  if (typeof slide?.load !== 'function') throw new Error('office_api_unsupported')
  ;(slide.load as (properties: string) => void)('id')
  await sync(context, signal)
  if (slide.id !== id) throw new Error('office_read_failed')
  return slide
}
async function pageTextRange(
  context: RuntimeRecord,
  slide: RuntimeRecord,
  shapeId: string,
  signal?: AbortSignal,
): Promise<RuntimeRecord> {
  const shapes = slide.shapes as RuntimeRecord
  if (typeof shapes?.getItem !== 'function') throw new Error('office_api_unsupported')
  const shape = (shapes.getItem as (id: string) => RuntimeRecord)(shapeId)
  const range = (shape?.textFrame as RuntimeRecord | undefined)?.textRange as
    RuntimeRecord | undefined
  if (typeof shape?.load !== 'function' || typeof range?.load !== 'function')
    throw new Error('office_api_unsupported')
  ;(shape.load as (properties: string) => void)('id')
  await sync(context, signal)
  if (shape.id !== shapeId) throw new Error('office_read_failed')
  return range
}
function validTextSpan(fullText: string, start: number, length: number): void {
  if (
    !Number.isSafeInteger(start) ||
    start < 0 ||
    !Number.isSafeInteger(length) ||
    length < 1 ||
    length > 128 ||
    start + length > fullText.length ||
    /[\uD800-\uDFFF]/u.test(fullText) ||
    /[\uD800-\uDFFF]/u.test(fullText.slice(start, start + length)) ||
    /[\r\n]/.test(fullText.slice(start, start + length))
  )
    throw new Error('office_api_unsupported')
}
async function readTextSpan(
  context: RuntimeRecord,
  fullRange: RuntimeRecord,
  start: number,
  length: number,
  signal?: AbortSignal,
) {
  ;(fullRange.load as (properties: string) => void)('text')
  await sync(context, signal)
  const fullText = boundedPageText(fullRange.text)
  validTextSpan(fullText, start, length)
  if (typeof fullRange.getSubstring !== 'function') throw new Error('office_api_unsupported')
  const range = (fullRange.getSubstring as (start: number, length: number) => RuntimeRecord)(
    start,
    length,
  )
  const font = range?.font as RuntimeRecord | undefined
  if (typeof range?.load !== 'function' || typeof font?.load !== 'function')
    throw new Error('office_api_unsupported')
  ;(range.load as (properties: string) => void)('text')
  ;(font.load as (properties: string) => void)('name,size,color,bold,italic,underline')
  await sync(context, signal)
  const text = boundedPageText(range.text)
  if (text !== fullText.slice(start, start + length)) throw new Error('office_read_failed')
  const fontValue = {
    name: font.name,
    size: font.size,
    color: font.color,
    bold: font.bold,
    italic: font.italic,
    underline: font.underline,
  }
  if (
    (fontValue.name !== null &&
      (typeof fontValue.name !== 'string' || fontValue.name.length > 256)) ||
    (fontValue.size !== null &&
      (typeof fontValue.size !== 'number' ||
        !Number.isFinite(fontValue.size) ||
        fontValue.size < 0)) ||
    (fontValue.color !== null &&
      (typeof fontValue.color !== 'string' || fontValue.color.length > 256)) ||
    (fontValue.bold !== null && typeof fontValue.bold !== 'boolean') ||
    (fontValue.italic !== null && typeof fontValue.italic !== 'boolean') ||
    (fontValue.underline !== null &&
      (typeof fontValue.underline !== 'string' || fontValue.underline.length > 64))
  )
    throw new Error('office_read_failed')
  return { range, fullText, text, font: fontValue as PresentationTextRangeSnapshot['font'] }
}
async function pageTableCell(
  context: RuntimeRecord,
  slideId: string,
  shapeId: string,
  rowIndex: number,
  columnIndex: number,
  signal?: AbortSignal,
): Promise<{ cell: RuntimeRecord; rowCount: number; columnCount: number }> {
  if (
    !Number.isSafeInteger(rowIndex) ||
    rowIndex < 0 ||
    !Number.isSafeInteger(columnIndex) ||
    columnIndex < 0
  )
    throw new Error('invalid_tool_input')
  const slide = await getPageById(context, slideId, signal)
  const shapes = slide.shapes as RuntimeRecord
  if (typeof shapes?.getItem !== 'function') throw new Error('office_api_unsupported')
  const shape = (shapes.getItem as (id: string) => RuntimeRecord)(shapeId)
  if (typeof shape?.load !== 'function') throw new Error('office_api_unsupported')
  ;(shape.load as (properties: string) => void)('id,type')
  await sync(context, signal)
  if (shape.id !== shapeId) throw new Error('office_read_failed')
  if (shape.type !== 'Table') throw new Error('office_api_unsupported')
  if (typeof shape.getTable !== 'function') throw new Error('office_api_unsupported')
  const table = (shape.getTable as () => RuntimeRecord)()
  if (typeof table?.load !== 'function' || typeof table.getCellOrNullObject !== 'function')
    throw new Error('office_api_unsupported')
  ;(table.load as (properties: string) => void)('rowCount,columnCount')
  await sync(context, signal)
  const { rowCount, columnCount } = table
  if (
    typeof rowCount !== 'number' ||
    typeof columnCount !== 'number' ||
    !Number.isSafeInteger(rowCount) ||
    !Number.isSafeInteger(columnCount) ||
    rowCount < 1 ||
    columnCount < 1
  )
    throw new Error('office_read_failed')
  if (rowIndex >= rowCount || columnIndex >= columnCount) throw new Error('invalid_tool_input')
  const cell = (table.getCellOrNullObject as (row: number, column: number) => RuntimeRecord)(
    rowIndex,
    columnIndex,
  )
  if (typeof cell?.load !== 'function') throw new Error('office_api_unsupported')
  ;(cell.load as (properties: string) => void)('rowIndex,columnIndex,rowCount,columnCount,text')
  await sync(context, signal)
  if (cell.isNullObject || cell.rowCount !== 1 || cell.columnCount !== 1)
    throw new Error('office_api_unsupported')
  if (cell.rowIndex !== rowIndex || cell.columnIndex !== columnIndex)
    throw new Error('office_read_failed')
  return { cell, rowCount, columnCount }
}
async function writeTextRange(
  context: RuntimeRecord,
  textRange: RuntimeRecord,
  value: string,
  signal?: AbortSignal,
  expectedText?: string,
): Promise<void> {
  ;(textRange.load as (properties: string) => void)('text')
  await sync(context, signal)
  const readText = () =>
    expectedText === undefined
      ? string(textRange.text, MAX_POWERPOINT_TEXT)
      : boundedPageText(textRange.text)
  const before = readText()
  if (expectedText !== undefined && before !== expectedText)
    throw new Error('office_concurrent_change')
  cancelled(signal)
  textRange.text = value
  try {
    await sync(context, signal)
    const applied = await readUntilConverged({
      signal,
      read: async () => {
        ;(textRange.load as (properties: string) => void)('text')
        await sync(context, signal)
        return readText()
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
        return readText()
      },
      accept: (observed) => observed === value,
    })
    if (current === before) throw new Error(signal?.aborted ? 'cancelled' : 'office_write_failed')
    if (current !== value) throw new Error('office_concurrent_change')
    return
  }
}

const geometryFields = ['left', 'top', 'width', 'height'] as const
function validPageGeometry(value: unknown): value is PresentationPageGeometry {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const item = value as Record<string, unknown>
  return (
    Object.keys(item).length === 4 &&
    geometryFields.every((key) => Object.hasOwn(item, key)) &&
    geometryFields.every(
      (key) =>
        typeof item[key] === 'number' &&
        Number.isFinite(item[key]) &&
        Math.abs(item[key] as number) <= 100000,
    ) &&
    (item.width as number) >= 0 &&
    (item.height as number) >= 0
  )
}
function shapeGeometry(shape: RuntimeRecord): PresentationPageGeometry {
  const result = { left: shape.left, top: shape.top, width: shape.width, height: shape.height }
  if (!validPageGeometry(result)) throw new Error('office_read_failed')
  return result
}
export const nativeGeometryEditable = (type: unknown): boolean =>
  type === 'TextBox' || type === 'GeometricShape' || type === 'Image' || type === 'Line'
async function geometryShape(
  context: RuntimeRecord,
  slideId: string,
  shapeId: string,
  signal?: AbortSignal,
): Promise<RuntimeRecord> {
  const slide = await getPageById(context, slideId, signal)
  const shapes = slide.shapes as RuntimeRecord
  if (typeof shapes?.getItem !== 'function') throw new Error('office_api_unsupported')
  const shape = (shapes.getItem as (id: string) => RuntimeRecord)(shapeId)
  if (typeof shape?.load !== 'function') throw new Error('office_api_unsupported')
  ;(shape.load as (properties: string[]) => void)(['id', ...geometryFields])
  await sync(context, signal)
  if (shape.id !== shapeId) throw new Error('office_read_failed')
  return shape
}
function geometryApplied(
  current: PresentationPageGeometry,
  target: PresentationPageGeometry,
): boolean {
  return geometryFields.every((key) => Math.abs(current[key] - target[key]) <= 0.01)
}

async function inspectMasterState(
  context: RuntimeRecord,
  signal?: AbortSignal,
  beforeFinalSync?: () => void,
): Promise<PowerPointMasterState> {
  const presentation = context.presentation as RuntimeRecord
  const masters = presentation.slideMasters as RuntimeRecord
  if (!masters || typeof masters.load !== 'function') throw new Error('office_api_unsupported')
  ;(masters.load as (properties: string) => void)(
    'items/id,items/name,items/layouts/items/id,items/layouts/items/name',
  )
  await sync(context, signal)
  const masterItems = (masters.items as RuntimeRecord[]) ?? []
  if (masterItems.length > 32) throw new Error('office_read_failed')
  const themeSlots = [
    'Accent1',
    'Accent2',
    'Accent3',
    'Accent4',
    'Accent5',
    'Accent6',
    'Dark1',
    'Dark2',
    'Light1',
    'Light2',
    'Hyperlink',
    'FollowedHyperlink',
  ]
  const pending = masterItems.map((master) => {
    const fill = (master.background as RuntimeRecord)?.fill as RuntimeRecord | undefined
    if (!fill || typeof fill.load !== 'function') throw new Error('office_api_unsupported')
    ;(fill.load as (properties: string) => void)('type')
    const solid =
      typeof fill.getSolidFillOrNullObject === 'function'
        ? (fill.getSolidFillOrNullObject as () => RuntimeRecord)()
        : undefined
    if (solid) (solid.load as (properties: string[]) => void)(['color', 'transparency'])
    const gradient =
      typeof fill.getGradientFillOrNullObject === 'function'
        ? (fill.getGradientFillOrNullObject as () => RuntimeRecord)()
        : undefined
    if (gradient) (gradient.load as (properties: string[]) => void)(['type'])
    const pattern =
      typeof fill.getPatternFillOrNullObject === 'function'
        ? (fill.getPatternFillOrNullObject as () => RuntimeRecord)()
        : undefined
    if (pattern)
      (pattern.load as (properties: string[]) => void)([
        'pattern',
        'foregroundColor',
        'backgroundColor',
      ])
    const picture =
      typeof fill.getPictureOrTextureFillOrNullObject === 'function'
        ? (fill.getPictureOrTextureFillOrNullObject as () => RuntimeRecord)()
        : undefined
    if (picture) (picture.load as (properties: string[]) => void)(['transparency'])
    const scheme = master.themeColorScheme as RuntimeRecord
    const colors = Object.fromEntries(
      themeSlots.map((slot) => [
        slot,
        (scheme.getThemeColor as (slot: string) => RuntimeRecord)(slot),
      ]),
    )
    const layouts = ((master.layouts as RuntimeRecord)?.items as RuntimeRecord[]) ?? []
    if (layouts.length > 128) throw new Error('office_read_failed')
    for (const layout of layouts) {
      const background = layout.background as RuntimeRecord
      ;(background.load as (properties: string[]) => void)([
        'isMasterBackgroundFollowed',
        'areBackgroundGraphicsHidden',
      ])
      const layoutFill = background.fill as RuntimeRecord
      if (!layoutFill || typeof layoutFill.load !== 'function')
        throw new Error('office_api_unsupported')
      ;(layoutFill.load as (properties: string) => void)('type')
    }
    return { master, fill, solid, gradient, pattern, picture, colors, layouts }
  })
  beforeFinalSync?.()
  await sync(context, signal)
  return {
    masters: pending.map(
      ({ master, fill, solid, gradient, pattern, picture, colors, layouts }) => ({
        id: string(master.id),
        name: string(master.name),
        background: {
          type: string(fill.type, 64),
          ...(solid && !solid.isNullObject && typeof solid.color === 'string'
            ? { color: solid.color }
            : {}),
          ...(solid && !solid.isNullObject && typeof solid.transparency === 'number'
            ? { transparency: solid.transparency }
            : {}),
          ...(gradient && !gradient.isNullObject && typeof gradient.type === 'string'
            ? { gradientType: gradient.type }
            : {}),
          ...(pattern && !pattern.isNullObject && typeof pattern.pattern === 'string'
            ? { pattern: pattern.pattern }
            : {}),
          ...(pattern && !pattern.isNullObject && typeof pattern.foregroundColor === 'string'
            ? { foregroundColor: pattern.foregroundColor }
            : {}),
          ...(pattern && !pattern.isNullObject && typeof pattern.backgroundColor === 'string'
            ? { backgroundColor: pattern.backgroundColor }
            : {}),
          ...(picture && !picture.isNullObject && typeof picture.transparency === 'number'
            ? { pictureTransparency: picture.transparency }
            : {}),
        },
        themeColors: Object.fromEntries(
          Object.entries(colors).map(([slot, result]) => [
            slot,
            string((result as RuntimeRecord).value, 64),
          ]),
        ),
        layouts: layouts.map((layout) => {
          const background = layout.background as RuntimeRecord
          return {
            id: string(layout.id),
            name: string(layout.name),
            isMasterBackgroundFollowed: Boolean(background.isMasterBackgroundFollowed),
            areBackgroundGraphicsHidden: Boolean(background.areBackgroundGraphicsHidden),
            background: { type: string((background.fill as RuntimeRecord).type, 64) },
          }
        }),
      }),
    ),
  }
}

export class BrowserPowerPointAdapter implements PowerPointAdapter {
  private async runScreenshot<T>(
    minimumVersion: '1.8' | '1.10',
    callback: (context: RuntimeRecord) => Promise<T>,
    signal?: AbortSignal,
  ): Promise<T> {
    for (let attempt = 0; ; attempt++) {
      cancelled(signal)
      try {
        return await this.run(minimumVersion, callback)
      } catch (error) {
        cancelled(signal)
        // A fresh PowerPoint.run context is required after a rejected host batch.
        if (
          attempt > 0 ||
          !error ||
          typeof error !== 'object' ||
          !['ActivityLimitReached', 'Timeout'].includes(String((error as { code?: unknown }).code))
        )
          throw error
      }
    }
  }

  private run<T>(
    minimumVersion: '1.4' | '1.8' | '1.10',
    callback: (context: RuntimeRecord) => Promise<T>,
  ): Promise<T> {
    const powerPoint = runtime(minimumVersion)
    return (powerPoint.run as (callback: (context: RuntimeRecord) => Promise<T>) => Promise<T>)(
      callback,
    )
  }

  async inspectStyleDependencies(signal?: AbortSignal): Promise<PowerPointStyleDependencies> {
    cancelled(signal)
    return this.run('1.4', async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      if (typeof slides?.load !== 'function') throw new Error('office_api_unsupported')
      ;(slides.load as (properties: string) => void)(
        'items/id,items/slideMaster/id,items/layout/id',
      )
      await sync(context, signal)
      if (!Array.isArray(slides.items)) throw new Error('office_read_failed')
      return parsePowerPointStyleDependencies({
        slides: (slides.items as RuntimeRecord[]).map((slide) => ({
          slideId: slide.id,
          masterId: (slide.slideMaster as RuntimeRecord | undefined)?.id,
          layoutId: (slide.layout as RuntimeRecord | undefined)?.id,
        })),
      })
    })
  }

  async inspectSlideMasters(signal?: AbortSignal): Promise<PowerPointMasterState> {
    cancelled(signal)
    return this.run('1.10', (context) => inspectMasterState(context, signal))
  }

  async executeMasterOperations(
    operations: PowerPointMasterOperation[],
    signal?: AbortSignal,
    preimage?: PowerPointMasterExecutionPreimage,
    beforeWrite?: () => Promise<void>,
    writeGuard?: () => void,
  ): Promise<void> {
    const ownedOperations = structuredClone(operations),
      ownedPreimage = preimage && structuredClone(preimage)
    cancelled(signal)
    await this.run('1.10', async (context) => {
      const masters = (context.presentation as RuntimeRecord).slideMasters as RuntimeRecord
      if (typeof masters.getItem !== 'function') throw new Error('office_api_unsupported')
      await beforeWrite?.()
      cancelled(signal)
      if (ownedPreimage) {
        const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
        if (typeof slides.load !== 'function' || typeof slides.getCount !== 'function')
          throw new Error('office_api_unsupported')
        let count: RuntimeRecord | undefined
        const actual = await inspectMasterState(context, signal, () => {
          count = (slides.getCount as () => RuntimeRecord)()
          ;(slides.load as (properties: string) => void)(
            'items/id,items/slideMaster/id,items/layout/id',
          )
          ;(masters.load as (properties: string) => void)(
            'items/id,items/name,items/layouts/items/id,items/layouts/items/name',
          )
        })
        if (
          !Array.isArray(slides.items) ||
          count?.value !== slides.items.length ||
          !Number.isSafeInteger(count?.value) ||
          (count!.value as number) < 1
        )
          throw new Error('proposal_stale')
        const order = (slides.items as RuntimeRecord[]).map((slide) => slide.id)
        const dependencies = parsePowerPointStyleDependencies({
          slides: (slides.items as RuntimeRecord[]).map((slide) => ({
            slideId: slide.id,
            masterId: (slide.slideMaster as RuntimeRecord | undefined)?.id,
            layoutId: (slide.layout as RuntimeRecord | undefined)?.id,
          })),
        })
        if (
          JSON.stringify(order) !== JSON.stringify(ownedPreimage.slideIds) ||
          JSON.stringify(dependencies) !==
            JSON.stringify(parsePowerPointStyleDependencies(ownedPreimage.dependencies)) ||
          masterStateValuesFingerprint(actual) !==
            masterStateValuesFingerprint(ownedPreimage.before) ||
          JSON.stringify(
            (masters.items as RuntimeRecord[]).map((master) => ({
              id: master.id,
              layouts: ((master.layouts as RuntimeRecord).items as RuntimeRecord[]).map(
                (layout) => layout.id,
              ),
            })),
          ) !==
            JSON.stringify(
              actual.masters.map((master) => ({
                id: master.id,
                layouts: master.layouts.map((layout) => layout.id),
              })),
            ) ||
          ownedOperations.some(
            (op) =>
              !ownedPreimage.operations.some(
                (expected) => masterOperationKey(op) === masterOperationKey(expected),
              ),
          )
        )
          throw new Error('proposal_stale')
      }
      writeGuard?.()
      cancelled(signal)
      for (const operation of ownedOperations) {
        cancelled(signal)
        const master = (masters.getItem as (id: string) => RuntimeRecord)(operation.master_id)
        if (operation.op === 'set_master_theme_color') {
          const scheme = master.themeColorScheme as RuntimeRecord
          ;(scheme.setThemeColor as (slot: string, color: string) => void)(
            operation.theme_color,
            operation.color,
          )
        } else if (operation.op === 'set_layout_background_following') {
          const layouts = master.layouts as RuntimeRecord
          const layout = (layouts.getItem as (id: string) => RuntimeRecord)(operation.layout_id)
          const background = layout.background as RuntimeRecord
          background.isMasterBackgroundFollowed = operation.follow_master
          background.areBackgroundGraphicsHidden = !operation.show_master_graphics
        } else {
          const fill = (master.background as RuntimeRecord).fill as RuntimeRecord
          if (operation.fill.type === 'solid')
            (fill.setSolidFill as (options: unknown) => void)({
              color: operation.fill.color,
              transparency: operation.fill.transparency,
            })
          else if (operation.fill.type === 'gradient')
            (fill.setGradientFill as (options: unknown) => void)({
              type: operation.fill.gradient_type,
            })
          else if (operation.fill.type === 'pattern')
            (fill.setPatternFill as (options: unknown) => void)({
              pattern: operation.fill.pattern,
              foregroundColor: operation.fill.foreground_color,
              backgroundColor: operation.fill.background_color,
            })
          else
            (fill.setPictureOrTextureFill as (options: unknown) => void)({
              imageBase64: operation.fill.image_base64,
              transparency: operation.fill.transparency,
            })
        }
      }
      await sync(context, signal)
    })
  }

  async readPresentationPageGeometry(
    slideId: string,
    shapeId: string,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; shapeId: string; geometry: PresentationPageGeometry }> {
    cancelled(signal)
    pageId(slideId)
    pageId(shapeId)
    return this.run('1.10', async (context) => {
      const shape = await geometryShape(context, slideId, shapeId, signal)
      return { slideId, shapeId, geometry: shapeGeometry(shape) }
    })
  }

  async editPresentationPageGeometry(
    slideId: string,
    shapeId: string,
    geometry: PresentationPageGeometry,
    expectedGeometry: PresentationPageGeometry,
    signal?: AbortSignal,
  ): Promise<void> {
    cancelled(signal)
    pageId(slideId)
    pageId(shapeId)
    if (!validPageGeometry(geometry) || !validPageGeometry(expectedGeometry))
      throw new Error('invalid_tool_input')
    const target = { ...geometry },
      expected = { ...expectedGeometry }
    await this.run('1.10', async (context) => {
      const shape = await geometryShape(context, slideId, shapeId, signal)
      const before = shapeGeometry(shape)
      if (!geometryFields.every((key) => before[key] === expected[key]))
        throw new Error('office_concurrent_change')
      cancelled(signal)
      try {
        for (const key of geometryFields) shape[key] = target[key]
      } catch {
        // Do not sync merely to inspect a failed setter: that could dispatch the partial
        // batch still queued in this context. Its final host state remains uncertain.
        throw new Error('office_state_uncertain')
      }
      // Stop can also arrive from a setter before the batch has ever been dispatched.
      if (signal?.aborted) throw new Error('office_state_uncertain')
      try {
        await sync(context, signal)
      } catch {
        /* A rejected sync can still have applied some or all queued fields. */
      }
      let observed: PresentationPageGeometry
      try {
        observed = await readUntilConverged({
          read: async () => {
            ;(shape.load as (properties: string[]) => void)(['id', ...geometryFields])
            await sync(context)
            if (shape.id !== shapeId) throw new Error('office_read_failed')
            return shapeGeometry(shape)
          },
          accept: (current) => geometryApplied(current, target),
        })
      } catch {
        throw new Error('office_state_uncertain')
      }
      if (geometryApplied(observed, target)) return
      if (geometryFields.every((key) => observed[key] === before[key]))
        throw new Error(signal?.aborted ? 'cancelled' : 'office_write_failed')
      const partial = geometryFields.every(
        (key) => observed[key] === before[key] || Math.abs(observed[key] - target[key]) <= 0.01,
      )
      throw new Error(partial ? 'office_state_uncertain' : 'office_concurrent_change')
    })
  }

  async listPresentationPageShapes(
    slideId: string,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; shapes: PowerPointShape[]; shapesTruncated: boolean }> {
    cancelled(signal)
    pageId(slideId)
    return this.run('1.10', async (context) => {
      const slide = await getPageById(context, slideId, signal)
      const shapes = slide.shapes as RuntimeRecord
      if (typeof shapes?.load !== 'function') throw new Error('office_api_unsupported')
      loadShapes(shapes)
      await sync(context, signal)
      if (!Array.isArray(shapes.items)) throw new Error('office_read_failed')
      const items = (shapes.items as RuntimeRecord[]).slice(0, MAX_POWERPOINT_VERIFY_SHAPES)
      for (const item of items)
        if (!item || typeof item.id !== 'string' || !item.id.length || item.id.length > 256)
          throw new Error('office_read_failed')
      return {
        slideId,
        shapes: items.map(shapeInfo),
        shapesTruncated: shapes.items.length > MAX_POWERPOINT_VERIFY_SHAPES,
      }
    })
  }

  async readPresentationPageText(
    slideId: string,
    shapeId: string,
    signal?: AbortSignal,
  ): Promise<SlideTextResult> {
    cancelled(signal)
    pageId(slideId)
    pageId(shapeId)
    return this.run('1.10', async (context) => {
      const slide = await getPageById(context, slideId, signal)
      const range = await pageTextRange(context, slide, shapeId, signal)
      ;(range.load as (properties: string) => void)('text')
      await sync(context, signal)
      const text = boundedPageText(range.text)
      return { slideId, shapeId, text, paragraphs: text.split(/\r?\n/) }
    })
  }

  async editPresentationPageText(
    slideId: string,
    shapeId: string,
    text: string,
    expectedText: string,
    signal?: AbortSignal,
  ): Promise<void> {
    cancelled(signal)
    pageId(slideId)
    pageId(shapeId)
    if (
      typeof text !== 'string' ||
      text.length > MAX_POWERPOINT_TEXT ||
      typeof expectedText !== 'string' ||
      expectedText.length > MAX_POWERPOINT_TEXT
    )
      throw new Error('invalid_tool_input')
    await this.run('1.10', async (context) => {
      const slide = await getPageById(context, slideId, signal)
      const range = await pageTextRange(context, slide, shapeId, signal)
      await writeTextRange(context, range, text, signal, expectedText)
    })
  }

  async readPresentationPageTextRange(
    slideId: string,
    shapeId: string,
    start: number,
    length: number,
    signal?: AbortSignal,
  ): Promise<PresentationTextRangeSnapshot> {
    cancelled(signal)
    pageId(slideId)
    pageId(shapeId)
    return this.run('1.10', async (context) => {
      const slide = await getPageById(context, slideId, signal)
      const fullRange = await pageTextRange(context, slide, shapeId, signal)
      const observed = await readTextSpan(context, fullRange, start, length, signal)
      return {
        slideId,
        shapeId,
        start,
        length,
        fullText: observed.fullText,
        text: observed.text,
        font: observed.font,
      }
    })
  }

  async editPresentationPageTextRange(
    expected: PresentationTextRangeSnapshot,
    after: string,
    signal?: AbortSignal,
  ): Promise<void> {
    cancelled(signal)
    pageId(expected.slideId)
    pageId(expected.shapeId)
    if (
      typeof after !== 'string' ||
      after.length !== expected.length ||
      after === expected.text ||
      /[\r\n]/.test(after) ||
      /[\uD800-\uDFFF]/u.test(after) ||
      typeof expected.fullText !== 'string' ||
      expected.fullText.length > MAX_POWERPOINT_TEXT ||
      Object.values(expected.font).some((value) => value === null)
    )
      throw new Error('invalid_tool_input')
    validTextSpan(expected.fullText, expected.start, expected.length)
    if (expected.fullText.slice(expected.start, expected.start + expected.length) !== expected.text)
      throw new Error('invalid_tool_input')
    await this.run('1.10', async (context) => {
      const slide = await getPageById(context, expected.slideId, signal)
      const fullRange = await pageTextRange(context, slide, expected.shapeId, signal)
      const before = await readTextSpan(context, fullRange, expected.start, expected.length, signal)
      if (
        before.fullText !== expected.fullText ||
        JSON.stringify(before.font) !== JSON.stringify(expected.font)
      )
        throw new Error('office_concurrent_change')
      await writeTextRange(context, before.range, after, signal, expected.text)
      const observed = await readTextSpan(
        context,
        fullRange,
        expected.start,
        expected.length,
        signal,
      )
      const target =
        expected.fullText.slice(0, expected.start) +
        after +
        expected.fullText.slice(expected.start + expected.length)
      if (
        observed.fullText !== target ||
        observed.text !== after ||
        JSON.stringify(observed.font) !== JSON.stringify(expected.font)
      )
        throw new Error('office_verify_failed')
    })
  }

  async readPresentationTableCell(
    slideId: string,
    shapeId: string,
    rowIndex: number,
    columnIndex: number,
    signal?: AbortSignal,
  ): Promise<{
    slideId: string
    shapeId: string
    rowIndex: number
    columnIndex: number
    text: string
    rowCount: number
    columnCount: number
  }> {
    cancelled(signal)
    pageId(slideId)
    pageId(shapeId)
    return this.run('1.8', async (context) => {
      const { cell, rowCount, columnCount } = await pageTableCell(
        context,
        slideId,
        shapeId,
        rowIndex,
        columnIndex,
        signal,
      )
      return {
        slideId,
        shapeId,
        rowIndex,
        columnIndex,
        text: boundedPageText(cell.text),
        rowCount,
        columnCount,
      }
    })
  }

  async editPresentationTableCell(
    slideId: string,
    shapeId: string,
    rowIndex: number,
    columnIndex: number,
    text: string,
    expectedText: string,
    signal?: AbortSignal,
  ): Promise<void> {
    cancelled(signal)
    pageId(slideId)
    pageId(shapeId)
    if (
      typeof text !== 'string' ||
      text.length > MAX_POWERPOINT_TEXT ||
      typeof expectedText !== 'string' ||
      expectedText.length > MAX_POWERPOINT_TEXT
    )
      throw new Error('invalid_tool_input')
    await this.run('1.8', async (context) => {
      const { cell } = await pageTableCell(context, slideId, shapeId, rowIndex, columnIndex, signal)
      await writeTextRange(context, cell, text, signal, expectedText)
    })
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
  ): Promise<{ slideId: string; base64: string; mime: 'image/png' }> {
    cancelled(signal)
    let targetSlideId: string | undefined
    try {
      return await this.runScreenshot(
        '1.8',
        async (context) => {
          const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
          const slide = await getSlide(context, slides, slideIndex, signal)
          const slideId = string(slide.id)
          if (targetSlideId && targetSlideId !== slideId)
            throw new Error('office_concurrent_change')
          targetSlideId = slideId
          if (typeof slide.getImageAsBase64 !== 'function')
            throw Object.assign(new Error('office_api_unsupported'), {
              code: 'office_screenshot_unavailable',
            })
          const image = (slide.getImageAsBase64 as (options: { width: number }) => RuntimeRecord)({
            width: 960,
          })
          await sync(context, signal)
          if (slide.id !== slideId) throw new Error('office_concurrent_change')
          if (
            typeof image.value !== 'string' ||
            typeof slide.id !== 'string' ||
            !slide.id ||
            slide.id.length > 256
          )
            throw new Error('office_read_failed')
          return { slideId, base64: image.value, mime: 'image/png' }
        },
        signal,
      )
    } catch (error) {
      const code =
        error && typeof error === 'object' ? (error as { code?: unknown }).code : undefined
      if (
        !targetSlideId ||
        !['office_screenshot_unavailable', 'ActivityLimitReached', 'Timeout'].includes(String(code))
      )
        throw error
      const failure = error instanceof Error ? error : new Error(String(error))
      throw Object.assign(new Error(failure.message, { cause: failure }), {
        code,
        targetSlideId,
      })
    }
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

  async readSlideTable(
    slideIndex: number,
    shapeId: string,
    signal?: AbortSignal,
  ): Promise<string[][]> {
    cancelled(signal)
    return this.run('1.8', async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      const slide = await getSlide(context, slides, slideIndex, signal)
      const shapes = slide.shapes as RuntimeRecord
      if (typeof shapes.getItem !== 'function') throw new Error('office_api_unsupported')
      const shape = (shapes.getItem as (id: string) => RuntimeRecord)(shapeId)
      if (typeof shape.getTable !== 'function') throw new Error('office_api_unsupported')
      const table = (shape.getTable as () => RuntimeRecord)()
      if (typeof table.load !== 'function') throw new Error('office_api_unsupported')
      ;(table.load as (properties: string) => void)('values,rowCount,columnCount')
      await sync(context, signal)
      const columnCount = table.columnCount
      if (
        !Array.isArray(table.values) ||
        table.values.length < 1 ||
        table.values.length !== table.rowCount ||
        table.values.length > 20 ||
        typeof columnCount !== 'number' ||
        !Number.isSafeInteger(columnCount) ||
        columnCount < 1 ||
        columnCount > 12 ||
        table.values.some(
          (row: unknown) =>
            !Array.isArray(row) ||
            row.length !== columnCount ||
            row.some((cell: unknown) => typeof cell !== 'string' || cell.length > 256),
        )
      )
        throw new Error('office_read_failed')
      return table.values as string[][]
    })
  }

  async inspectPresentationPage(
    slideId: string,
    signal?: AbortSignal,
    fallbackBase64?: string,
  ): Promise<PowerPointPageInspection> {
    cancelled(signal)
    if (typeof slideId !== 'string' || !slideId.length || slideId.length > 256)
      throw new Error('invalid_tool_input')
    return this.runScreenshot(
      '1.10',
      async (context) => {
        const presentation = context.presentation as RuntimeRecord
        const slides = presentation.slides as RuntimeRecord
        const pageSetup = presentation.pageSetup as RuntimeRecord | undefined
        if (typeof slides?.getItem !== 'function' || typeof pageSetup?.load !== 'function')
          throw new Error('office_api_unsupported')
        // Resolve the durable host ID directly; global verification only visits the first 20 pages.
        const slide = (slides.getItem as (id: string) => RuntimeRecord)(slideId)
        const collection = slide?.shapes as RuntimeRecord | undefined
        if (typeof slide?.load !== 'function' || typeof collection?.load !== 'function')
          throw new Error('office_api_unsupported')
        if (!fallbackBase64 && typeof slide.getImageAsBase64 !== 'function')
          throw Object.assign(new Error('office_api_unsupported'), {
            code: 'office_screenshot_unavailable',
          })
        ;(slide.load as (properties: string) => void)('id')
        ;(pageSetup.load as (properties: string[]) => void)(['slideWidth', 'slideHeight'])
        loadShapes(collection)
        const image = fallbackBase64
          ? undefined
          : (slide.getImageAsBase64 as (options: { width: number }) => RuntimeRecord)({
              width: 960,
            })
        await sync(context, signal)
        if (slide.id !== slideId || !Array.isArray(collection.items))
          throw new Error('office_read_failed')
        const slideWidth = pageSetup.slideWidth,
          slideHeight = pageSetup.slideHeight
        if (
          typeof slideWidth !== 'number' ||
          !Number.isFinite(slideWidth) ||
          slideWidth <= 0 ||
          typeof slideHeight !== 'number' ||
          !Number.isFinite(slideHeight) ||
          slideHeight <= 0
        )
          throw new Error('office_read_failed')
        const raw = (collection.items as RuntimeRecord[]).slice(0, MAX_POWERPOINT_VERIFY_SHAPES)
        for (const shape of raw) {
          if (
            !shape ||
            typeof shape.id !== 'string' ||
            !shape.id.length ||
            shape.id.length > 256 ||
            ['left', 'top', 'width', 'height'].some(
              (key) => typeof shape[key] !== 'number' || !Number.isFinite(shape[key]),
            ) ||
            (shape.width as number) < 0 ||
            (shape.height as number) < 0 ||
            !Number.isFinite((shape.left as number) + (shape.width as number)) ||
            !Number.isFinite((shape.top as number) + (shape.height as number))
          )
            throw new Error('office_read_failed')
        }
        const shapes = raw.map(shapeInfo)
        if (new Set(shapes.map((shape) => shape.id)).size !== shapes.length)
          throw new Error('office_read_failed')
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
        let overlapsTruncated = false
        for (let i = 0; i < shapes.length; i++)
          for (let j = i + 1; j < shapes.length; j++) {
            const a = shapes[i]!,
              b = shapes[j]!
            if (
              explicitCanvasBackground(a, slideWidth, slideHeight) ||
              explicitCanvasBackground(b, slideWidth, slideHeight)
            )
              continue
            const overlapX = Math.min(a.left + a.width, b.left + b.width) - Math.max(a.left, b.left)
            const overlapY = Math.min(a.top + a.height, b.top + b.height) - Math.max(a.top, b.top)
            if (overlapX > 0 && overlapY > 0) {
              if (overlaps.length === MAX_POWERPOINT_VERIFY_OVERLAPS) {
                overlapsTruncated = true
                break
              }
              overlaps.push({ shapeAId: a.id, shapeBId: b.id, overlapX, overlapY })
            }
          }
        const screenshot = (value: unknown): string => {
          try {
            return validatePowerPointPageScreenshot(value)
          } catch {
            throw Object.assign(new Error('office_read_failed'), {
              code: 'office_screenshot_unavailable',
            })
          }
        }
        let base64 = screenshot(fallbackBase64 ?? image?.value)
        // Keep a single PNG well below the Office transport's 256 KiB total request limit.
        // Use the same deterministic widths when recapturing for a review.
        const fitsModelBudget = () => atob(base64).length <= 64 * 1024
        for (const width of fallbackBase64 ? [] : [640, 480, 320, 240]) {
          if (fitsModelBudget()) break
          const smaller = (slide.getImageAsBase64 as (options: { width: number }) => RuntimeRecord)(
            {
              width,
            },
          )
          await sync(context, signal)
          base64 = screenshot(smaller.value)
        }
        if (!fitsModelBudget()) throw new Error('office_image_too_large')
        cancelled(signal)
        return {
          slideId,
          slideWidth,
          slideHeight,
          shapes,
          shapesTruncated: collection.items.length > MAX_POWERPOINT_VERIFY_SHAPES,
          overflows,
          overlaps,
          overlapsTruncated,
          screenshot: {
            mime: 'image/png',
            base64,
            ...(fallbackBase64 ? { renderer: 'libreoffice' as const } : {}),
          },
        }
      },
      signal,
    )
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
            if (
              explicitCanvasBackground(a, slideWidth, slideHeight) ||
              explicitCanvasBackground(b, slideWidth, slideHeight)
            )
              continue
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
    includeShapes = false,
  ): Promise<{
    slideId: string
    fingerprint: string
    shapes?: Array<
      PowerPointShape & {
        text: string
        tableValues?: string[][]
        rotation?: number
        altTextTitle?: string
        altTextDescription?: string
      }
    >
  }> {
    cancelled(signal)
    return this.run('1.4', async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      const slide = await getSlide(context, slides, slideIndex, signal)
      const shapes = slide.shapes as RuntimeRecord
      if (!shapes || typeof shapes.load !== 'function') throw new Error('office_api_unsupported')
      loadShapes(shapes, MAX_POWERPOINT_SHAPES, includeShapes)
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
      const tables = new Map<RuntimeRecord, RuntimeRecord>()
      for (const shape of items) {
        const textFrame = shape.textFrame as RuntimeRecord | undefined
        const textRange = textFrame?.textRange as RuntimeRecord | undefined
        if (textFrame?.hasText === true && textRange && typeof textRange.load === 'function')
          (textRange.load as (properties: string) => void)('text')
        if (includeShapes && shape.type === 'Table') {
          if (typeof shape.getTable !== 'function') throw new Error('office_api_unsupported')
          const table = (shape.getTable as () => RuntimeRecord)()
          if (!table || typeof table.load !== 'function') throw new Error('office_api_unsupported')
          ;(table.load as (properties: string) => void)('values,rowCount,columnCount')
          tables.set(shape, table)
        }
      }
      await sync(context, signal)
      const slideId = string(slide.id)
      const semanticShapes = items
        .map((shape) => {
          if (
            includeShapes &&
            (typeof shape.id !== 'string' ||
              !shape.id ||
              shape.id.length > 256 ||
              typeof shape.name !== 'string' ||
              shape.name.length > 256 ||
              typeof shape.type !== 'string' ||
              shape.type.length > 64 ||
              ![shape.left, shape.top, shape.width, shape.height].every(
                (value) => typeof value === 'number' && Number.isFinite(value),
              ))
          )
            throw new Error('office_read_failed')
          const textFrame = shape.textFrame as RuntimeRecord | undefined
          const textRange = textFrame?.textRange as RuntimeRecord | undefined
          const table = tables.get(shape)
          let tableValues: string[][] | undefined
          if (table) {
            const values = table.values
            if (
              !Array.isArray(values) ||
              values.length < 1 ||
              values.length > 20 ||
              values.length !== table.rowCount ||
              !Number.isSafeInteger(table.columnCount) ||
              (table.columnCount as number) < 1 ||
              (table.columnCount as number) > 12 ||
              values.some(
                (row: unknown) =>
                  !Array.isArray(row) ||
                  row.length !== table.columnCount ||
                  row.some((cell: unknown) => typeof cell !== 'string' || cell.length > 256),
              )
            )
              throw new Error('office_read_failed')
            tableValues = (values as string[][]).map((row) => [...row])
          }
          const picture = includeShapes && ['Image', 'Picture'].includes(String(shape.type))
          if (
            picture &&
            (typeof shape.rotation !== 'number' ||
              !Number.isFinite(shape.rotation) ||
              Math.abs(shape.rotation) > 360 ||
              typeof shape.altTextTitle !== 'string' ||
              shape.altTextTitle.length > 12000 ||
              typeof shape.altTextDescription !== 'string' ||
              shape.altTextDescription.length > 12000)
          )
            throw new Error('office_read_failed')
          return {
            ...shapeInfo(shape),
            text:
              textFrame?.hasText === true
                ? includeShapes
                  ? boundedPageText(textRange?.text)
                  : string(textRange?.text, MAX_POWERPOINT_TEXT)
                : '',
            ...(tableValues ? { tableValues } : {}),
            ...(picture
              ? {
                  rotation: shape.rotation as number,
                  altTextTitle: shape.altTextTitle as string,
                  altTextDescription: shape.altTextDescription as string,
                }
              : {}),
          }
        })
        .sort((first, second) => first.id.localeCompare(second.id))
      return {
        slideId,
        fingerprint: `${slideId}:${hash(JSON.stringify(semanticShapes))}`,
        ...(includeShapes ? { shapes: semanticShapes } : {}),
      }
    })
  }

  async readSlideOrder(signal?: AbortSignal): Promise<string[]> {
    cancelled(signal)
    return this.run('1.8', async (context) =>
      readCompleteSlideOrder(
        context,
        (context.presentation as RuntimeRecord).slides as RuntimeRecord,
        signal,
      ),
    )
  }

  async exportPresentationPagePackage(
    slideId: string,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; slideIds: string[]; base64: string }> {
    cancelled(signal)
    pageId(slideId)
    return this.run('1.8', async (context) => {
      const slides = (context.presentation as RuntimeRecord).slides as RuntimeRecord
      if (typeof slides?.load !== 'function') throw new Error('office_api_unsupported')
      const readOrder = async (): Promise<string[]> => {
        ;(slides.load as (properties: unknown) => void)({ $top: 514, id: true })
        await sync(context, signal)
        if (!Array.isArray(slides.items) || slides.items.length < 1 || slides.items.length > 513)
          throw new Error('office_read_failed')
        const ids = (slides.items as RuntimeRecord[]).map((item) => item?.id)
        if (
          ids.some(
            (id) =>
              typeof id !== 'string' ||
              !id ||
              id.length > 256 ||
              Array.from(id).some(
                (char) =>
                  char.charCodeAt(0) < 32 ||
                  (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
              ),
          ) ||
          new Set(ids).size !== ids.length ||
          !ids.includes(slideId)
        )
          throw new Error('office_read_failed')
        return ids as string[]
      }
      const slideIds = await readOrder()
      const slide = await getPageById(context, slideId, signal)
      if (typeof slide.exportAsBase64 !== 'function') throw new Error('office_api_unsupported')
      const exported = (slide.exportAsBase64 as () => RuntimeRecord)()
      await sync(context, signal)
      if (
        typeof exported.value !== 'string' ||
        !exported.value ||
        exported.value.length > Math.ceil((8 * 1024 * 1024) / 3) * 4
      )
        throw new Error('office_read_failed')
      if (JSON.stringify(await readOrder()) !== JSON.stringify(slideIds))
        throw new Error('office_concurrent_change')
      return { slideId, slideIds, base64: exported.value }
    })
  }

  async inspectSlidePictureFingerprints(
    slideId: string,
    shapeIds: string[],
    signal?: AbortSignal,
  ): Promise<{
    slideId: string
    slideIds: string[]
    fingerprints: Record<string, string>
    mediaDigests: Record<string, string>
  }> {
    cancelled(signal)
    const exported = await this.exportPresentationPagePackage(slideId, signal)
    const inspected = await inspectPowerPointPictureMediaBatch(exported.base64, shapeIds, signal)
    cancelled(signal)
    if (
      inspected.unsupported.length ||
      Object.keys(inspected.pictureFingerprints).length !== shapeIds.length ||
      Object.keys(inspected.mediaDigests).length !== shapeIds.length ||
      shapeIds.some(
        (id) =>
          !Object.hasOwn(inspected.pictureFingerprints, id) ||
          !Object.hasOwn(inspected.mediaDigests, id),
      )
    )
      throw new Error('office_api_unsupported')
    return {
      slideId: exported.slideId,
      slideIds: exported.slideIds,
      fingerprints: inspected.pictureFingerprints,
      mediaDigests: inspected.mediaDigests,
    }
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
    preimage?: { slideId: string; packageDigest: string; slideIds?: string[] },
  ): Promise<{ slideId: string }> {
    const ownedPreimage = preimage && structuredClone(preimage)
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
      const originalBase64 = original.value
      const originalSlideId = string(slide.id)
      const originalExpected = await capturePowerPointPackage(originalBase64, signal)
      const replacementExactProof = await capturePowerPointPackage(base64, signal)
      if (
        ownedPreimage &&
        (originalSlideId !== ownedPreimage.slideId ||
          (await presentationPackageDigest(originalBase64, signal)) !== ownedPreimage.packageDigest)
      )
        throw new Error('proposal_stale')

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
      if (ownedPreimage?.slideIds) {
        const order = await readCompleteSlideOrder(context, slides, signal)
        if (
          JSON.stringify(order) !== JSON.stringify(ownedPreimage.slideIds) ||
          order[slideIndex] !== ownedPreimage.slideId
        )
          throw new Error('proposal_stale')
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
          )(originalBase64, insertOptions)
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
      if (
        operation.op === 'add_geometric_shape' &&
        (!['rect', 'ellipse', 'roundRect'].includes(operation.shape) ||
          !operation.name ||
          operation.name.length > 256 ||
          !/^[0-9A-Fa-f]{6}$/.test(operation.fill) ||
          !/^[0-9A-Fa-f]{6}$/.test(operation.lineColor) ||
          [operation.left, operation.top, operation.width, operation.height].some(
            (value) => !Number.isFinite(value),
          ) ||
          operation.width <= 0 ||
          operation.height <= 0)
      )
        throw new Error('invalid_tool_input')
      if (
        operation.op === 'add_native_table' &&
        (!operation.name ||
          operation.name.length > 256 ||
          !Array.isArray(operation.rows) ||
          operation.rows.length < 1 ||
          operation.rows.length > 20 ||
          !operation.rows[0]?.length ||
          operation.rows[0].length > 12 ||
          operation.rows.length * operation.rows[0].length > 128 ||
          operation.rows.some(
            (row) =>
              !Array.isArray(row) ||
              row.length !== operation.rows[0].length ||
              row.some((cell) => typeof cell !== 'string' || cell.length > 256),
          ) ||
          JSON.stringify(operation.rows).length > 12_000 ||
          [
            operation.left,
            operation.top,
            operation.width,
            operation.height,
            operation.fontSize,
          ].some((value) => !Number.isFinite(value)) ||
          operation.width <= 0 ||
          operation.height <= 0 ||
          operation.fontSize < 6 ||
          operation.fontSize > 48 ||
          !operation.fontFace ||
          operation.fontFace.length > 128 ||
          !/^[0-9A-Fa-f]{6}$/.test(operation.color) ||
          (operation.borderColor !== undefined &&
            !/^[0-9A-Fa-f]{6}$/.test(operation.borderColor)) ||
          (operation.cellMargin !== undefined &&
            (!Number.isFinite(operation.cellMargin) ||
              operation.cellMargin < 0 ||
              operation.cellMargin > 36)))
      )
        throw new Error('invalid_tool_input')
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
            if (operation.margin !== undefined || operation.verticalAlignment) {
              const frame = created.textFrame as RuntimeRecord
              if (operation.margin !== undefined) {
                frame.leftMargin = operation.margin
                frame.rightMargin = operation.margin
                frame.topMargin = operation.margin
                frame.bottomMargin = operation.margin
              }
              if (operation.verticalAlignment)
                frame.verticalAlignment = { top: 'Top', middle: 'Middle', bottom: 'Bottom' }[
                  operation.verticalAlignment
                ]
            }
            if (
              operation.fontFace ||
              operation.fontSize ||
              operation.color ||
              operation.bold !== undefined
            ) {
              const font = ((created.textFrame as RuntimeRecord).textRange as RuntimeRecord)
                .font as RuntimeRecord
              if (operation.fontFace) font.name = operation.fontFace
              if (operation.fontSize) font.size = operation.fontSize
              if (operation.color) font.color = `#${operation.color}`
              if (operation.bold !== undefined) font.bold = operation.bold
            }
            if (operation.align) {
              const paragraph = ((created.textFrame as RuntimeRecord).textRange as RuntimeRecord)
                .paragraphFormat as RuntimeRecord
              paragraph.horizontalAlignment = { left: 'Left', center: 'Center', right: 'Right' }[
                operation.align
              ]
            }
            if (typeof created.load !== 'function') throw new Error('office_api_unsupported')
            ;(created.load as (properties: string) => void)('id')
            createdShapes.push(created)
          })
          hasUnrecoverableMutation = true
        } else if (operation.op === 'add_geometric_shape') {
          const shapes = slide.shapes as RuntimeRecord
          if (typeof shapes.addGeometricShape !== 'function')
            throw new Error('office_api_unsupported')
          queued.push(() => {
            const shapeType = {
              rect: 'Rectangle',
              ellipse: 'Ellipse',
              roundRect: 'RoundRectangle',
            }[operation.shape]
            const created = (
              shapes.addGeometricShape as (
                kind: string,
                options: Record<string, number>,
              ) => RuntimeRecord
            )(shapeType, {
              left: operation.left,
              top: operation.top,
              width: operation.width,
              height: operation.height,
            })
            created.name = operation.name
            ;((created.fill as RuntimeRecord).setSolidColor as (color: string) => void)(
              `#${operation.fill}`,
            )
            ;(created.lineFormat as RuntimeRecord).color = `#${operation.lineColor}`
            if (typeof created.load !== 'function') throw new Error('office_api_unsupported')
            ;(created.load as (properties: string) => void)('id')
            createdShapes.push(created)
          })
          hasUnrecoverableMutation = true
        } else if (operation.op === 'add_native_table') {
          const shapes = slide.shapes as RuntimeRecord
          if (typeof shapes.addTable !== 'function') throw new Error('office_api_unsupported')
          queued.push(() => {
            const created = (
              shapes.addTable as (
                rows: number,
                columns: number,
                options: Record<string, unknown>,
              ) => RuntimeRecord
            )(operation.rows.length, operation.rows[0]!.length, {
              left: operation.left,
              top: operation.top,
              width: operation.width,
              height: operation.height,
              values: operation.rows,
              uniformCellProperties: {
                font: {
                  name: operation.fontFace,
                  size: operation.fontSize,
                  color: `#${operation.color}`,
                },
                ...(operation.borderColor
                  ? {
                      borders: Object.fromEntries(
                        ['top', 'right', 'bottom', 'left'].map((side) => [
                          side,
                          { color: `#${operation.borderColor}`, weight: 1 },
                        ]),
                      ),
                    }
                  : {}),
                ...(operation.cellMargin !== undefined
                  ? {
                      margins: {
                        top: operation.cellMargin,
                        right: operation.cellMargin,
                        bottom: operation.cellMargin,
                        left: operation.cellMargin,
                      },
                    }
                  : {}),
              },
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
      await writeTextRange(context, textRange, value, signal)
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
