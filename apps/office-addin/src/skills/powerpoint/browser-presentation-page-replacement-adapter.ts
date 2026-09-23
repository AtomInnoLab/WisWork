import type { PresentationPageReplacement } from './presentation-page-replacement-record.js'
import { presentationPackageDigest } from './powerpoint-package.js'

type AssertCurrent = () => void | Promise<void>
export interface PresentationPageReplacementInspection {
  status: 'baseline' | 'staged' | 'conflict'
  slideIds: string[]
}
export interface PresentationPageReplacementAdapter {
  inspect(
    record: PresentationPageReplacement,
    signal?: AbortSignal,
  ): Promise<PresentationPageReplacementInspection>
  stage(
    record: PresentationPageReplacement,
    base64: string,
    onInserted: (newId: string) => Promise<void>,
    assertCurrent: AssertCurrent,
    signal?: AbortSignal,
  ): Promise<void>
  discard(
    record: PresentationPageReplacement,
    assertCurrent: AssertCurrent,
    signal?: AbortSignal,
  ): Promise<void>
}
function check(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('cancelled')
}
function runtime(): typeof PowerPoint {
  const root = globalThis as typeof globalThis & {
    Office?: typeof Office
    PowerPoint?: typeof PowerPoint
  }
  if (
    String(root.Office?.context?.host) !== 'PowerPoint' ||
    !root.Office?.context?.requirements?.isSetSupported?.('PowerPointApi', '1.8') ||
    typeof root.PowerPoint?.run !== 'function'
  )
    throw new Error('office_api_unsupported')
  return root.PowerPoint
}
async function sync(context: PowerPoint.RequestContext, signal?: AbortSignal) {
  check(signal)
  await context.sync()
  check(signal)
}
const same = (a: string[], b: string[]) => JSON.stringify(a) === JSON.stringify(b)
function stagedIds(record: PresentationPageReplacement): string[] {
  const ids = [...record.beforeSlideIds]
  if (record.newSlideId) ids.splice(ids.indexOf(record.oldSlideId) + 1, 0, record.newSlideId)
  return ids
}
async function order(context: PowerPoint.RequestContext, signal?: AbortSignal): Promise<string[]> {
  const slides = context.presentation.slides
  if (typeof slides?.load !== 'function') throw new Error('office_api_unsupported')
  slides.load({ $top: 514, id: true })
  await sync(context, signal)
  if (!Array.isArray(slides.items) || !slides.items.length || slides.items.length > 513)
    throw new Error('office_read_failed')
  const ids = slides.items.map((item: PowerPoint.Slide) => item?.id)
  if (
    ids.some(
      (id: unknown) =>
        typeof id !== 'string' ||
        !id ||
        id.length > 256 ||
        Array.from(id).some(
          (char) =>
            char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
        ),
    ) ||
    new Set(ids).size !== ids.length
  )
    throw new Error('office_read_failed')
  return ids
}
async function page(
  context: PowerPoint.RequestContext,
  id: string,
  signal?: AbortSignal,
): Promise<PowerPoint.Slide> {
  if (typeof context.presentation.slides.getItem !== 'function')
    throw new Error('office_api_unsupported')
  const slide = context.presentation.slides.getItem(id)
  if (typeof slide?.load !== 'function') throw new Error('office_api_unsupported')
  slide.load('id')
  await sync(context, signal)
  if (slide.id !== id) throw new Error('office_concurrent_change')
  return slide
}
async function digest(
  context: PowerPoint.RequestContext,
  id: string,
  signal?: AbortSignal,
): Promise<string> {
  const slide = await page(context, id, signal)
  if (typeof slide.exportAsBase64 !== 'function') throw new Error('office_api_unsupported')
  const exported = slide.exportAsBase64()
  await sync(context, signal)
  if (typeof exported?.value !== 'string') throw new Error('office_read_failed')
  return presentationPackageDigest(exported.value, signal)
}
async function inspect(
  context: PowerPoint.RequestContext,
  record: PresentationPageReplacement,
  signal?: AbortSignal,
): Promise<PresentationPageReplacementInspection> {
  const slideIds = await order(context, signal)
  let status: PresentationPageReplacementInspection['status'] = 'conflict'
  if (same(slideIds, record.beforeSlideIds)) status = 'baseline'
  else if (record.newSlideId && same(slideIds, stagedIds(record))) status = 'staged'
  if (status !== 'conflict') {
    if ((await digest(context, record.oldSlideId, signal)) !== record.originalPackageDigest)
      status = 'conflict'
    if (
      status === 'staged' &&
      (await digest(context, record.newSlideId!, signal)) !== record.replacementPackageDigest
    )
      status = 'conflict'
    if (!same(await order(context, signal), slideIds)) status = 'conflict'
  }
  return { status, slideIds }
}

/** A durable pending journal must exist before stage; this adapter never deletes the original. */
export class BrowserPresentationPageReplacementAdapter implements PresentationPageReplacementAdapter {
  async inspect(
    record: PresentationPageReplacement,
    signal?: AbortSignal,
  ): Promise<PresentationPageReplacementInspection> {
    check(signal)
    const saved = structuredClone(record)
    return runtime().run((context: PowerPoint.RequestContext) => inspect(context, saved, signal))
  }
  async stage(
    record: PresentationPageReplacement,
    base64: string,
    onInserted: (newId: string) => Promise<void>,
    assertCurrent: AssertCurrent,
    signal?: AbortSignal,
  ): Promise<void> {
    const saved = structuredClone(record)
    check(signal)
    if (saved.state !== 'pending' || saved.newSlideId) throw new Error('office_concurrent_change')
    if ((await presentationPackageDigest(base64, signal)) !== saved.replacementPackageDigest)
      throw new Error('office_concurrent_change')
    await assertCurrent()
    check(signal)
    await runtime().run(async (context: PowerPoint.RequestContext) => {
      if (typeof context.presentation.insertSlidesFromBase64 !== 'function')
        throw new Error('office_api_unsupported')
      if ((await inspect(context, saved, signal)).status !== 'baseline')
        throw new Error('office_concurrent_change')
      await assertCurrent()
      check(signal)
      // Office 1.2: targetSlideId inserts AFTER that slide; explicit source selects one page.
      context.presentation.insertSlidesFromBase64(base64, {
        targetSlideId: saved.oldSlideId,
        sourceSlideIds: [saved.sourceSlideId],
        formatting: 'KeepSourceFormatting',
      })
      await sync(context, signal)
      const after = await order(context, signal)
      const added = after.filter((id) => !saved.beforeSlideIds.includes(id))
      if (added.length !== 1 || !same(after, stagedIds({ ...saved, newSlideId: added[0]! })))
        throw new Error('office_concurrent_change')
      // Save the exact new ID before content verification; any failure remains recoverable.
      await assertCurrent()
      check(signal)
      await onInserted(added[0]!)
      await assertCurrent()
      check(signal)
      if ((await inspect(context, { ...saved, newSlideId: added[0]! }, signal)).status !== 'staged')
        throw new Error('office_concurrent_change')
      await assertCurrent()
      check(signal)
    })
  }
  async discard(
    record: PresentationPageReplacement,
    assertCurrent: AssertCurrent,
    signal?: AbortSignal,
  ): Promise<void> {
    const saved = structuredClone(record)
    check(signal)
    if (saved.state !== 'discard_pending' || !saved.newSlideId)
      throw new Error('office_concurrent_change')
    await assertCurrent()
    check(signal)
    await runtime().run(async (context: PowerPoint.RequestContext) => {
      const current = await inspect(context, saved, signal)
      if (current.status === 'conflict') throw new Error('office_concurrent_change')
      if (current.status === 'baseline') {
        await assertCurrent()
        check(signal)
        return
      }
      const slide = await page(context, saved.newSlideId!, signal)
      if (typeof slide.delete !== 'function') throw new Error('office_api_unsupported')
      await assertCurrent()
      check(signal)
      slide.delete()
      await sync(context, signal)
      if ((await inspect(context, saved, signal)).status !== 'baseline')
        throw new Error('office_concurrent_change')
      await assertCurrent()
      check(signal)
    })
  }
}
