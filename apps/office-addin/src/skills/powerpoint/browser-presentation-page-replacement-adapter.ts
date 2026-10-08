import type { PresentationPageReplacement } from './presentation-page-replacement-record.js'
import { presentationPackageDigest } from './powerpoint-package.js'

type AssertCurrent = () => void | Promise<void>
export interface PresentationPageReplacementInspection {
  status: 'baseline' | 'staged' | 'applied' | 'restore_staged' | 'undone' | 'conflict'
  slideIds: string[]
}
export interface PresentationPageReplacementAdapter {
  captureUnchangedPageDigests(
    slideIds: string[],
    beforeSlideIds: string[],
    signal?: AbortSignal,
  ): Promise<{ slideId: string; digest: string }[]>
  reconcilePending(
    record: PresentationPageReplacement,
    signal?: AbortSignal,
  ): Promise<{ status: 'baseline' | 'inserted' | 'conflict'; newSlideId?: string }>
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
  commit(
    record: PresentationPageReplacement,
    assertCurrent: AssertCurrent,
    signal?: AbortSignal,
  ): Promise<void>
  undo(
    record: PresentationPageReplacement,
    backupBase64: string,
    onRestored: (id: string) => Promise<void>,
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
function appliedIds(record: PresentationPageReplacement): string[] {
  return record.beforeSlideIds.map((id) => (id === record.oldSlideId ? record.newSlideId! : id))
}
function restoredIds(record: PresentationPageReplacement, staged = false): string[] {
  const ids = appliedIds(record)
  ids.splice(
    ids.indexOf(record.newSlideId!) + (staged ? 1 : 0),
    staged ? 0 : 1,
    record.restoredSlideId!,
  )
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
  else if (record.newSlideId && same(slideIds, appliedIds(record))) status = 'applied'
  else if (record.newSlideId && record.restoredSlideId && same(slideIds, restoredIds(record, true)))
    status = 'restore_staged'
  else if (record.restoredSlideId && same(slideIds, restoredIds(record))) status = 'undone'
  if (status !== 'conflict') {
    const checks: [string, string][] = []
    if (status === 'baseline' || status === 'staged')
      checks.push([record.oldSlideId, record.originalPackageDigest])
    if (status === 'staged' || status === 'applied' || status === 'restore_staged')
      checks.push([record.newSlideId!, record.replacementPackageDigest])
    if (status === 'restore_staged' || status === 'undone')
      checks.push([record.restoredSlideId!, record.originalPackageDigest])
    for (const [id, expected] of checks) {
      if ((await digest(context, id, signal)) !== expected) status = 'conflict'
    }
    for (const item of record.untouchedSlideDigests ?? []) {
      if ((await digest(context, item.slideId, signal)) !== item.digest) status = 'conflict'
    }
    if (!same(await order(context, signal), slideIds)) status = 'conflict'
  }
  return { status, slideIds }
}

/** Every write requires its durable intent and exact page identity/content proof. */
export class BrowserPresentationPageReplacementAdapter implements PresentationPageReplacementAdapter {
  async captureUnchangedPageDigests(
    slideIds: string[],
    beforeSlideIds: string[],
    signal?: AbortSignal,
  ): Promise<{ slideId: string; digest: string }[]> {
    check(signal)
    if (
      slideIds.length > 31 ||
      new Set(slideIds).size !== slideIds.length ||
      !same(
        slideIds,
        beforeSlideIds.filter((id) => slideIds.includes(id)),
      )
    )
      throw new Error('office_concurrent_change')
    return runtime().run(async (context: PowerPoint.RequestContext) => {
      if (!same(await order(context, signal), beforeSlideIds))
        throw new Error('office_concurrent_change')
      const captured = []
      for (const slideId of slideIds)
        captured.push({ slideId, digest: await digest(context, slideId, signal) })
      if (!same(await order(context, signal), beforeSlideIds))
        throw new Error('office_concurrent_change')
      return captured
    })
  }
  async reconcilePending(
    record: PresentationPageReplacement,
    signal?: AbortSignal,
  ): Promise<{ status: 'baseline' | 'inserted' | 'conflict'; newSlideId?: string }> {
    check(signal)
    const saved = structuredClone(record)
    if (saved.state !== 'pending' || saved.newSlideId) throw new Error('office_concurrent_change')
    return runtime().run(async (context: PowerPoint.RequestContext) => {
      const ids = await order(context, signal)
      if (same(ids, saved.beforeSlideIds)) {
        return (await inspect(context, saved, signal)).status === 'baseline'
          ? { status: 'baseline' as const }
          : { status: 'conflict' as const }
      }
      const added = ids.filter((id) => !saved.beforeSlideIds.includes(id))
      if (
        added.length !== 1 ||
        !same(ids, stagedIds({ ...saved, newSlideId: added[0]! })) ||
        (await inspect(context, { ...saved, newSlideId: added[0]! }, signal)).status !== 'staged'
      )
        return { status: 'conflict' }
      return { status: 'inserted', newSlideId: added[0]! }
    })
  }
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
      if (current.status !== 'baseline' && current.status !== 'staged')
        throw new Error('office_concurrent_change')
      if (current.status === 'baseline') {
        await assertCurrent()
        check(signal)
        return
      }
      const slide = await page(context, saved.newSlideId!, signal)
      if (typeof slide.delete !== 'function') throw new Error('office_api_unsupported')
      // Obtaining the deletion proxy syncs with Office: recheck content after that await.
      if ((await inspect(context, saved, signal)).status !== 'staged')
        throw new Error('office_concurrent_change')
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
  async commit(
    record: PresentationPageReplacement,
    assertCurrent: AssertCurrent,
    signal?: AbortSignal,
  ): Promise<void> {
    const saved = structuredClone(record)
    check(signal)
    if (saved.state !== 'commit_pending' || !saved.newSlideId)
      throw new Error('office_concurrent_change')
    await assertCurrent()
    check(signal)
    await runtime().run(async (context: PowerPoint.RequestContext) => {
      const current = await inspect(context, saved, signal)
      if (current.status === 'applied') {
        await assertCurrent()
        check(signal)
        return
      }
      if (current.status !== 'staged') throw new Error('office_concurrent_change')
      const slide = await page(context, saved.oldSlideId, signal)
      if (typeof slide.delete !== 'function') throw new Error('office_api_unsupported')
      if ((await inspect(context, saved, signal)).status !== 'staged')
        throw new Error('office_concurrent_change')
      await assertCurrent()
      check(signal)
      slide.delete()
      await sync(context, signal)
      if ((await inspect(context, saved, signal)).status !== 'applied')
        throw new Error('office_concurrent_change')
      await assertCurrent()
      check(signal)
    })
  }
  async undo(
    record: PresentationPageReplacement,
    backupBase64: string,
    onRestored: (id: string) => Promise<void>,
    assertCurrent: AssertCurrent,
    signal?: AbortSignal,
  ): Promise<void> {
    let saved = structuredClone(record)
    check(signal)
    if (
      !saved.newSlideId ||
      !['undo_pending', 'restore_inserted'].includes(saved.state) ||
      (saved.state === 'undo_pending' ? !!saved.restoredSlideId : !saved.restoredSlideId)
    )
      throw new Error('office_concurrent_change')
    if ((await presentationPackageDigest(backupBase64, signal)) !== saved.originalPackageDigest)
      throw new Error('office_concurrent_change')
    await assertCurrent()
    check(signal)
    await runtime().run(async (context: PowerPoint.RequestContext) => {
      const current = await inspect(context, saved, signal)
      if (saved.state === 'undo_pending') {
        if (current.status !== 'applied') throw new Error('office_concurrent_change')
        if (typeof context.presentation.insertSlidesFromBase64 !== 'function')
          throw new Error('office_api_unsupported')
        await assertCurrent()
        check(signal)
        context.presentation.insertSlidesFromBase64(backupBase64, {
          targetSlideId: saved.newSlideId,
          formatting: 'KeepSourceFormatting',
        })
        await sync(context, signal)
        const after = await order(context, signal)
        const added = after.filter((id) => !current.slideIds.includes(id))
        if (
          added.length !== 1 ||
          saved.beforeSlideIds.includes(added[0]!) ||
          !same(after, restoredIds({ ...saved, restoredSlideId: added[0]! }, true))
        )
          throw new Error('office_concurrent_change')
        await assertCurrent()
        check(signal)
        await onRestored(added[0]!)
        saved = { ...saved, state: 'restore_inserted', restoredSlideId: added[0]! }
        await assertCurrent()
        check(signal)
      } else if (current.status === 'undone') {
        await assertCurrent()
        check(signal)
        return
      } else if (current.status !== 'restore_staged') throw new Error('office_concurrent_change')
      const slide = await page(context, saved.newSlideId!, signal)
      if (typeof slide.delete !== 'function') throw new Error('office_api_unsupported')
      if ((await inspect(context, saved, signal)).status !== 'restore_staged')
        throw new Error('office_concurrent_change')
      await assertCurrent()
      check(signal)
      slide.delete()
      await sync(context, signal)
      if ((await inspect(context, saved, signal)).status !== 'undone')
        throw new Error('office_concurrent_change')
      await assertCurrent()
      check(signal)
    })
  }
}
