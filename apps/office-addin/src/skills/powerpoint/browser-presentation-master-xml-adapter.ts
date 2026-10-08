import { presentationPackageDigest } from './powerpoint-package.js'
import {
  inspectPowerPointPackageHost,
  readPackageSourceSlideId,
  readPowerPointPackageHostOrder,
  type PackageHostSnapshot,
} from './browser-presentation-package-edit-adapter.js'

export interface MasterXmlHostSnapshot extends PackageHostSnapshot {
  masters: { masterId: string; name: string; layouts: { layoutId: string; name: string }[] }[]
  dependencies: { slideId: string; masterId: string; layoutId: string }[]
}
export interface MasterXmlInsertedIdentity {
  slideId: string
  masterId: string
  layoutId: string
}
export interface MasterXmlStageRequest {
  base64: string
  sourceSlideId: string
  packageSourceSlideId: string
  preimage: MasterXmlHostSnapshot
}
export interface MasterXmlRemoveRequest {
  slideId: string
  preimage: MasterXmlHostSnapshot
}
export interface MasterXmlApplyLayoutRequest {
  slideId: string
  masterId: string
  layoutId: string
  preimage: MasterXmlHostSnapshot
}
const MAX_METADATA_BYTES = 8 * 1024 * 1024
function check(signal?: AbortSignal) {
  if (signal?.aborted) throw Error('cancelled')
}
function hostId(value: unknown): asserts value is string {
  if (
    typeof value !== 'string' ||
    !value.trim() ||
    value.length > 256 ||
    Array.from(value).some(
      (char) => char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
    )
  )
    throw Error('office_read_failed')
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
const metadata = (value: MasterXmlHostSnapshot) => ({
  slideIds: value.slideIds,
  pages: value.pages.map(({ slideId, digest }) => ({ slideId, digest })),
  masters: value.masters,
  dependencies: value.dependencies,
})
type Native = Pick<MasterXmlHostSnapshot, 'slideIds' | 'masters' | 'dependencies'>
function validNative(value: Native): void {
  if (
    !value ||
    !Array.isArray(value.slideIds) ||
    !value.slideIds.length ||
    !Array.isArray(value.masters) ||
    !value.masters.length ||
    !Array.isArray(value.dependencies) ||
    value.dependencies.length !== value.slideIds.length
  )
    throw Error('office_read_failed')
  value.slideIds.forEach(hostId)
  const masterIds = new Set<string>(),
    layoutIds = new Set<string>()
  for (const master of value.masters) {
    hostId(master.masterId)
    if (
      masterIds.has(master.masterId) ||
      typeof master.name !== 'string' ||
      master.name.length > 2048 ||
      !Array.isArray(master.layouts) ||
      !master.layouts.length
    )
      throw Error('office_read_failed')
    masterIds.add(master.masterId)
    for (const layout of master.layouts) {
      hostId(layout.layoutId)
      if (
        layoutIds.has(layout.layoutId) ||
        typeof layout.name !== 'string' ||
        layout.name.length > 2048
      )
        throw Error('office_read_failed')
      layoutIds.add(layout.layoutId)
    }
  }
  if (
    new Set(value.slideIds).size !== value.slideIds.length ||
    value.dependencies.some(
      (page, i) =>
        page.slideId !== value.slideIds[i] ||
        !value.masters.some(
          (master) =>
            master.masterId === page.masterId &&
            master.layouts.some((layout) => layout.layoutId === page.layoutId),
        ),
    ) ||
    new TextEncoder().encode(JSON.stringify(value)).byteLength > MAX_METADATA_BYTES
  )
    throw Error('office_read_failed')
}
function valid(value: MasterXmlHostSnapshot): void {
  validNative({
    slideIds: value.slideIds,
    masters: value.masters,
    dependencies: value.dependencies,
  })
  if (
    !Array.isArray(value.pages) ||
    value.pages.length !== value.slideIds.length ||
    value.pages.some(
      (page, i) => page.slideId !== value.slideIds[i] || !/^[a-f0-9]{64}$/.test(page.digest),
    ) ||
    new TextEncoder().encode(JSON.stringify(metadata(value))).byteLength > MAX_METADATA_BYTES
  )
    throw Error('office_read_failed')
}
function runtime(): typeof PowerPoint {
  if (
    String(globalThis.Office?.context?.host) !== 'PowerPoint' ||
    String(Office.context.platform).toLowerCase() === 'mac' ||
    !Office.context.requirements.isSetSupported('PowerPointApi', '1.8') ||
    typeof globalThis.PowerPoint?.run !== 'function'
  )
    throw Error('office_api_unsupported')
  return PowerPoint
}
async function sync(context: PowerPoint.RequestContext, signal?: AbortSignal) {
  check(signal)
  await context.sync()
  check(signal)
}
async function native(context: PowerPoint.RequestContext, signal?: AbortSignal): Promise<Native> {
  const { slides, slideMasters } = context.presentation
  if (
    typeof slideMasters?.getCount !== 'function' ||
    typeof slideMasters.load !== 'function' ||
    typeof slides.getCount !== 'function'
  )
    throw Error('office_api_unsupported')
  const load = () => {
    slides.load('items/id,items/slideMaster/id,items/layout/id')
    slideMasters.load('items/id,items/name,items/layouts/items/id,items/layouts/items/name')
  }
  load()
  const initialMasterCount = slideMasters.getCount(),
    initialSlideCount = slides.getCount()
  await sync(context, signal)
  if (
    initialMasterCount.value !== slideMasters.items.length ||
    initialSlideCount.value !== slides.items.length
  )
    throw Error('office_read_failed')
  const initial = slideMasters.items.map((master) => ({
    id: master.id,
    layouts: master.layouts.items.map((layout) => layout.id),
  }))
  const counts = slideMasters.items.map((master) => {
    if (typeof master.layouts.getCount !== 'function') throw Error('office_api_unsupported')
    master.layouts.load('items/id,items/name')
    return { masterId: master.id, count: master.layouts.getCount() }
  })
  load()
  const masterCount = slideMasters.getCount(),
    slideCount = slides.getCount()
  await sync(context, signal)
  if (
    masterCount.value !== slideMasters.items.length ||
    slideCount.value !== slides.items.length ||
    !Number.isSafeInteger(masterCount.value) ||
    !Number.isSafeInteger(slideCount.value) ||
    !same(
      initial,
      slideMasters.items.map((master) => ({
        id: master.id,
        layouts: master.layouts.items.map((layout) => layout.id),
      })),
    ) ||
    counts.some(
      (entry) =>
        entry.count.value !==
          slideMasters.items.find((master) => master.id === entry.masterId)?.layouts.items.length ||
        !Number.isSafeInteger(entry.count.value),
    )
  )
    throw Error('office_read_failed')
  const slideIds = slides.items.map((slide) => slide.id)
  const result = {
    slideIds,
    masters: slideMasters.items
      .map((master) => ({
        masterId: master.id,
        name: master.name,
        layouts: master.layouts.items
          .map((layout) => ({ layoutId: layout.id, name: layout.name }))
          .sort((a, b) => a.layoutId.localeCompare(b.layoutId)),
      }))
      .sort((a, b) => a.masterId.localeCompare(b.masterId)),
    dependencies: slides.items.map((slide) => ({
      slideId: slide.id,
      masterId: slide.slideMaster.id,
      layoutId: slide.layout.id,
    })),
  }
  validNative(result)
  return result
}
async function inspect(
  context: PowerPoint.RequestContext,
  exports: readonly string[],
  signal?: AbortSignal,
): Promise<MasterXmlHostSnapshot> {
  const before = structuredClone(await native(context, signal))
  const packages = await inspectPowerPointPackageHost(context, exports, signal)
  const after = await native(context, signal)
  if (!same(before, after) || !same(packages.slideIds, after.slideIds))
    throw Error('proposal_stale')
  const result = { ...packages, masters: after.masters, dependencies: after.dependencies }
  valid(result)
  return result
}
async function preflight(
  context: PowerPoint.RequestContext,
  before: MasterXmlHostSnapshot,
  signal?: AbortSignal,
) {
  const actual = await inspect(context, [], signal)
  if (!same(metadata(actual), metadata(before))) throw Error('proposal_stale')
}
/** Native membership is observed by ID. Package/layout correspondence must be separately proved by the engine. */
export class BrowserPresentationMasterXmlAdapter {
  /** A single-page observation; callers close the complete document proof separately. */
  async readPage(
    slideId: string,
    signal?: AbortSignal,
  ): Promise<{
    slideId: string
    masterId: string
    layoutId: string
    digest: string
    base64: string
  }> {
    const ownedId = slideId
    hostId(ownedId)
    check(signal)
    return runtime().run(async (context) => {
      const slide = context.presentation.slides.getItem(ownedId)
      const identity = () => {
        const result = {
          slideId: slide.id,
          masterId: slide.slideMaster.id,
          layoutId: slide.layout.id,
        }
        Object.values(result).forEach(hostId)
        if (result.slideId !== ownedId) throw Error('proposal_stale')
        return result
      }
      slide.load('id,slideMaster/id,layout/id')
      await sync(context, signal)
      const before = structuredClone(identity())
      if (typeof slide.exportAsBase64 !== 'function') throw Error('office_api_unsupported')
      const exported = slide.exportAsBase64()
      await sync(context, signal)
      const base64 = exported.value
      if (typeof base64 !== 'string') throw Error('office_read_failed')
      const digest = await presentationPackageDigest(base64, signal)
      slide.load('id,slideMaster/id,layout/id')
      await sync(context, signal)
      if (!same(before, identity())) throw Error('proposal_stale')
      return { ...before, digest, base64 }
    })
  }
  async inspect(exportIds: string[] = [], signal?: AbortSignal): Promise<MasterXmlHostSnapshot> {
    const owned = structuredClone(exportIds)
    owned.forEach(hostId)
    check(signal)
    return runtime().run((context) => inspect(context, owned, signal))
  }
  async stage(
    request: MasterXmlStageRequest,
    onInserted: (actual: MasterXmlInsertedIdentity) => Promise<void>,
    beforeWrite: () => Promise<void>,
    writeGuard: () => void,
    signal?: AbortSignal,
  ): Promise<MasterXmlInsertedIdentity> {
    const owned = structuredClone(request)
    valid(owned.preimage)
    hostId(owned.sourceSlideId)
    if (
      !owned.preimage.slideIds.includes(owned.sourceSlideId) ||
      (await readPackageSourceSlideId(owned.base64, signal)) !== owned.packageSourceSlideId
    )
      throw Error('invalid_tool_input')
    await beforeWrite()
    check(signal)
    return runtime().run(async (context) => {
      await preflight(context, owned.preimage, signal)
      if (typeof context.presentation.insertSlidesFromBase64 !== 'function')
        throw Error('office_api_unsupported')
      check(signal)
      writeGuard()
      context.presentation.insertSlidesFromBase64(owned.base64, {
        targetSlideId: owned.sourceSlideId,
        sourceSlideIds: [owned.packageSourceSlideId],
        formatting: 'KeepSourceFormatting',
      })
      try {
        await sync(context, signal)
      } catch (cause) {
        throw Error('office_state_uncertain', { cause })
      }
      const current = await native(context, signal),
        added = current.slideIds.filter((id) => !owned.preimage.slideIds.includes(id)),
        expected = [...owned.preimage.slideIds]
      if (added.length !== 1) throw Error('office_state_uncertain')
      expected.splice(expected.indexOf(owned.sourceSlideId) + 1, 0, added[0]!)
      if (!same(expected, current.slideIds)) throw Error('office_state_uncertain')
      const actual = current.dependencies.find((page) => page.slideId === added[0])!
      if (
        !actual ||
        !current.masters.some(
          (master) =>
            master.masterId === actual.masterId &&
            master.layouts.some((layout) => layout.layoutId === actual.layoutId),
        )
      )
        throw Error('office_state_uncertain')
      await onInserted(structuredClone(actual))
      check(signal)
      return actual
    })
  }
  async remove(
    request: MasterXmlRemoveRequest,
    beforeWrite: () => Promise<void>,
    writeGuard: () => void,
    signal?: AbortSignal,
  ): Promise<void> {
    const owned = structuredClone(request)
    valid(owned.preimage)
    hostId(owned.slideId)
    if (!owned.preimage.slideIds.includes(owned.slideId) || owned.preimage.slideIds.length < 2)
      throw Error('invalid_tool_input')
    await beforeWrite()
    check(signal)
    await runtime().run(async (context) => {
      await preflight(context, owned.preimage, signal)
      const slide = context.presentation.slides.getItem(owned.slideId)
      if (typeof slide.delete !== 'function') throw Error('office_api_unsupported')
      check(signal)
      writeGuard()
      slide.delete()
      try {
        await sync(context, signal)
      } catch (cause) {
        throw Error('office_state_uncertain', { cause })
      }
      if (
        !same(
          await readPowerPointPackageHostOrder(context, signal),
          owned.preimage.slideIds.filter((id) => id !== owned.slideId),
        )
      )
        throw Error('office_state_uncertain')
    })
  }
  async applyLayout(
    request: MasterXmlApplyLayoutRequest,
    beforeWrite: () => Promise<void>,
    writeGuard: () => void,
    signal?: AbortSignal,
  ): Promise<void> {
    const owned = structuredClone(request)
    valid(owned.preimage)
    hostId(owned.slideId)
    hostId(owned.masterId)
    hostId(owned.layoutId)
    if (
      !owned.preimage.slideIds.includes(owned.slideId) ||
      !owned.preimage.masters.some(
        (master) =>
          master.masterId === owned.masterId &&
          master.layouts.some((layout) => layout.layoutId === owned.layoutId),
      )
    )
      throw Error('invalid_tool_input')
    await beforeWrite()
    check(signal)
    await runtime().run(async (context) => {
      await preflight(context, owned.preimage, signal)
      const slide = context.presentation.slides.getItem(owned.slideId),
        layout = context.presentation.slideMasters
          .getItem(owned.masterId)
          .layouts.getItem(owned.layoutId)
      if (typeof slide.applyLayout !== 'function') throw Error('office_api_unsupported')
      check(signal)
      writeGuard()
      slide.applyLayout(layout)
      try {
        await sync(context, signal)
      } catch (cause) {
        throw Error('office_state_uncertain', { cause })
      }
      const actual = await native(context, signal),
        page = actual.dependencies.find((page) => page.slideId === owned.slideId)
      if (
        !same(actual.slideIds, owned.preimage.slideIds) ||
        page?.masterId !== owned.masterId ||
        page.layoutId !== owned.layoutId
      )
        throw Error('office_state_uncertain')
    })
  }
}
