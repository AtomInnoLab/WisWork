import { XMLParser, XMLValidator } from 'fast-xml-parser'
import {
  capturePowerPointPackage,
  loadBoundedZip,
  MAX_PPTX_XML_BYTES,
  presentationPackageDigest,
} from './powerpoint-package.js'

export interface PackageHostSnapshot {
  slideIds: string[]
  pages: { slideId: string; digest: string; base64?: string }[]
}
export interface PackageStageRequest {
  base64: string
  sourceSlideId: string
  packageSourceSlideId: string
  preimage: PackageHostSnapshot
}
export interface PackageRemoveRequest {
  slideId: string
  preimage: PackageHostSnapshot
}
const MAX_SNAPSHOT_METADATA_BYTES = 4 * 1024 * 1024
function check(signal?: AbortSignal) {
  if (signal?.aborted) throw Error('cancelled')
}
function id(value: unknown): asserts value is string {
  if (
    typeof value !== 'string' ||
    !value.trim() ||
    value.length > 256 ||
    Array.from(value).some(
      (char) => char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
    )
  )
    throw Error('invalid_tool_input')
}
function snapshot(value: PackageHostSnapshot): PackageHostSnapshot {
  const owned = structuredClone(value)
  if (
    !Array.isArray(owned.slideIds) ||
    !owned.slideIds.length ||
    !Array.isArray(owned.pages) ||
    owned.pages.length !== owned.slideIds.length ||
    new TextEncoder().encode(
      JSON.stringify({
        slideIds: owned.slideIds,
        pages: owned.pages.map(({ slideId, digest }) => ({ slideId, digest })),
      }),
    ).byteLength > MAX_SNAPSHOT_METADATA_BYTES
  )
    throw Error('invalid_tool_input')
  owned.slideIds.forEach(id)
  if (
    new Set(owned.slideIds).size !== owned.slideIds.length ||
    owned.pages.some(
      (page, i) => page.slideId !== owned.slideIds[i] || !/^[a-f0-9]{64}$/.test(page.digest),
    )
  )
    throw Error('invalid_tool_input')
  return owned
}
function runtime(): typeof PowerPoint {
  if (
    String(globalThis.Office?.context?.host) !== 'PowerPoint' ||
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
async function order(context: PowerPoint.RequestContext, signal?: AbortSignal): Promise<string[]> {
  const slides = context.presentation.slides
  if (typeof slides.getCount !== 'function' || typeof slides.load !== 'function')
    throw Error('office_api_unsupported')
  const count = slides.getCount()
  slides.load('items/id')
  await sync(context, signal)
  const ids = slides.items.map((slide) => slide.id)
  if (
    !Number.isSafeInteger(count.value) ||
    count.value < 1 ||
    count.value !== ids.length ||
    new TextEncoder().encode(JSON.stringify(ids)).byteLength > MAX_SNAPSHOT_METADATA_BYTES
  )
    throw Error('office_read_failed')
  ids.forEach(id)
  if (new Set(ids).size !== ids.length) throw Error('office_read_failed')
  return ids
}
const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b)
async function inspect(
  context: PowerPoint.RequestContext,
  exports: readonly string[],
  signal?: AbortSignal,
): Promise<PackageHostSnapshot> {
  const slideIds = await order(context, signal)
  if (exports.some((target) => !slideIds.includes(target))) throw Error('invalid_tool_input')
  const pages: PackageHostSnapshot['pages'] = []
  for (const slideId of slideIds) {
    const slide = context.presentation.slides.getItem(slideId)
    if (typeof slide.exportAsBase64 !== 'function') throw Error('office_api_unsupported')
    const result = slide.exportAsBase64()
    await sync(context, signal)
    const base64 = result.value
    if (typeof base64 !== 'string') throw Error('office_read_failed')
    const digest = await presentationPackageDigest(base64, signal)
    pages.push({ slideId, digest, ...(exports.includes(slideId) ? { base64 } : {}) })
  }
  if (!same(await order(context, signal), slideIds)) throw Error('proposal_stale')
  return { slideIds, pages }
}
export async function readPackageSourceSlideId(
  base64: string,
  signal?: AbortSignal,
): Promise<string> {
  const zip = await loadBoundedZip(base64, signal)
  await capturePowerPointPackage(base64, signal)
  const xml = await zip.file('ppt/presentation.xml')?.async('string')
  if (
    !xml ||
    new TextEncoder().encode(xml).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(DOCTYPE|ENTITY)\b/i.test(xml) ||
    XMLValidator.validate(xml) !== true
  )
    throw Error('invalid_tool_input')
  const parser = new XMLParser({
    ignoreAttributes: false,
    parseAttributeValue: false,
    processEntities: false,
  })
  const entry = parser.parse(xml)?.['p:presentation']?.['p:sldIdLst']?.['p:sldId']
  if (
    !entry ||
    Array.isArray(entry) ||
    !/^[1-9][0-9]{0,9}$/.test(entry['@_id']) ||
    Number(entry['@_id']) < 256 ||
    Number(entry['@_id']) > 4294967295
  )
    throw Error('invalid_tool_input')
  const relXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string')
  if (
    !relXml ||
    new TextEncoder().encode(relXml).byteLength > MAX_PPTX_XML_BYTES ||
    /<!\s*(DOCTYPE|ENTITY)\b/i.test(relXml) ||
    XMLValidator.validate(relXml) !== true
  )
    throw Error('invalid_tool_input')
  const all = parser.parse(relXml)?.Relationships?.Relationship
  const rels = Array.isArray(all) ? all : [all]
  const slidePaths = Object.keys(zip.files).filter((path) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(path),
  )
  const matches = rels.filter(
    (rel) =>
      rel?.['@_Id'] === entry['@_r:id'] &&
      rel['@_Type'] ===
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' &&
      rel['@_Target'] === `slides/${slidePaths[0]?.split('/').at(-1)}` &&
      (rel['@_TargetMode'] === undefined || rel['@_TargetMode'] === 'Internal'),
  )
  if (slidePaths.length !== 1 || matches.length !== 1) throw Error('invalid_tool_input')
  return `${entry['@_id']}#`
}
async function preflight(
  context: PowerPoint.RequestContext,
  before: PackageHostSnapshot,
  signal?: AbortSignal,
) {
  const current = await inspect(context, [], signal)
  if (
    !same(current.slideIds, before.slideIds) ||
    current.pages.some((page, i) => page.digest !== before.pages[i]!.digest)
  )
    throw Error('proposal_stale')
}
/** These guarded SDK writes do not provide host-side atomic CAS. Unknown acknowledgements remain unknown. */
export class BrowserPresentationPackageEditAdapter {
  async inspect(exportIds: string[] = [], signal?: AbortSignal): Promise<PackageHostSnapshot> {
    const exports = structuredClone(exportIds)
    exports.forEach(id)
    check(signal)
    return runtime().run((context) => inspect(context, exports, signal))
  }
  async stage(
    request: PackageStageRequest,
    onInserted: (slideId: string) => Promise<void>,
    beforeWrite: () => Promise<void>,
    writeGuard: () => void,
    signal?: AbortSignal,
  ): Promise<{ slideId: string }> {
    const owned = structuredClone(request),
      before = snapshot(owned.preimage)
    id(owned.sourceSlideId)
    if (
      !before.slideIds.includes(owned.sourceSlideId) ||
      (await readPackageSourceSlideId(owned.base64, signal)) !== owned.packageSourceSlideId
    )
      throw Error('invalid_tool_input')
    await beforeWrite()
    check(signal)
    return runtime().run(async (context) => {
      await preflight(context, before, signal)
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
      const after = await order(context, signal),
        added = after.filter((value) => !before.slideIds.includes(value))
      const expected = [...before.slideIds]
      if (added.length !== 1) throw Error('office_state_uncertain')
      expected.splice(expected.indexOf(owned.sourceSlideId) + 1, 0, added[0]!)
      if (!same(after, expected)) throw Error('office_state_uncertain')
      await onInserted(added[0]!)
      check(signal)
      return { slideId: added[0]! }
    })
  }
  async remove(
    request: PackageRemoveRequest,
    beforeWrite: () => Promise<void>,
    writeGuard: () => void,
    signal?: AbortSignal,
  ): Promise<void> {
    const owned = structuredClone(request),
      before = snapshot(owned.preimage)
    id(owned.slideId)
    if (!before.slideIds.includes(owned.slideId) || before.slideIds.length < 2)
      throw Error('invalid_tool_input')
    await beforeWrite()
    check(signal)
    await runtime().run(async (context) => {
      await preflight(context, before, signal)
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
          await order(context, signal),
          before.slideIds.filter((value) => value !== owned.slideId),
        )
      )
        throw Error('office_state_uncertain')
    })
  }
}
