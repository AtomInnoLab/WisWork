import { XMLValidator } from 'fast-xml-parser'
import {
  loadBoundedZip,
  MAX_PPTX_PACKAGE_BYTES,
  MAX_PPTX_XML_BYTES,
  presentationPackageDigest,
} from './powerpoint-package.js'

export interface NativePageGeometry {
  left: number
  top: number
  width: number
  height: number
}

const emu = (points: number) => Math.round(points * 12700)
const fail = (): never => {
  throw new Error('presentation_existing_target_unsupported')
}
const valid = (value: NativePageGeometry): boolean =>
  value !== null &&
  typeof value === 'object' &&
  Object.keys(value).sort().join(',') === 'height,left,top,width' &&
  [value.left, value.top, value.width, value.height].every(
    (part) => Number.isFinite(part) && part >= 0 && part <= 1000,
  ) &&
  value.width > 0 &&
  value.height > 0 &&
  value.left + value.width <= 1000 &&
  value.top + value.height <= 1000

/** Revise an ordinary, unrotated native shape's OOXML transform in a one-slide package. */
export async function replacePowerPointShapeGeometryPackage(
  source: string,
  shapeId: string,
  before: NativePageGeometry,
  after: NativePageGeometry,
  signal?: AbortSignal,
): Promise<{ base64: string; beforeDigest: string; afterDigest: string }> {
  if (!/^[1-9]\d{0,9}$/.test(shapeId) || !valid(before) || !valid(after))
    throw new Error('invalid_tool_input')
  if (signal?.aborted) throw new Error('cancelled')
  if (JSON.stringify(before) === JSON.stringify(after)) throw new Error('invalid_tool_input')
  const zip = await loadBoundedZip(source, signal, true, MAX_PPTX_PACKAGE_BYTES)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) fail()
  const slidePath = slides[0]!
  const slideXml = await zip.file(slidePath)!.async('string')
  if (
    new TextEncoder().encode(slideXml).byteLength > MAX_PPTX_XML_BYTES ||
    XMLValidator.validate(slideXml) !== true
  )
    fail()
  const targets = [...slideXml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].filter(([xml]) =>
    new RegExp(`<p:cNvPr\\b[^>]*\\bid=["']${shapeId}["']`).test(xml),
  )
  if (targets.length !== 1) fail()
  const original = targets[0]![0]
  const properties = [...original.matchAll(/<p:spPr\b[^>]*>[\s\S]*?<\/p:spPr>/g)]
  if (properties.length !== 1) fail()
  const transforms = [...properties[0]![0].matchAll(/<a:xfrm\b[^>]*>[\s\S]*?<\/a:xfrm>/g)]
  if (transforms.length !== 1) fail()
  const transform = transforms[0]![0]
  const values =
    /^<a:xfrm>(\s*<a:off x="(\d+)" y="(\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>\s*)<\/a:xfrm>$/.exec(
      transform,
    )
  if (!values) throw new Error('presentation_existing_target_unsupported')
  const actual = [Number(values[2]), Number(values[3]), Number(values[4]), Number(values[5])]
  if (actual.some((part) => !Number.isSafeInteger(part))) fail()
  if (
    actual.some(
      (part, index) =>
        Math.abs(part - emu([before.left, before.top, before.width, before.height][index]!)) > 127,
    )
  )
    throw new Error('presentation_baseline_changed')
  const revised = `<a:xfrm><a:off x="${emu(after.left)}" y="${emu(after.top)}"/><a:ext cx="${emu(after.width)}" cy="${emu(after.height)}"/></a:xfrm>`
  const targetIndex = targets[0]!.index!
  const nextShape = original.replace(transform, revised)
  const nextSlide =
    slideXml.slice(0, targetIndex) + nextShape + slideXml.slice(targetIndex + original.length)
  if (nextSlide === slideXml || XMLValidator.validate(nextSlide) !== true) fail()
  zip.file(slidePath, nextSlide)
  const base64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' })
  if (signal?.aborted) throw new Error('cancelled')
  if (base64.length > Math.ceil(MAX_PPTX_PACKAGE_BYTES / 3) * 4)
    throw new Error('invalid_tool_input')
  await loadBoundedZip(base64, signal)
  const [beforeDigest, afterDigest] = await Promise.all([
    presentationPackageDigest(source, signal),
    presentationPackageDigest(base64, signal),
  ])
  if (beforeDigest === afterDigest) fail()
  return { base64, beforeDigest, afterDigest }
}
