import { XMLValidator } from 'fast-xml-parser'
import { inspectPowerPointPicturePackage, loadBoundedZip, MAX_PPTX_PACKAGE_BYTES, MAX_PPTX_XML_BYTES, presentationPackageDigest } from './powerpoint-package.js'

type Image = { mime: 'image/png' | 'image/jpeg'; base64: string }
const imageRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
const fail = (): never => { throw new Error('office_api_unsupported') }

/** Replace one ordinary embedded picture's media without changing any other shape or relation. */
export async function replacePowerPointPictureMediaPackage(
  source: string,
  shapeId: string,
  image: Image,
  signal?: AbortSignal,
): Promise<{ base64: string; beforeDigest: string; afterDigest: string; mediaDigest: string }> {
  if (!/^[1-9]\d{0,9}$/.test(shapeId) || signal?.aborted) throw new Error(signal?.aborted ? 'cancelled' : 'invalid_tool_input')
  if (!image.base64 || image.base64.length > Math.ceil((2 * 1024 * 1024) / 3) * 4 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(image.base64)) throw new Error('invalid_tool_input')
  const bytes = Uint8Array.from(atob(image.base64), (value) => value.charCodeAt(0))
  if (!bytes.length || bytes.length > 2 * 1024 * 1024 ||
    (image.mime === 'image/png' ? !image.base64.startsWith('iVBORw0KGgo') :
      image.mime === 'image/jpeg' ? !(bytes[0] === 0xff && bytes[1] === 0xd8) : true)) throw new Error('invalid_tool_input')
  const original = await inspectPowerPointPicturePackage(source, shapeId, signal)
  const mediaDigest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (byte) => byte.toString(16).padStart(2, '0')).join('')
  if (mediaDigest === original.mediaDigest) throw new Error('invalid_tool_input')
  const zip = await loadBoundedZip(source, signal)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) fail()
  const slidePath = slides[0]!
  const relPath = slidePath.replace('/slides/', '/slides/_rels/') + '.rels'
  const slideXml = await zip.file(slidePath)!.async('string')
  const relXml = await zip.file(relPath)?.async('string')
  const typesXml = await zip.file('[Content_Types].xml')?.async('string')
  if (!relXml || !typesXml) throw new Error('office_api_unsupported')
  if ([slideXml, relXml, typesXml].some((xml) => xml.length > MAX_PPTX_XML_BYTES || XMLValidator.validate(xml) !== true)) fail()
  const pictures = [...slideXml.matchAll(/<p:pic\b[^>]*>[\s\S]*?<\/p:pic>/g)]
  const target = pictures.filter(([xml]) => new RegExp(`<p:cNvPr\\b[^>]*\\bid=["']${shapeId}["']`).test(xml))
  if (target.length !== 1) fail()
  const oldEmbed = [...target[0]![0].matchAll(/<a:blip\b[^>]*\br:embed=["']([^"']+)["']/g)]
  if (oldEmbed.length !== 1) fail()
  const relationIds = [...relXml.matchAll(/\bId=["']rId(\d+)["']/g)].map((match) => Number(match[1]))
  if (relationIds.some((value) => !Number.isSafeInteger(value)) || relationIds.length > 1024) fail()
  const rid = `rId${Math.max(0, ...relationIds) + 1}`
  const ext = image.mime === 'image/png' ? 'png' : 'jpeg'
  const mediaName = `wiswork-picture-${crypto.randomUUID()}.${ext}`
  const mediaPath = `ppt/media/${mediaName}`
  if (zip.file(mediaPath)) fail()
  const rewrittenPicture = target[0]![0].replace(`r:embed="${oldEmbed[0]![1]}"`, `r:embed="${rid}"`)
  if (rewrittenPicture === target[0]![0]) fail()
  const nextSlide = slideXml.replace(target[0]![0], rewrittenPicture)
  const relation = `<Relationship Id="${rid}" Type="${imageRel}" Target="../media/${mediaName}"/>`
  const nextRels = relXml.replace('</Relationships>', `${relation}</Relationships>`)
  if (nextRels === relXml) fail()
  let nextTypes = typesXml
  if (!new RegExp(`<Default\\b[^>]*\\bExtension=["']${ext}["']`, 'i').test(typesXml))
    nextTypes = typesXml.replace('</Types>', `<Default Extension="${ext}" ContentType="${image.mime}"/></Types>`)
  if ([nextSlide, nextRels, nextTypes].some((xml) => xml.length > MAX_PPTX_XML_BYTES || XMLValidator.validate(xml) !== true)) fail()
  zip.file(slidePath, nextSlide)
  zip.file(relPath, nextRels)
  zip.file('[Content_Types].xml', nextTypes)
  zip.file(mediaPath, bytes)
  const base64 = await zip.generateAsync({ type: 'base64', compression: 'DEFLATE' })
  if (signal?.aborted) throw new Error('cancelled')
  if (base64.length > Math.ceil(MAX_PPTX_PACKAGE_BYTES / 3) * 4) throw new Error('invalid_tool_input')
  await loadBoundedZip(base64, signal)
  const updated = await inspectPowerPointPicturePackage(base64, shapeId, signal)
  if (updated.mediaDigest !== mediaDigest || JSON.stringify(updated.shapeIds) !== JSON.stringify(original.shapeIds)) fail()
  const [beforeDigest, afterDigest] = await Promise.all([presentationPackageDigest(source, signal), presentationPackageDigest(base64, signal)])
  if (beforeDigest === afterDigest) fail()
  return { base64, beforeDigest, afterDigest, mediaDigest }
}
