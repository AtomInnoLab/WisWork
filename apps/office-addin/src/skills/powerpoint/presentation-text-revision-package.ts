import { XMLParser, XMLValidator } from 'fast-xml-parser'
import {
  loadBoundedZip,
  MAX_PPTX_PACKAGE_BYTES,
  MAX_PPTX_XML_BYTES,
  presentationPackageDigest,
} from './powerpoint-package.js'

const parser = new XMLParser({
  ignoreAttributes: false,
  parseTagValue: false,
  trimValues: false,
})
const fail = (): never => {
  throw new Error('presentation_existing_target_unsupported')
}
const escape = (value: string) =>
  value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
const invalidText = (value: string) => /[\r\n]/.test(value) || /[\uD800-\uDFFF]/u.test(value)

type Run = { xml: string; text: string; textStart: number; textEnd: number }

function plainRuns(shapeXml: string): Run[] {
  const bodies = [...shapeXml.matchAll(/<p:txBody\b[^>]*>[\s\S]*?<\/p:txBody>/g)]
  if (bodies.length !== 1) fail()
  const body = bodies[0]![0]
  if (
    [...body.matchAll(/<a:p\b[^>]*>/g)].length !== 1 ||
    /<a:(?:fld|br|tab|hlinkClick|hlinkMouseOver)\b/.test(body)
  )
    fail()
  const result: Run[] = []
  for (const match of body.matchAll(/<a:r\b[^>]*>[\s\S]*?<\/a:r>/g)) {
    const xml = match[0]
    const texts = [...xml.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)]
    if (texts.length !== 1 || /<!\[CDATA\[/.test(texts[0]![1]!)) fail()
    const parsed = parser.parse(texts[0]![0]) as Record<string, unknown>
    const value = parsed['a:t']
    if (typeof value !== 'string' || invalidText(value)) fail()
    result.push({
      xml,
      text: value as string,
      textStart:
        bodies[0]!.index! + match.index! + texts[0]!.index! + texts[0]![0].indexOf('>') + 1,
      textEnd:
        bodies[0]!.index! + match.index! + texts[0]!.index! + texts[0]![0].lastIndexOf('</a:t>'),
    })
  }
  if (!result.length || result.length > 2000) fail()
  if ([...body.matchAll(/<a:t(?:\s[^>]*)?>/g)].length !== result.length) fail()
  const parsedBody = parser.parse(body) as Record<string, unknown>
  const paragraphs = (parsedBody['p:txBody'] as Record<string, unknown> | undefined)?.['a:p'] as
    Record<string, unknown> | undefined
  const parsedRuns = paragraphs?.['a:r']
  if (Array.isArray(parsedRuns) ? parsedRuns.length !== result.length : result.length !== 1) fail()
  return result
}

/** Prepare an equal-length native text revision while retaining each run's formatting. */
export async function replacePowerPointTextRangePackage(
  source: string,
  shapeId: string,
  start: number,
  before: string,
  after: string,
  signal?: AbortSignal,
  runReplacements?: string[],
): Promise<{ base64: string; beforeDigest: string; afterDigest: string; changedRuns: number }> {
  if (
    !/^[1-9]\d{0,9}$/.test(shapeId) ||
    !Number.isSafeInteger(start) ||
    start < 0 ||
    typeof before !== 'string' ||
    typeof after !== 'string' ||
    !before.length ||
    before.length > 128 ||
    !after.length ||
    after.length > 128 ||
    before === after ||
    invalidText(before) ||
    invalidText(after) ||
    (runReplacements !== undefined &&
      (before.length === after.length ||
        !Array.isArray(runReplacements) ||
        !runReplacements.length ||
        runReplacements.some(
          (value) => typeof value !== 'string' || !value || invalidText(value),
        ) ||
        runReplacements.join('') !== after))
  )
    throw new Error('invalid_tool_input')
  if (signal?.aborted) throw new Error('cancelled')
  const zip = await loadBoundedZip(source, signal, true, 8 * 1024 * 1024)
  const slides = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
  if (slides.length !== 1) fail()
  const slidePath = slides[0]!
  const slideXml = await zip.file(slidePath)!.async('string')
  if (
    new TextEncoder().encode(slideXml).byteLength > MAX_PPTX_XML_BYTES ||
    XMLValidator.validate(slideXml) !== true
  )
    fail()
  const shapes = [...slideXml.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].filter(([xml]) =>
    new RegExp(`<p:cNvPr\\b[^>]*\\bid=["']${shapeId}["']`).test(xml),
  )
  if (shapes.length !== 1) fail()
  const original = shapes[0]![0]
  const runs = plainRuns(original)
  const fullText = runs.map((run) => run.text).join('')
  if (fullText.length > 12_000) fail()
  if (
    start + before.length > fullText.length ||
    fullText.slice(start, start + before.length) !== before
  )
    throw new Error('presentation_baseline_changed')
  if (fullText.length - before.length + after.length > 12_000) fail()
  let cursor = 0
  let changedRuns = 0
  let rewritten = original
  const patches: Array<{ from: number; to: number; value: string }> = []
  if (before.length !== after.length) {
    if (runReplacements) {
      for (const run of runs) {
        const first = Math.max(start, cursor)
        const last = Math.min(start + before.length, cursor + run.text.length)
        if (first < last) {
          const part = runReplacements[changedRuns]
          if (!part) throw new Error('invalid_tool_input')
          const replacement =
            run.text.slice(0, first - cursor) + part + run.text.slice(last - cursor)
          if (invalidText(replacement)) fail()
          patches.push({ from: run.textStart, to: run.textEnd, value: escape(replacement) })
          changedRuns++
        }
        cursor += run.text.length
      }
      if (changedRuns !== runReplacements.length) throw new Error('invalid_tool_input')
    } else
      for (const run of runs) {
        if (start >= cursor && start + before.length <= cursor + run.text.length) {
          const local = start - cursor
          const replacement =
            run.text.slice(0, local) + after + run.text.slice(local + before.length)
          if (invalidText(replacement)) fail()
          patches.push({ from: run.textStart, to: run.textEnd, value: escape(replacement) })
          changedRuns++
          break
        }
        cursor += run.text.length
      }
    if (!changedRuns) fail()
  } else {
    for (const run of runs) {
      const first = Math.max(start, cursor)
      const last = Math.min(start + before.length, cursor + run.text.length)
      if (first < last) {
        const local = first - cursor
        const replacement =
          run.text.slice(0, local) +
          after.slice(first - start, last - start) +
          run.text.slice(last - cursor)
        if (invalidText(replacement)) fail()
        patches.push({ from: run.textStart, to: run.textEnd, value: escape(replacement) })
        changedRuns++
      }
      cursor += run.text.length
    }
  }
  if (!changedRuns) fail()
  for (const patch of patches.reverse())
    rewritten = rewritten.slice(0, patch.from) + patch.value + rewritten.slice(patch.to)
  const nextSlide =
    slideXml.slice(0, shapes[0]!.index!) +
    rewritten +
    slideXml.slice(shapes[0]!.index! + original.length)
  if (XMLValidator.validate(nextSlide) !== true) fail()
  const nextShape = [...nextSlide.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g)].filter(([xml]) =>
    new RegExp(`<p:cNvPr\\b[^>]*\\bid=["']${shapeId}["']`).test(xml),
  )
  if (
    nextShape.length !== 1 ||
    plainRuns(nextShape[0]![0])
      .map((run) => run.text)
      .join('') !==
      fullText.slice(0, start) + after + fullText.slice(start + before.length)
  )
    fail()
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
  return { base64, beforeDigest, afterDigest, changedRuns }
}
