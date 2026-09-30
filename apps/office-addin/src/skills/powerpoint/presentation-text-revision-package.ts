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

type Run = {
  xml: string
  text: string
  globalStart: number
  openTag: string
  openStart: number
  openEnd: number
  textStart: number
  textEnd: number
}

function plainRuns(shapeXml: string): { runs: Run[]; text: string } {
  const bodies = [...shapeXml.matchAll(/<p:txBody\b[^>]*>[\s\S]*?<\/p:txBody>/g)]
  if (bodies.length !== 1) fail()
  const body = bodies[0]![0]
  if (/<a:(?:fld|br|tab|hlinkClick|hlinkMouseOver)\b/.test(body)) fail()
  const paragraphs = [...body.matchAll(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g)]
  if (
    !paragraphs.length ||
    paragraphs.length > 2000 ||
    paragraphs.length !== [...body.matchAll(/<a:p\b/g)].length
  )
    fail()
  const result: Run[] = []
  let fullText = ''
  for (const [paragraphIndex, paragraph] of paragraphs.entries()) {
    if (paragraphIndex) fullText += '\n'
    const paragraphRuns = [...paragraph[0].matchAll(/<a:r\b[^>]*>[\s\S]*?<\/a:r>/g)]
    if ([...paragraph[0].matchAll(/<a:r\b/g)].length !== paragraphRuns.length) fail()
    for (const match of paragraphRuns) {
      const xml = match[0]
      const texts = [...xml.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)]
      if (texts.length !== 1 || /<!\[CDATA\[/.test(texts[0]![1]!)) fail()
      const parsed = parser.parse(texts[0]![0]) as Record<string, unknown>
      const parsedText = parsed['a:t']
      const value =
        typeof parsedText === 'string'
          ? parsedText
          : (parsedText as Record<string, unknown> | undefined)?.['#text']
      if (typeof value !== 'string' || invalidText(value)) fail()
      const openTag = texts[0]![0].slice(0, texts[0]![0].indexOf('>') + 1)
      const openStart = bodies[0]!.index! + paragraph.index! + match.index! + texts[0]!.index!
      result.push({
        xml,
        text: value as string,
        globalStart: fullText.length,
        openTag,
        openStart,
        openEnd: openStart + openTag.length,
        textStart: openStart + openTag.length,
        textEnd:
          bodies[0]!.index! +
          paragraph.index! +
          match.index! +
          texts[0]!.index! +
          texts[0]![0].lastIndexOf('</a:t>'),
      })
      fullText += value
    }
    if (fullText.length > 12_000) fail()
    const parsedParagraph = parser.parse(paragraph[0]) as Record<string, unknown>
    const parsedRuns = (parsedParagraph['a:p'] as Record<string, unknown> | undefined)?.['a:r']
    if (
      Array.isArray(parsedRuns)
        ? parsedRuns.length !== paragraphRuns.length
        : paragraphRuns.length !== (parsedRuns === undefined ? 0 : 1)
    )
      fail()
  }
  if (!result.length || result.length > 2000) fail()
  if ([...body.matchAll(/<a:t(?:\s[^>]*)?>/g)].length !== result.length) fail()
  return { runs: result, text: fullText }
}

/** Revise native text in one paragraph, counting paragraph boundaries as newlines in start. */
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
  const { runs, text: fullText } = plainRuns(original)
  if (fullText.length > 12_000) fail()
  if (
    start + before.length > fullText.length ||
    fullText.slice(start, start + before.length) !== before
  )
    throw new Error('presentation_baseline_changed')
  if (fullText.length - before.length + after.length > 12_000) fail()
  let changedRuns = 0
  let rewritten = original
  const patches: Array<{ from: number; to: number; value: string }> = []
  const patchRun = (run: Run, replacement: string) => {
    if (invalidText(replacement)) fail()
    if (/^[ \t]|[ \t]$/.test(replacement)) {
      const space = /\bxml:space\s*=\s*(['"])([^'"]*)\1/.exec(run.openTag)
      if (space?.[2] !== 'preserve')
        patches.push({
          from: run.openStart,
          to: run.openEnd,
          value: space
            ? run.openTag.replace(space[0], 'xml:space="preserve"')
            : `${run.openTag.slice(0, -1)} xml:space="preserve">`,
        })
    }
    patches.push({ from: run.textStart, to: run.textEnd, value: escape(replacement) })
    changedRuns++
  }
  if (before.length !== after.length) {
    if (runReplacements) {
      for (const run of runs) {
        const first = Math.max(start, run.globalStart)
        const last = Math.min(start + before.length, run.globalStart + run.text.length)
        if (first < last) {
          const part = runReplacements[changedRuns]
          if (!part) throw new Error('invalid_tool_input')
          const replacement =
            run.text.slice(0, first - run.globalStart) +
            part +
            run.text.slice(last - run.globalStart)
          patchRun(run, replacement)
        }
      }
      if (changedRuns !== runReplacements.length) throw new Error('invalid_tool_input')
    } else
      for (const run of runs) {
        if (
          start >= run.globalStart &&
          start + before.length <= run.globalStart + run.text.length
        ) {
          const local = start - run.globalStart
          const replacement =
            run.text.slice(0, local) + after + run.text.slice(local + before.length)
          patchRun(run, replacement)
          break
        }
      }
    if (!changedRuns) fail()
  } else {
    for (const run of runs) {
      const first = Math.max(start, run.globalStart)
      const last = Math.min(start + before.length, run.globalStart + run.text.length)
      if (first < last) {
        const local = first - run.globalStart
        const replacement =
          run.text.slice(0, local) +
          after.slice(first - start, last - start) +
          run.text.slice(last - run.globalStart)
        patchRun(run, replacement)
      }
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
    plainRuns(nextShape[0]![0]).text !==
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
