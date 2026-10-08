import { afterEach, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../../shell/src/main/presentation-service'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { createPresentationNativeModifySkill } from '../src/skills/powerpoint/presentation-native-modify'
import {
  inspectPowerPointRichText,
  inspectPowerPointTextShapeFingerprints,
} from '../src/skills/powerpoint/presentation-rich-text-package'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
import type { createStructuredProposalController } from '../src/agent/proposal-controller'
const durableRoots: string[] = []
afterEach(() => {
  for (const root of durableRoots.splice(0)) rmSync(root, { recursive: true, force: true })
})
export async function durableCompatibility(
  fake: PowerPointAdapter,
  proposals: ReturnType<typeof createStructuredProposalController>,
) {
  fake.inspectSlideNativePackage = undefined
  const root = mkdtempSync(join(tmpdir(), 'ppt-compat-durable-'))
  durableRoots.push(root)
  const service = createPresentationService({ userDataPath: root })
  const initial = await fake.listSlideShapes(0)
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[0]!]
  deck.slides[0]!.elements = initial.shapes.map((shape) => ({
    kind: 'text' as const,
    id: shape.name.replace(/[^A-Za-z0-9_-]/g, '_'),
    x: shape.left / 72,
    y: shape.top / 72,
    w: shape.width / 72,
    h: shape.height / 72,
    text: 'Hello',
    fontSize: 20,
  }))
  const { bytes } = await compilePresentationDeck(deck),
    zip = await JSZip.loadAsync(bytes)
  let base64 = await zip.generateAsync({ type: 'base64' })
  const originalWrite = fake.executeDeclarative.bind(fake)
  fake.executeDeclarative = vi.fn(async (ops: any[], signal?: AbortSignal) => {
    const result = await originalWrite(ops, signal)
    for (const op of ops) {
      const shape = initial.shapes.find((shape) => shape.id === op.shape_id)!
      const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
      const node = [...xml.matchAll(/<p:sp\b[^]*?<\/p:sp>/g)]
        .map((match) => match[0])
        .find((node) => node.includes(`name="${shape.name.replace(/[^A-Za-z0-9_-]/g, '_')}"`))
      if (!node) throw Error('fixture_native_shape_missing')
      const replacement =
        op.op === 'delete_shape'
          ? ''
          : op.op === 'set_shape_text'
            ? node.replace(/<a:t>[^]*?<\/a:t>/, `<a:t>${op.text}</a:t>`)
            : op.op === 'set_shape_text_style'
              ? node
                  .replace(/<a:rPr\b([^>]*?)(\/?)>/g, (_match, attrs, close) => {
                    const values = {
                      ...(op.fontSize === undefined ? {} : { sz: Math.round(op.fontSize * 100) }),
                      ...(op.bold === undefined ? {} : { b: op.bold ? 1 : 0 }),
                      ...(op.italic === undefined ? {} : { i: op.italic ? 1 : 0 }),
                    }
                    for (const [key, value] of Object.entries(values))
                      attrs =
                        attrs.replace(new RegExp(`\\s${key}="[^"]*"`, 'g'), '') +
                        ` ${key}="${value}"`
                    return `<a:rPr${attrs}${close}>`
                  })
                  .replace(/<a:rPr\b([^>]*)>([\s\S]*?)<\/a:rPr>/g, (_match, attrs, body) => {
                    if (op.color)
                      body =
                        body.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/g, '') +
                        `<a:solidFill><a:srgbClr val="${op.color.slice(1)}"/></a:solidFill>`
                    if (op.fontFamily)
                      body =
                        body.replace(/<a:latin\b[^>]*\/>/g, '') +
                        `<a:latin typeface="${op.fontFamily}"/>`
                    return `<a:rPr${attrs}>${body}</a:rPr>`
                  })
              : node
                  .replace(
                    /<a:off\b[^>]*\/>/,
                    `<a:off x="${Math.round(op.left * 12700)}" y="${Math.round(op.top * 12700)}"/>`,
                  )
                  .replace(
                    /<a:ext\b[^>]*\/>/,
                    `<a:ext cx="${Math.round(op.width * 12700)}" cy="${Math.round(op.height * 12700)}"/>`,
                  )
      zip.file('ppt/slides/slide1.xml', xml.replace(node, replacement))
    }
    base64 = await zip.generateAsync({ type: 'base64' })
    return result
  })
  fake.snapshotSlide = vi.fn(async (_slideIndex: number, signal?: AbortSignal) => ({
    slideId: initial.slideId,
    fingerprint: 'initial-native-page',
    shapes: await Promise.all(
      (await fake.listSlideShapes(0, signal)).shapes.map(async (shape) => ({
        ...shape,
        text: await fake.readSlideText(0, shape.id, signal).then(
          (value) => value.text,
          () => '',
        ),
      })),
    ),
  }))
  fake.exportPresentationPagePackage = vi.fn(async (slideId: string) => ({
    slideId,
    slideIds: [initial.slideId],
    base64,
  }))
  fake.inspectSlideRichText = vi.fn(async (slideId: string, shapeIds: string[]) => {
    const parsed = await inspectPowerPointRichText(base64)
    const packageFingerprints = await inspectPowerPointTextShapeFingerprints(
      base64,
      parsed.shapes.map((item) => item.packageShapeId),
    )
    const matched = shapeIds.map((id) => {
      const name = initial.shapes
        .find((item) => item.id === id)!
        .name.replace(/[^A-Za-z0-9_-]/g, '_')
      const found = parsed.shapes.find((item) => item.name === name)
      if (!found) throw Error('office_api_unsupported')
      return [id, found] as const
    })
    return {
      slideId,
      slideIds: [initial.slideId],
      shapes: Object.fromEntries(matched),
      fingerprints: Object.fromEntries(
        matched.map(([id, item]) => [id, packageFingerprints[item.packageShapeId]]),
      ),
    }
  })
  const values = new Map<string, string>()
  const binding = createPresentationDocumentBinding(
    {
      get: (key) => values.get(key),
      set: (key, value) => {
        values.set(key, value)
      },
      save: async () => {},
      location: () => 'compatibility-deck',
    },
    () => 'compatibility-doc',
  )
  const generic = createPresentationNativeModifySkill({
    adapter: fake,
    proposals,
    available: () => true,
    documentId: () => binding.documentId(),
    request: async (body, signal) =>
      new Response(
        Buffer.from(await service(body, signal ?? new AbortController().signal)).toString('utf8'),
      ),
    readExistingBatch: (id) => binding.readExistingBatch(id),
    writeExistingBatch: (next, expected) => binding.writeExistingBatch(next, expected),
  })
  return {
    binding,
    skill: createPowerPointSkill({
      adapter: fake,
      proposals,
      durableModify: (operations, explanation, signal) =>
        generic.propose(operations, explanation, signal),
    }),
  }
}
