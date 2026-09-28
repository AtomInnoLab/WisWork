import { afterEach, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationExistingPageEditingSkill } from '../src/skills/powerpoint/presentation-existing-page-editing.js'
import { inspectPowerPointPicturePackage } from '../src/skills/powerpoint/powerpoint-package.js'

afterEach(() => {
  delete (globalThis as Record<string, unknown>).createImageBitmap
})

it('prepares a three-object revision for one confirmed page change without writing PowerPoint', async () => {
  const deck = benchmarkDeck()
  const slide = deck.slides[2]!
  slide.elements.push({ kind: 'text', id: 'caption', x: 1, y: 6, w: 5, h: 0.5, text: '图片说明' })
  deck.slides = [slide]
  const original = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const xml = await (
    await JSZip.loadAsync(original, { base64: true })
  )
    .file('ppt/slides/slide1.xml')!
    .async('string')
  const shape = (tag: 'sp' | 'pic', name: string) =>
    [...xml.matchAll(new RegExp(`<p:${tag}\\b[^>]*>[\\s\\S]*?<\\/p:${tag}>`, 'g'))].find(([part]) =>
      part.includes(`name="${name}"`),
    )![0]
  const shapeId = (value: string) => /<p:cNvPr\b[^>]*\bid="(\d+)"/.exec(value)![1]!
  const caption = shape('sp', 'caption')
  const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(caption)!
  const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(caption)!
  const before = {
    left: Number(off[1]) / 12700,
    top: Number(off[2]) / 12700,
    width: Number(ext[1]) / 12700,
    height: Number(ext[2]) / 12700,
  }
  const pictureId = shapeId(shape('pic', 'image'))
  const pictureBefore = await inspectPowerPointPicturePackage(original, pictureId)
  const png = new PNG({ width: 2, height: 1 })
  png.data[0] = 255
  const files = new Map<string, Uint8Array>([['/home/user/replacement.png', PNG.sync.write(png)]])
  ;(globalThis as Record<string, unknown>).createImageBitmap = vi.fn(async (blob: Blob) => {
    const decoded = PNG.sync.read(Buffer.from(await blob.arrayBuffer()))
    return { width: decoded.width, height: decoded.height, close: vi.fn() }
  })
  let hostWrites = 0
  const baseline = {
    baselineId: 'baseline',
    documentId: 'doc',
    scope: { slideIds: ['slide-1'] },
    context: { slideIds: ['slide-1'] },
  }
  const skill = createPresentationExistingPageEditingSkill({
    baseline: {
      snapshot: () => baseline,
      executeTool: async () => ({ output: '{"unchanged":true}' }),
    },
    exportAdapter: {
      exportPresentationPagePackage: async () => ({
        slideId: 'slide-1',
        slideIds: ['slide-1'],
        base64: original,
      }),
    },
    vfs: {
      readBytes: (path: string) => files.get(path),
      writeFile: (path: string, value: Uint8Array) => {
        files.set(path, value)
      },
    },
    documentId: async () => 'doc',
    available: () => true,
    adapter: {
      insert: async () => {
        hostWrites++
      },
    },
  } as unknown as Parameters<typeof createPresentationExistingPageEditingSkill>[0])
  const prepared = await skill.executeTool({
    id: 'composite',
    name: 'prepare_existing_presentation_composite_revision',
    input: {
      baseline_id: 'baseline',
      slide_id: 'slide-1',
      text: {
        shape_id: shapeId(shape('sp', 'title')),
        start: 0,
        before: '研究图文',
        after: '研究图表页',
      },
      geometry: {
        shape_id: shapeId(caption),
        before,
        after: { ...before, top: before.top + 5.76 },
      },
      picture: { shape_id: pictureId, path: '/home/user/replacement.png' },
    },
  })
  expect(prepared.isError, prepared.output).not.toBe(true)
  const result = JSON.parse(prepared.output)
  expect(result.nextTool).toBe('stage_existing_presentation_page_change')
  expect(result.nextInput).toMatchObject({
    baseline_id: 'baseline',
    slide_id: 'slide-1',
    path: result.path,
    picture_shape_id: pictureId,
  })
  expect(hostWrites).toBe(0)
  const revised = Buffer.from(files.get(result.path)!).toString('base64')
  const revisedXml = await (
    await JSZip.loadAsync(revised, { base64: true })
  )
    .file('ppt/slides/slide1.xml')!
    .async('string')
  expect(revisedXml).toContain('<a:t>研究图表页</a:t>')
  expect(revisedXml).toContain(`y="${Math.round((before.top + 5.76) * 12700)}"`)
  expect((await inspectPowerPointPicturePackage(revised, pictureId)).mediaDigest).not.toBe(
    pictureBefore.mediaDigest,
  )
})
