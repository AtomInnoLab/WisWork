import { expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { PNG } from 'pngjs'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { createStructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationExistingPageEditingSkill } from '../src/skills/powerpoint/presentation-existing-page-editing'
import { inspectPowerPointPicturePackage } from '../src/skills/powerpoint/powerpoint-package'
import { InMemoryVfs } from '../src/skills/shared/vfs'
import { readBoundedImage } from '../src/skills/shared/import-media'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'

it('prepares a native picture revision for the confirmed page transaction without writing the host', async () => {
  const deck = benchmarkDeck()
  deck.slides = [deck.slides[2]!]
  const source = Buffer.from((await compilePresentationDeck(deck)).bytes).toString('base64')
  const zip = await JSZip.loadAsync(source, { base64: true })
  const xml = await zip.file('ppt/slides/slide1.xml')!.async('string')
  const shapeId = xml.match(/<p:pic\b[^]*?<p:cNvPr\b[^>]*\bid="(\d+)"/)![1]!
  const original = await inspectPowerPointPicturePackage(source, shapeId)
  const vfs = new InMemoryVfs()
  const png = new PNG({ width: 2, height: 1 })
  png.data[0] = 255
  vfs.writeFile('/home/user/new.png', PNG.sync.write(png))
  vi.stubGlobal('createImageBitmap', async () => ({ width: 2, height: 1, close() {} }))
  expect((await readBoundedImage(vfs, '/home/user/new.png')).mime).toBe('image/png')
  let hostWrites = 0
  const baseline = {
    baselineId: 'baseline', documentId: 'doc', contentDigest: 'a'.repeat(64),
    scope: { kind: 'current', slideIds: ['old'] },
    context: { slideIds: ['old'], selectedSlideIds: ['old'], selectedShapeIds: [] },
    pages: [{ slideId: 'old', shapes: [] }],
  }
  const options = {
    baseline: { snapshot: () => structuredClone(baseline), executeTool: async () => ({ output: '{"unchanged":true}', mutated: false }) },
    adapter: { stage: async () => { hostWrites++ } },
    inspectPage: async () => ({ slideId: 'old', shapesTruncated: false }),
    exportAdapter: { exportPresentationPagePackage: async () => ({ slideId: 'old', slideIds: ['old'], base64: source }) },
    vfs, request: async () => new Response('{}'), proposals: createStructuredProposalController(),
    documentId: async () => 'doc', readExistingPageChange: () => undefined,
    writeExistingPageChange: async () => { hostWrites++ }, available: () => true,
  } as unknown as Parameters<typeof createPresentationExistingPageEditingSkill>[0]
  const skill = createPresentationExistingPageEditingSkill(options)
  const result = await skill.executeTool({ id: 'prepare', name: 'prepare_existing_presentation_image_revision', input: {
    baseline_id: 'baseline', slide_id: 'old', shape_id: shapeId, path: '/home/user/new.png',
  } })
  expect(result.isError, result.output).not.toBe(true)
  const prepared = JSON.parse(result.output) as { path: string; beforeDigest: string; afterDigest: string; nextTool: string }
  expect(prepared.path).toMatch(/^\/home\/user\/presentation-image-revision-.*\.pptx$/)
  expect(prepared.beforeDigest).not.toBe(prepared.afterDigest)
  expect(prepared.nextTool).toBe('stage_existing_presentation_page_change')
  const changed = Buffer.from(vfs.readBytes(prepared.path, { maxBytes: 8 * 1024 * 1024 })).toString('base64')
  expect((await inspectPowerPointPicturePackage(changed, shapeId)).mediaDigest).not.toBe(original.mediaDigest)
  const staged = await skill.executeTool({ id: 'stage', name: 'stage_existing_presentation_page_change', input: {
    baseline_id: 'baseline', slide_id: 'old', path: prepared.path,
  } })
  expect(staged.isError, staged.output).not.toBe(true)
  expect(JSON.parse(staged.output)).toMatchObject({ status: 'awaiting_confirmation' })
  expect(hostWrites).toBe(0)
  vi.unstubAllGlobals()
})
