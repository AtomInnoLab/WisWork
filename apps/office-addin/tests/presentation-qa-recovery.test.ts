import { createHash } from 'node:crypto'
import { expect, it } from 'vitest'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { createPresentationQaSkill } from '../src/skills/powerpoint/presentation-qa'
import { InMemoryVfs } from '../src/skills/shared/vfs'

it('persists captured evidence and Agent review through settings, requiring recapture after restart', async () => {
  const values = new Map<string, string>()
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save: async () => {},
    location: () => 'file://qa-recovery.pptx',
  }
  let binding = createPresentationDocumentBinding(settings, () => 'qa-document')
  const documentId = await binding.documentId()
  const artifact = {
    documentId,
    projectId: 'project',
    requestId: 'request',
    pptxBase64: 'UEsDBAAAAAA=',
    slideCount: 1,
    pages: [{ id: 'page1', title: '测试页', sourceSlideId: '256#' }],
  }
  await binding.writeReceipt('project/request', {
    state: 'complete',
    documentId,
    slideIds: ['host-page'],
    checkpoint: {
      version: 1,
      artifactDigest: createHash('sha256').update(artifact.pptxBase64).digest('hex'),
      sourceSlideIds: ['256#'],
      baselineSlideIds: [],
      completed: [{ sourceSlideId: '256#', slideId: 'host-page' }],
    },
  })
  const png =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
  const vfs = new InMemoryVfs()
  const setup = () =>
    createPresentationQaSkill({
      ...binding,
      vfs,
      available: () => true,
      artifact: () => artifact,
      inspectPage: async (id) => ({
        slideId: id,
        slideWidth: 960,
        slideHeight: 540,
        shapes: [],
        shapesTruncated: false,
        overflows: [],
        overlaps: [],
        overlapsTruncated: false,
        screenshot: { mime: 'image/png', base64: png },
      }),
    })
  let qa = setup()
  const capture = await qa.executeTool({
    id: 'capture',
    name: 'capture_presentation_page_qa',
    input: { page_id: 'page1' },
  })
  expect(capture.isError).not.toBe(true)
  expect(capture.modelContent).toEqual([
    { type: 'image', image: { mime: 'image/png', base64: png } },
  ])
  const page = JSON.parse(capture.output).page
  expect(page.visual.status).toBe('needs_review')
  const review = {
    id: 'review',
    name: 'record_presentation_page_review',
    input: {
      page_id: 'page1',
      screenshot_digest: page.screenshotDigest,
      outcome: 'pass',
      notes: '测试中的 Agent 复核说明',
    },
  }
  expect((await qa.executeTool(review)).isError).not.toBe(true)
  expect(binding.readQa('project/request')?.pages[0]?.visual).toMatchObject({
    status: 'pass',
    reviewer: 'agent',
  })
  qa.clear()
  vfs.clear()
  binding = createPresentationDocumentBinding(settings, () => 'unused')
  qa = setup()
  const history = JSON.parse(
    (await qa.executeTool({ id: 'history', name: 'read_presentation_qa', input: {} })).output,
  )
  expect(history.needs_recapture).toBe(true)
  expect(history.checks).toMatchObject({ sources: 'not_verified', saveReopen: 'not_run' })
  expect(history.record.pages[0].visual.status).toBe('pass')
  expect((await qa.executeTool(review)).output).toBe('presentation_qa_capture_required')
})
