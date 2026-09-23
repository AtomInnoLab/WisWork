import { createHash } from 'node:crypto'
import { afterEach, expect, it, vi } from 'vitest'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
import type { StructuredProposalController } from '../src/agent/proposal-controller'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'

const beforePng =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
const afterPng =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})
async function fixture() {
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
  const values = new Map<string, string>()
  const save = vi.fn(async () => {})
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save,
    location: () => 'file://edit-qa.pptx',
  }
  const binding = createPresentationDocumentBinding(settings, () => 'doc')
  const documentId = await binding.documentId()
  const pptxBase64 = 'UEsDBAAAAAA='
  await binding.writeReceipt('project/request', {
    state: 'complete',
    documentId,
    slideIds: ['host'],
    checkpoint: {
      version: 1,
      artifactDigest: createHash('sha256').update(pptxBase64).digest('hex'),
      sourceSlideIds: ['256#'],
      baselineSlideIds: [],
      completed: [{ sourceSlideId: '256#', slideId: 'host' }],
    },
  })
  let text = 'before'
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'verifySlides').mockResolvedValue({
    slideWidth: 960,
    slideHeight: 540,
    slides: [],
  })
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'readSlideText').mockImplementation(async () => ({
    slideId: 'host',
    shapeId: 'shape',
    text,
    paragraphs: [text],
  }))
  const edit = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editSlideText')
    .mockImplementation(async (_index, _id, next) => {
      text = next
    })
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectPresentationPage').mockImplementation(
    async () => ({
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes: [],
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: text === 'before' ? beforePng : afterPng },
    }),
  )
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      ...binding,
      available: () => true,
      request: async () =>
        new Response(
          JSON.stringify({
            projectId: 'project',
            requestId: 'request',
            status: 'compiled',
            pptxBase64,
            report: { deckId: 'project', slideCount: 1 },
            pages: [{ id: 'page1', title: 'Page', sourceSlideId: '256#' }],
          }),
        ),
    },
  })
  const restored = await runtime.skill.executeTool({
    id: 'restore',
    name: 'restore_presentation_project',
    input: { project_id: 'project' },
  })
  expect(restored.isError).not.toBe(true)
  const capture = () =>
    runtime.skill.executeTool({
      id: 'capture',
      name: 'capture_presentation_page_qa',
      input: { page_id: 'page1' },
    })
  const review = (digest: string) =>
    runtime.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: 'page1',
        screenshot_digest: digest,
        outcome: 'pass',
        notes: '测试Agent复核',
      },
    })
  const first = await capture()
  expect(first.isError).not.toBe(true)
  const digest = JSON.parse(first.output).page.screenshotDigest
  expect((await review(digest)).isError).not.toBe(true)
  // Keep live evidence too: a pending write must invalidate even unreviewed captures.
  expect((await capture()).isError).not.toBe(true)
  const proposals = runtime.proposals as StructuredProposalController
  const propose = async () => {
    const result = await runtime.skill.executeTool({
      id: 'edit',
      name: 'edit_slide_text',
      input: { slide_index: 1, shape_id: 'shape', text: 'after' },
    })
    expect(result.isError).not.toBe(true)
    return proposals.pending()!.id
  }
  return {
    binding,
    setText: (value: string) => {
      text = value
    },
    save,
    edit,
    runtime,
    proposals,
    propose,
    capture,
    review,
    digest,
    page: () => binding.readQa('project/request')!.pages[0]!,
  }
}
it('invalidates before a confirmed edit, blocks concurrent QA, and requires fresh review after it', async () => {
  const f = await fixture()
  const rejected = await f.propose()
  expect(rejected).toBeTruthy()
  f.proposals.reject()
  expect(f.page().recheckRequired).toBeUndefined()
  const id = await f.propose()
  let finish!: () => void
  f.edit.mockImplementationOnce(async () => {
    expect(f.page().recheckRequired).toBe(true)
    await new Promise<void>((resolve) => {
      finish = resolve
    })
  })
  const pending = f.proposals.confirm(id)
  await vi.waitFor(() => expect(f.edit).toHaveBeenCalledOnce())
  expect((await f.capture()).output).toBe('presentation_qa_busy')
  // Simulate the queued Office write and then reconcile its text readback.
  f.setText('after')
  finish()
  await pending
  expect(f.page().recheckRequired).toBe(true)
  expect((await f.review(f.digest)).output).toBe('presentation_qa_capture_required')
  const fresh = await f.capture()
  expect(fresh.isError).not.toBe(true)
  expect(JSON.parse(fresh.output).page.screenshotDigest).not.toBe(f.digest)
  expect(f.page().recheckRequired).toBeUndefined()
  expect((await f.review(JSON.parse(fresh.output).page.screenshotDigest)).isError).not.toBe(true)
  expect(f.page().visual.status).toBe('pass')
  f.runtime.dispose()
})
it('prevents document writes when QA invalidation cannot be saved', async () => {
  const f = await fixture(),
    id = await f.propose()
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(f.proposals.confirm(id)).rejects.toThrow('save_failed')
  expect(f.edit).not.toHaveBeenCalled()
  expect(f.page().recheckRequired).toBeUndefined()
  expect((await f.review(f.digest)).output).toBe('presentation_qa_capture_required')
  expect((await f.capture()).isError).not.toBe(true)
  f.runtime.dispose()
})
it('retains recheck status after uncertain host write failure and releases the QA lock', async () => {
  const f = await fixture(),
    id = await f.propose()
  f.edit.mockRejectedValueOnce(new Error('office_state_uncertain'))
  await expect(f.proposals.confirm(id)).rejects.toThrow('office_state_uncertain')
  expect(f.page().recheckRequired).toBe(true)
  expect((await f.review(f.digest)).output).toBe('presentation_qa_capture_required')
  expect((await f.capture()).isError).not.toBe(true)
  f.runtime.dispose()
})
