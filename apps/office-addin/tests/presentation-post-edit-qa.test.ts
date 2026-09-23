import { PNG } from 'pngjs'
import { BrowserPresentationImageAdapter } from '../src/skills/powerpoint/browser-presentation-image-adapter'
import { imageReplacementKey } from '../src/skills/powerpoint/presentation-image-replacement-record'
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

function stablePageHost(f: Awaited<ReturnType<typeof fixture>>) {
  let text = 'before'
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'listPresentationPageShapes').mockResolvedValue({
    slideId: 'host',
    shapesTruncated: false,
    shapes: [
      { id: 'shape', name: 'Title', type: 'TextBox', left: 0, top: 0, width: 400, height: 100 },
    ],
  })
  const read = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationPageText')
    .mockImplementation(async (slideId, shapeId) => {
      expect(slideId).toBe('host')
      expect(shapeId).toBe('shape')
      return { slideId, shapeId, text, paragraphs: [text] }
    })
  const edit = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationPageText')
    .mockImplementation(async (slideId, shapeId, next, expected) => {
      expect(slideId).toBe('host')
      expect(shapeId).toBe('shape')
      expect(expected).toBe(text)
      expect(f.page().recheckRequired).toBe(true)
      text = next
      f.setText(next)
    })
  const propose = () =>
    f.runtime.skill.executeTool({
      id: 'stable-edit',
      name: 'edit_presentation_page_text',
      input: { page_id: 'page1', shape_id: 'shape', text: 'after' },
    })
  return { read, edit, propose }
}
it('routes a business-page text edit through its host ID and then fresh QA', async () => {
  const f = await fixture(),
    host = stablePageHost(f)
  const objects = await f.runtime.skill.executeTool({
    id: 'objects',
    name: 'read_presentation_page',
    input: { page_id: 'page1' },
  })
  expect(objects.isError, objects.output).not.toBe(true)
  expect(objects.output).toContain('shape')
  const text = await f.runtime.skill.executeTool({
    id: 'text',
    name: 'read_presentation_page',
    input: { page_id: 'page1', shape_id: 'shape' },
  })
  expect(text.isError, text.output).not.toBe(true)
  expect(text.output).toContain('before')
  const proposal = await host.propose()
  expect(proposal.isError, proposal.output).not.toBe(true)
  // The original index now addresses a different page; the new workflow must not consult it.
  vi.mocked(BrowserPowerPointAdapter.prototype.readSlideText).mockResolvedValue({
    slideId: 'other-page',
    shapeId: 'shape',
    text: 'unrelated',
    paragraphs: ['unrelated'],
  })
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(host.edit).toHaveBeenCalledOnce()
  expect(f.edit).not.toHaveBeenCalled()
  expect(BrowserPowerPointAdapter.prototype.readSlideText).not.toHaveBeenCalled()
  expect(f.page().recheckRequired).toBe(true)
  const fresh = await f.capture()
  expect(fresh.isError).not.toBe(true)
  expect(JSON.parse(fresh.output).page.screenshotDigest).not.toBe(f.digest)
  expect((await f.review(JSON.parse(fresh.output).page.screenshotDigest)).isError).not.toBe(true)
  expect(f.page().visual.status).toBe('pass')
  f.runtime.dispose()
})
it.each(['deleted', 'restored'] as const)(
  'rejects a %s stable page proposal before writing',
  async (change) => {
    const f = await fixture(),
      host = stablePageHost(f)
    const proposal = await host.propose()
    expect(proposal.isError, proposal.output).not.toBe(true)
    const id = f.proposals.pending()!.id
    if (change === 'deleted') host.read.mockRejectedValue(new Error('office_read_failed'))
    else {
      const restored = await f.runtime.skill.executeTool({
        id: 'restore-again',
        name: 'restore_presentation_project',
        input: { project_id: 'project' },
      })
      expect(restored.isError).not.toBe(true)
    }
    await expect(f.proposals.confirm(id)).rejects.toThrow()
    expect(host.edit).not.toHaveBeenCalled()
    expect(f.edit).not.toHaveBeenCalled()
    expect(f.page().recheckRequired).toBeUndefined()
    f.runtime.dispose()
  },
)

function stableGeometryHost() {
  let geometry = { left: 10, top: 20, width: 300, height: 100 }
  const read = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationPageGeometry')
    .mockImplementation(async (slideId, shapeId) => {
      expect([slideId, shapeId]).toEqual(['host', 'shape'])
      return { slideId, shapeId, geometry: { ...geometry } }
    })
  const edit = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationPageGeometry')
    .mockImplementation(async (slideId, shapeId, next, expected) => {
      expect([slideId, shapeId]).toEqual(['host', 'shape'])
      expect(expected).toEqual(geometry)
      geometry = { ...next }
    })
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectPresentationPage').mockImplementation(
    async () => ({
      slideId: 'host',
      slideWidth: 960,
      slideHeight: 540,
      shapes: [{ id: 'shape', name: 'Photo', type: 'Image', ...geometry }],
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: afterPng },
    }),
  )
  return {
    read,
    edit,
    moveManually: () => {
      geometry.left += 20
    },
  }
}
it('moves and resizes a stable page object through confirmation and new QA without index calls', async () => {
  const f = await fixture(),
    host = stableGeometryHost()
  const before = await f.runtime.skill.executeTool({
    id: 'geometry',
    name: 'read_presentation_page_geometry',
    input: { page_id: 'page1', shape_id: 'shape' },
  })
  expect(before.isError, before.output).not.toBe(true)
  expect(JSON.parse(before.output)).toMatchObject({
    unit: 'pt',
    geometry: { left: 10, top: 20, width: 300, height: 100 },
  })
  const next = { left: 100, top: 60, width: 500, height: 200 }
  const proposal = await f.runtime.skill.executeTool({
    id: 'layout',
    name: 'edit_presentation_page_geometry',
    input: { page_id: 'page1', shape_id: 'shape', geometry: next },
  })
  expect(proposal.isError, proposal.output).not.toBe(true)
  expect(f.proposals.pending()?.after).toEqual(next)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(host.edit).toHaveBeenCalledOnce()
  expect(f.edit).not.toHaveBeenCalled()
  expect(BrowserPowerPointAdapter.prototype.readSlideText).not.toHaveBeenCalled()
  expect(f.page().recheckRequired).toBe(true)
  const after = await f.runtime.skill.executeTool({
    id: 'geometry-after',
    name: 'read_presentation_page_geometry',
    input: { page_id: 'page1', shape_id: 'shape' },
  })
  expect(JSON.parse(after.output).geometry).toEqual(next)
  const captured = await f.capture()
  expect(captured.isError, captured.output).not.toBe(true)
  expect(JSON.parse(captured.output).page.structure.shapeCount).toBe(1)
  expect((await f.review(JSON.parse(captured.output).page.screenshotDigest)).isError).not.toBe(true)
  expect(f.page().visual.status).toBe('pass')
  f.runtime.dispose()
})
it('preserves a manual geometry change made after proposing layout adjustments', async () => {
  const f = await fixture(),
    host = stableGeometryHost()
  const proposal = await f.runtime.skill.executeTool({
    id: 'layout',
    name: 'edit_presentation_page_geometry',
    input: {
      page_id: 'page1',
      shape_id: 'shape',
      geometry: { left: 100, top: 60, width: 500, height: 200 },
    },
  })
  expect(proposal.isError, proposal.output).not.toBe(true)
  const id = f.proposals.pending()!.id
  host.moveManually()
  await expect(f.proposals.confirm(id)).rejects.toThrow('proposal_stale')
  expect(host.edit).not.toHaveBeenCalled()
  expect(f.page().recheckRequired).toBeUndefined()
  f.runtime.dispose()
})

it.each([false, true])(
  'persists image replacement through runtime and blocks replay (interrupted=%s)',
  async (interrupted) => {
    const f = await fixture()
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 1, height: 1, close() {} })),
    )
    const png = new PNG({ width: 1, height: 1 })
    png.data.fill(120)
    f.runtime.vfs.writeFile('/home/user/replacement.png', PNG.sync.write(png))
    const snapshot = {
      slideId: 'host',
      shapeId: 'old',
      geometry: { left: 1, top: 2, width: 100, height: 50 },
      rotation: 0,
      name: 'Picture',
      altTextTitle: '',
      altTextDescription: '',
      zOrderPosition: 0,
      shapeIds: ['old'],
      pictureFingerprint: 'a'.repeat(64),
      mediaDigest: 'b'.repeat(64),
    }
    vi.spyOn(BrowserPresentationImageAdapter.prototype, 'inspect').mockResolvedValue(snapshot)
    const key = await imageReplacementKey('project', 'request', 'page1', 'old')
    const native = vi
      .spyOn(BrowserPresentationImageAdapter.prototype, 'replace')
      .mockImplementation(async (_slide, _shape, _base64, _expected, onInserted) => {
        expect(f.page().recheckRequired).toBe(true)
        expect(f.binding.readImageReplacement(key)?.state).toBe('pending')
        await onInserted('new')
        expect(f.binding.readImageReplacement(key)?.newShapeId).toBe('new')
        if (interrupted) throw new Error('office_write_uncertain')
        return { shapeId: 'new' }
      })
    const call = {
      id: 'replace-image',
      name: 'replace_presentation_page_image',
      input: { page_id: 'page1', shape_id: 'old', path: '/home/user/replacement.png' },
    }
    expect((await f.runtime.skill.executeTool(call)).isError).not.toBe(true)
    const confirmation = f.proposals.confirm(f.proposals.pending()!.id)
    if (interrupted) await expect(confirmation).rejects.toThrow('office_write_uncertain')
    else await confirmation
    expect(f.binding.readImageReplacement(key)).toMatchObject({
      state: interrupted ? 'pending' : 'complete',
      newShapeId: 'new',
      hostSlideId: 'host',
    })
    const status = await f.runtime.skill.executeTool({
      id: 'status',
      name: 'read_presentation_image_replacement',
      input: { page_id: 'page1', shape_id: 'old' },
    })
    expect(status.isError).not.toBe(true)
    expect(status.output).toContain(interrupted ? 'pending' : 'complete')
    await f.runtime.skill.executeTool(call)
    expect(native).toHaveBeenCalledOnce()
    expect(f.proposals.pending()).toBeUndefined()
    f.runtime.dispose()
  },
)
