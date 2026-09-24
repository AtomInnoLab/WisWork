import { PNG } from 'pngjs'
import { BrowserPresentationImageAdapter } from '../src/skills/powerpoint/browser-presentation-image-adapter'
import { imageReplacementKey } from '../src/skills/powerpoint/presentation-image-replacement-record'
import { createHash } from 'node:crypto'
import { afterEach, expect, it, vi } from 'vitest'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
import type { StructuredProposalController } from '../src/agent/proposal-controller'
import {
  BrowserPowerPointAdapter,
  type PowerPointMasterState,
  type PowerPointMasterOperation,
} from '../src/skills/powerpoint/browser-powerpoint-adapter'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'

const beforePng =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
const afterPng =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})
async function fixture(withSecondPage = false) {
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
    slideIds: withSecondPage ? ['host', 'host-2'] : ['host'],
    checkpoint: {
      version: 1,
      artifactDigest: createHash('sha256').update(pptxBase64).digest('hex'),
      sourceSlideIds: withSecondPage ? ['256#', '257#'] : ['256#'],
      baselineSlideIds: [],
      completed: [
        { sourceSlideId: '256#', slideId: 'host' },
        ...(withSecondPage ? [{ sourceSlideId: '257#', slideId: 'host-2' }] : []),
      ],
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
    async (slideId) => ({
      slideId,
      slideWidth: 960,
      slideHeight: 540,
      shapes: [],
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: {
        mime: 'image/png',
        base64: slideId === 'host-2' || text === 'before' ? beforePng : afterPng,
      },
    }),
  )
  const createRuntime = () =>
    createOfficeHostRuntime('powerpoint', {
      presentation: {
        ...createPresentationDocumentBinding(settings, () => 'doc'),
        available: () => true,
        request: async () =>
          new Response(
            JSON.stringify({
              projectId: 'project',
              requestId: 'request',
              status: 'compiled',
              pptxBase64,
              report: { deckId: 'project', slideCount: withSecondPage ? 2 : 1 },
              pages: [
                { id: 'page1', title: 'Page', sourceSlideId: '256#' },
                ...(withSecondPage
                  ? [{ id: 'page2', title: 'Page 2', sourceSlideId: '257#' }]
                  : []),
              ],
            }),
          ),
      },
    })
  const runtime = createRuntime()
  const restored = await runtime.skill.executeTool({
    id: 'restore',
    name: 'restore_presentation_project',
    input: { project_id: 'project' },
  })
  expect(restored.isError).not.toBe(true)
  const capture = (pageId = 'page1') =>
    runtime.skill.executeTool({
      id: 'capture',
      name: 'capture_presentation_page_qa',
      input: { page_id: pageId },
    })
  const review = (digest: string, pageId = 'page1') =>
    runtime.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: pageId,
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
  let secondDigest: string | undefined
  if (withSecondPage) {
    const second = await capture('page2')
    expect(second.isError, second.output).not.toBe(true)
    secondDigest = JSON.parse(second.output).page.screenshotDigest
    expect((await review(secondDigest!, 'page2')).isError).not.toBe(true)
    expect((await capture('page2')).isError).not.toBe(true)
  }
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
    createRuntime,
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
    secondDigest,
    secondPage: () =>
      binding.readQa('project/request')!.pages.find((page) => page.pageId === 'page2')!,
    page: () => binding.readQa('project/request')!.pages.find((page) => page.pageId === 'page1')!,
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
  const f = await fixture(true),
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
  expect(f.secondPage().recheckRequired).toBeUndefined()
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
    const f = await fixture(true)
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
    expect(f.secondPage().recheckRequired).toBeUndefined()
    f.runtime.dispose()
  },
)

it.each(['ready_to_finish', 'already_applied', 'completion_save_failed'] as const)(
  'recovers an image replacement after reopening without the original VFS asset (%s)',
  async (scenario) => {
    const f = await fixture(true),
      key = await imageReplacementKey('project', 'request', 'page1', 'old')
    const baseline = {
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
    const record = {
      version: 1 as const,
      documentId: await f.binding.documentId(),
      projectId: 'project',
      requestId: 'request',
      pageId: 'page1',
      hostSlideId: 'host',
      oldShapeId: 'old',
      assetDigest: 'c'.repeat(64),
      state: 'pending' as const,
      baseline,
    }
    await f.binding.writeImageReplacement(key, record)
    await f.binding.writeImageReplacement(key, { ...record, newShapeId: 'new' })
    f.runtime.dispose()
    vi.stubGlobal('createImageBitmap', undefined)
    let status: 'ready_to_finish' | 'already_applied' =
      scenario === 'already_applied' ? 'already_applied' : 'ready_to_finish'
    const inspect = vi
      .spyOn(BrowserPresentationImageAdapter.prototype, 'inspectRecovery')
      .mockImplementation(async () => ({ status }))
    const insert = vi.spyOn(BrowserPresentationImageAdapter.prototype, 'replace')
    let failOnce = scenario === 'completion_save_failed'
    const finish = vi
      .spyOn(BrowserPresentationImageAdapter.prototype, 'finishRecovery')
      .mockImplementation(async () => {
        expect(f.page().recheckRequired).toBe(true)
        status = 'already_applied'
        if (failOnce) {
          f.save.mockRejectedValueOnce(new Error('save_failed'))
          failOnce = false
        }
        return { shapeId: 'new' }
      })
    const runtime = f.createRuntime(),
      proposals = runtime.proposals as StructuredProposalController
    expect(
      (
        await runtime.skill.executeTool({
          id: 'restore',
          name: 'restore_presentation_project',
          input: { project_id: 'project' },
        })
      ).isError,
    ).not.toBe(true)
    const input = { page_id: 'page1', shape_id: 'old' }
    const live = await runtime.skill.executeTool({
      id: 'inspect',
      name: 'inspect_presentation_image_replacement',
      input,
    })
    expect(live.isError, live.output).not.toBe(true)
    expect(live.output).toContain(status)
    expect(f.page().recheckRequired).toBeUndefined()
    const propose = async () => {
      const result = await runtime.skill.executeTool({
        id: 'resume',
        name: 'resume_presentation_image_replacement',
        input,
      })
      expect(result.isError, result.output).not.toBe(true)
      return proposals.pending()!.id
    }
    const confirmed = proposals.confirm(await propose())
    if (scenario === 'completion_save_failed') {
      await expect(confirmed).rejects.toThrow('save_failed')
      expect(f.binding.readImageReplacement(key)?.state).toBe('pending')
      await proposals.confirm(await propose())
    } else await confirmed
    expect(f.binding.readImageReplacement(key)).toMatchObject({
      state: 'complete',
      newShapeId: 'new',
      baseline,
    })
    expect(inspect).toHaveBeenCalled()
    expect(finish).toHaveBeenCalledTimes(scenario === 'completion_save_failed' ? 2 : 1)
    expect(insert).not.toHaveBeenCalled()
    expect(f.page().recheckRequired).toBe(true)
    expect(f.secondPage().recheckRequired).toBeUndefined()
    runtime.dispose()
  },
)

it('keeps another page review and live capture valid after a stable page text edit', async () => {
  const f = await fixture(true),
    host = stablePageHost(f)
  const previous = structuredClone(f.secondPage())
  expect((await host.propose()).isError).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  expect(f.page().recheckRequired).toBe(true)
  expect(f.secondPage()).toEqual(previous)
  expect((await f.review(f.digest)).output).toBe('presentation_qa_capture_required')
  const review = await f.review(f.secondDigest!, 'page2')
  expect(review.isError, review.output).not.toBe(true)
  f.runtime.dispose()
})
it('still invalidates every page for an index-based operation with uncertain impact', async () => {
  const f = await fixture(true)
  await f.proposals.confirm(await f.propose())
  expect(f.page().recheckRequired).toBe(true)
  expect(f.secondPage().recheckRequired).toBe(true)
  expect((await f.review(f.secondDigest!, 'page2')).output).toBe('presentation_qa_capture_required')
  f.runtime.dispose()
})
it('does not trust a stable tool label on a general script proposal', async () => {
  const f = await fixture(true)
  const proposal = f.proposals.propose({
    operation: 'execute_office_js',
    toolName: 'edit_presentation_page_text',
    title: 'Script',
    preview: { qaScope: { basis: 'native_master_layout', hostSlideIds: [] } },
    impact: { host: 'powerpoint', targets: ['host'], count: 1 },
    fingerprint: 'test',
    validate: () => true,
    execute: () => {},
  })
  await f.proposals.confirm(proposal.id)
  expect(f.page().recheckRequired).toBe(true)
  expect(f.secondPage().recheckRequired).toBe(true)
  f.runtime.dispose()
})

it.each([{ scope: ['host'] }, { scope: [] }])(
  'uses the native master dependency scope $scope for saved and live QA',
  async ({ scope }) => {
    const f = await fixture(true)
    const previous = structuredClone(f.secondPage())
    const beforeSave = f.save.mock.calls.length
    const proposal = f.proposals.propose({
      operation: 'edit_slide_master',
      toolName: 'edit_slide_master',
      title: 'Native master edit',
      preview: { qaScope: { basis: 'native_master_layout', hostSlideIds: scope } },
      impact: { host: 'powerpoint', targets: ['master:master1'], count: 1 },
      fingerprint: 'test',
      validate: () => true,
      execute: () => {
        expect(f.page().recheckRequired).toBe(scope.length ? true : undefined)
      },
    })
    await f.proposals.confirm(proposal.id)
    expect(f.secondPage()).toEqual(previous)
    expect((await f.review(f.secondDigest!, 'page2')).isError).not.toBe(true)
    if (!scope.length) {
      expect(f.save.mock.calls.length).toBe(beforeSave + 1) // only the unrelated review was saved
      expect((await f.review(f.digest)).isError).not.toBe(true)
    } else expect((await f.review(f.digest)).output).toBe('presentation_qa_capture_required')
    f.runtime.dispose()
  },
)

it.each([
  { basis: 'document' },
  { basis: 'native_master_layout', hostSlideIds: ['host', 'host'] },
  { basis: 'native_master_layout', hostSlideIds: ['host'], partial: true },
])('keeps malformed or unknown native scopes document-wide (%j)', async (qaScope) => {
  const f = await fixture(true)
  const proposal = f.proposals.propose({
    operation: 'edit_slide_master',
    toolName: 'edit_slide_master',
    title: 'Native master edit',
    preview: { qaScope },
    impact: { host: 'powerpoint', targets: ['master:master1'], count: 1 },
    fingerprint: 'test',
    validate: () => true,
    execute: () => {},
  })
  await f.proposals.confirm(proposal.id)
  expect(f.page().recheckRequired).toBe(true)
  expect(f.secondPage().recheckRequired).toBe(true)
  f.runtime.dispose()
})

function nativeStyleHost(dependencies: { slideId: string; masterId: string; layoutId: string }[]) {
  let slides = structuredClone(dependencies)
  const masterState: PowerPointMasterState = {
    masters: ['master1', 'master2', 'unused'].map((id) => ({
      id,
      name: id,
      background: { type: 'Solid', color: '#FFFFFF', transparency: 0 },
      themeColors: { Light1: '#FFFFFF', Dark1: '#000000' },
      layouts: ['layout1', 'layout2'].map((id) => ({
        id,
        name: id,
        isMasterBackgroundFollowed: true,
        areBackgroundGraphicsHidden: false,
        background: { type: 'Solid' },
      })),
    })),
  }
  const inspect = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'inspectStyleDependencies')
    .mockImplementation(async () => ({ slides: structuredClone(slides) }))
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectSlideMasters').mockImplementation(async () =>
    structuredClone(masterState),
  )
  const execute = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'executeMasterOperations')
    .mockImplementation(async (operations) => {
      for (const operation of operations) {
        const master = masterState.masters.find((master) => master.id === operation.master_id)!
        if (operation.op === 'set_master_theme_color')
          master.themeColors[operation.theme_color] = operation.color
        else if (operation.op === 'set_layout_background_following') {
          const layout = master.layouts.find((layout) => layout.id === operation.layout_id)!
          layout.isMasterBackgroundFollowed = operation.follow_master
          layout.areBackgroundGraphicsHidden = !operation.show_master_graphics
        } else if (operation.fill.type === 'solid')
          master.background = {
            type: 'Solid',
            color: operation.fill.color,
            transparency: operation.fill.transparency,
          }
      }
    })
  return {
    inspect,
    execute,
    change: (next: typeof slides) => {
      slides = structuredClone(next)
    },
  }
}
const stylePages = [
  { slideId: 'host', masterId: 'master1', layoutId: 'layout1' },
  { slideId: 'host-2', masterId: 'master1', layoutId: 'layout2' },
]
const themeOperation: PowerPointMasterOperation = {
  op: 'set_master_theme_color',
  master_id: 'master1',
  theme_color: 'Dark1',
  color: '#333333',
}
async function proposeStyle(
  f: Awaited<ReturnType<typeof fixture>>,
  operation: PowerPointMasterOperation,
) {
  const result = await f.runtime.skill.executeTool({
    id: 'style',
    name: 'edit_slide_master',
    input: { program: { version: 2, operations: [operation] } },
  })
  expect(result.isError, result.output).not.toBe(true)
  return f.proposals.pending()!
}
it.each(['theme', 'layout', 'different_master', 'unused_master', 'unknown'] as const)(
  'connects real native-style proposals to durable and live QA (%s)',
  async (scenario) => {
    const f = await fixture(true)
    const pages = structuredClone(stylePages)
    if (scenario === 'different_master') pages[1]!.masterId = 'master2'
    const native = nativeStyleHost(pages)
    if (scenario === 'unknown')
      native.inspect.mockRejectedValue(new Error('office_api_unsupported'))
    const operation: PowerPointMasterOperation =
      scenario === 'layout'
        ? {
            op: 'set_layout_background_following',
            master_id: 'master1',
            layout_id: 'layout1',
            follow_master: false,
            show_master_graphics: false,
          }
        : { ...themeOperation, master_id: scenario === 'unused_master' ? 'unused' : 'master1' }
    const before = [structuredClone(f.page()), structuredClone(f.secondPage())]
    const proposal = await proposeStyle(f, operation)
    const affected =
      scenario === 'unused_master'
        ? []
        : scenario === 'theme' || scenario === 'unknown'
          ? ['host', 'host-2']
          : ['host']
    expect(proposal.preview.qaScope).toEqual(
      scenario === 'unknown'
        ? { basis: 'document' }
        : { basis: 'native_master_layout', hostSlideIds: affected },
    )
    await f.proposals.confirm(proposal.id)
    expect(native.execute).toHaveBeenCalledOnce()
    for (const [index, entry] of [f.page(), f.secondPage()].entries()) {
      if (affected.includes(entry.hostSlideId)) expect(entry.recheckRequired).toBe(true)
      else expect(entry).toEqual(before[index])
    }
    expect((await f.review(f.digest)).isError).toBe(affected.includes('host') ? true : undefined)
    expect((await f.review(f.secondDigest!, 'page2')).isError).toBe(
      affected.includes('host-2') ? true : undefined,
    )
    f.runtime.dispose()
    const reopened = f.createRuntime()
    await reopened.skill.executeTool({
      id: 'restore',
      name: 'restore_presentation_project',
      input: { project_id: 'project' },
    })
    expect(
      reopened
        .qa!.read()!
        .pages.filter((page) => page.recheckRequired)
        .map((page) => page.hostSlideId)
        .sort(),
    ).toEqual(affected)
    reopened.dispose()
  },
)
it.each(['before_confirm', 'during_save', 'save_failed'] as const)(
  'blocks native shared-style writes when the dependency checkpoint is unsafe (%s)',
  async (scenario) => {
    const f = await fixture(true)
    const native = nativeStyleHost(stylePages)
    const proposal = await proposeStyle(f, themeOperation)
    const next = structuredClone(stylePages)
    next[1]!.masterId = 'master2'
    if (scenario === 'before_confirm') native.change(next)
    else if (scenario === 'during_save')
      f.save.mockImplementationOnce(async () => {
        native.change(next)
      })
    else f.save.mockRejectedValueOnce(new Error('save_failed'))
    await expect(f.proposals.confirm(proposal.id)).rejects.toThrow(
      scenario === 'save_failed' ? 'save_failed' : 'proposal_stale',
    )
    expect(native.execute).not.toHaveBeenCalled()
    expect(f.page().recheckRequired).toBe(scenario === 'during_save' ? true : undefined)
    expect(f.secondPage().recheckRequired).toBe(scenario === 'during_save' ? true : undefined)
    expect((await f.capture()).isError).not.toBe(true)
    f.runtime.dispose()
  },
)

it('restores text differences after reopening and undoes from the workbench with scoped QA', async () => {
  const f = await fixture(true),
    host = stablePageHost(f)
  const unrelated = structuredClone(f.secondPage())
  expect((await host.propose()).isError).not.toBe(true)
  await f.proposals.confirm(f.proposals.pending()!.id)
  await f.runtime.changes!.refresh()
  expect(
    f.runtime.changes!.snapshot().entries.find((entry) => entry.kind === 'text'),
  ).toMatchObject({ before: 'before', after: 'after', state: 'applied', actions: ['undo'] })
  f.runtime.dispose()
  const reopened = f.createRuntime()
  try {
    await reopened.skill.executeTool({
      id: 'restore',
      name: 'restore_presentation_project',
      input: { project_id: 'project' },
    })
    await reopened.changes!.refresh()
    const entry = reopened.changes!.snapshot().entries.find((entry) => entry.kind === 'text')!
    await reopened.changes!.run(entry.id, 'undo')
    expect(host.edit).toHaveBeenCalledOnce()
    expect(reopened.proposals.pending()?.operation).toBe('undo_presentation_text_change')
    await reopened.proposals.confirm(reopened.proposals.pending()!.id)
    await reopened.changes!.refresh()
    expect(host.edit).toHaveBeenCalledTimes(2)
    expect(host.edit.mock.calls[1]!.slice(0, 4)).toEqual(['host', 'shape', 'before', 'after'])
    expect(
      reopened.changes!.snapshot().entries.find((value) => value.kind === 'text'),
    ).toMatchObject({ state: 'undone', actions: [] })
    expect(f.page().recheckRequired).toBe(true)
    expect(f.secondPage()).toEqual(unrelated)
  } finally {
    reopened.dispose()
  }
})

it('recovers a text completion-save failure from the workbench without repeating the native write', async () => {
  const f = await fixture(true),
    host = stablePageHost(f)
  const write = host.edit.getMockImplementation()!
  host.edit.mockImplementationOnce(async (...args) => {
    await write(...args)
    f.save.mockRejectedValueOnce(new Error('completion_save_failed'))
  })
  try {
    await host.propose()
    await expect(f.proposals.confirm(f.proposals.pending()!.id)).rejects.toThrow(
      'completion_save_failed',
    )
    await f.runtime.changes!.refresh()
    const entry = f.runtime.changes!.snapshot().entries.find((entry) => entry.kind === 'text')!
    expect(entry.state).toBe('pending')
    await f.runtime.changes!.run(entry.id, 'inspect')
    expect(host.edit).toHaveBeenCalledOnce()
    expect(f.proposals.pending()).toBeUndefined()
    await f.runtime.changes!.run(entry.id, 'resume')
    expect(f.proposals.pending()?.operation).toBe('resume_presentation_text_change')
    await f.proposals.confirm(f.proposals.pending()!.id)
    await f.runtime.changes!.refresh()
    expect(host.edit).toHaveBeenCalledOnce()
    expect(
      f.runtime.changes!.snapshot().entries.find((entry) => entry.kind === 'text')!.state,
    ).toBe('applied')
  } finally {
    f.runtime.dispose()
  }
})

it('keeps manual text changes intact when undo is requested from a historical workbench entry', async () => {
  const f = await fixture(),
    host = stablePageHost(f)
  try {
    await host.propose()
    await f.proposals.confirm(f.proposals.pending()!.id)
    await f.runtime.changes!.refresh()
    const entry = f.runtime.changes!.snapshot().entries.find((entry) => entry.kind === 'text')!
    host.read.mockResolvedValue({
      slideId: 'host',
      shapeId: 'shape',
      text: 'manual edit',
      paragraphs: ['manual edit'],
    })
    await f.runtime.changes!.run(entry.id, 'undo')
    expect(f.proposals.pending()).toBeUndefined()
    expect(host.edit).toHaveBeenCalledOnce()
    expect(f.runtime.changes!.snapshot().error).toBeTruthy()
    expect(f.binding.readTextChange()!.state).toBe('applied')
  } finally {
    f.runtime.dispose()
  }
})
