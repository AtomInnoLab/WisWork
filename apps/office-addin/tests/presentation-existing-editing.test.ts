import { afterEach, expect, it, vi } from 'vitest'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
import type { StructuredProposalController } from '../src/agent/proposal-controller'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document'
import { BrowserPresentationBaselineAdapter } from '../src/skills/powerpoint/browser-presentation-baseline-adapter'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
const otherPng =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})
async function fixture() {
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
  let location = 'file://existing.pptx',
    text = 'before',
    geometry = { left: 1, top: 2, width: 100, height: 40 },
    selection = ['shape'],
    screenshot = png
  const values = new Map<string, string>()
  const save = vi.fn(async () => {})
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save,
    location: () => location,
  }
  const bind = () => createPresentationDocumentBinding(settings, () => 'doc')
  const binding = bind()
  await binding.documentId()
  const shape = () => ({
    id: 'shape',
    name: 'Title',
    type: 'TextBox',
    ...geometry,
    text,
    font: { name: 'Arial', size: 20, color: '#000000' },
  })
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readContext').mockImplementation(
    async () => ({
      slideIds: ['slide', 'other'],
      selectedSlideIds: ['slide'],
      selectedShapeIds: selection,
      slideWidth: 960,
      slideHeight: 540,
    }),
  )
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readPage').mockImplementation(
    async (id) => ({ slideId: id, shapes: id === 'slide' ? [shape()] : [] }),
  )
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectSlideMasters').mockRejectedValue(
    new Error('office_api_unsupported'),
  )
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationPageText').mockImplementation(
    async (slideId, shapeId) => ({ slideId, shapeId, text, paragraphs: [text] }),
  )
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationPageGeometry').mockImplementation(
    async (slideId, shapeId) => ({ slideId, shapeId, geometry: { ...geometry } }),
  )
  const editText = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationPageText')
    .mockImplementation(async (_slide, _shape, next, expected) => {
      if (text !== expected) throw new Error('office_concurrent_change')
      text = next
    })
  const editGeometry = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationPageGeometry')
    .mockImplementation(async (_slide, _shape, next, expected) => {
      if (JSON.stringify(geometry) !== JSON.stringify(expected))
        throw new Error('office_concurrent_change')
      geometry = { ...next }
    })
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectPresentationPage').mockImplementation(
    async (slideId) => ({
      slideId,
      slideWidth: 960,
      slideHeight: 540,
      shapes: [shape()],
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: { mime: 'image/png', base64: screenshot },
    }),
  )
  const invalidateQa = vi.fn(async (_ids?: readonly string[]) => {})
  const create = () =>
    createOfficeHostRuntime('powerpoint', {
      presentation: {
        ...bind(),
        available: () => false,
        request: async () => {
          throw new Error('offline')
        },
        invalidateQa,
      },
    })
  let runtime = create()
  const call = (name: string, input: Record<string, unknown> = {}) =>
    runtime.skill.executeTool({ id: 'call', name, input })
  const baseline = async () => {
    const r = await call('read_presentation_baseline', { scope: 'selected' })
    expect(r.isError, r.output).not.toBe(true)
    return JSON.parse(r.output).baselineId as string
  }
  const propose = async (kind: 'text' | 'geometry' = 'text') => {
    const baseline_id = await baseline()
    const r = await call(`edit_existing_presentation_${kind}`, {
      baseline_id,
      slide_id: 'slide',
      shape_id: 'shape',
      ...(kind === 'text' ? { text: 'after' } : { geometry: { ...geometry, left: 30 } }),
    })
    expect(r.isError, r.output).not.toBe(true)
    return r
  }
  const confirm = () => {
    const proposals = runtime.proposals as StructuredProposalController
    return proposals.confirm(proposals.pending()!.id)
  }
  const records = () =>
    bind()
      .listChangeHistory()
      .filter((e) => e.kind === 'existing')
  return {
    call,
    baseline,
    propose,
    confirm,
    records,
    values,
    save,
    editText,
    editGeometry,
    invalidateQa,
    binding: bind,
    getRuntime: () => runtime,
    reopen: () => {
      runtime.dispose()
      runtime = create()
    },
    setText: (v: string) => {
      text = v
    },
    text: () => text,
    setLocation: (v: string) => {
      location = v
    },
    setSelection: (v: string[]) => {
      selection = v
    },
    setScreenshot: (v: string) => {
      screenshot = v
    },
    geometry: () => geometry,
  }
}
it.each(['text', 'geometry'] as const)(
  'persists a confirmed existing %s change and undoes it after reopening while offline',
  async (kind) => {
    const f = await fixture()
    await f.propose(kind)
    expect(f.records()).toHaveLength(0)
    expect(f.editText).not.toHaveBeenCalled()
    expect(f.editGeometry).not.toHaveBeenCalled()
    await f.confirm()
    const record = f.records()[0]!.record
    expect(record).toMatchObject({ state: 'applied', kind, hostSlideId: 'slide', shapeId: 'shape' })
    expect(record).not.toHaveProperty('projectId')
    expect(f.invalidateQa).toHaveBeenCalledWith(['slide'])
    f.reopen()
    const listed = await f.call('list_existing_presentation_changes')
    expect(JSON.parse(listed.output).changes).toHaveLength(1)
    expect(
      (await f.call('undo_existing_presentation_change', { change_id: record.changeId })).isError,
    ).not.toBe(true)
    await f.confirm()
    expect(f.records()[0]!.record.state).toBe('undone')
    expect(kind === 'text' ? f.text() : f.geometry().left).toBe(kind === 'text' ? 'before' : 1)
  },
)
it.each(['text', 'selection', 'document'] as const)(
  'rejects %s drift before confirmation without a host write',
  async (drift) => {
    const f = await fixture()
    await f.propose()
    if (drift === 'text') f.setText('manual')
    else if (drift === 'selection') f.setSelection([])
    else f.setLocation('file://save-as.pptx')
    await expect(f.confirm()).rejects.toThrow()
    expect(f.editText).not.toHaveBeenCalled()
    expect(f.records()).toHaveLength(0)
  },
)
it('saves pending before writing and only completes the receipt after a failed completion save', async () => {
  const f = await fixture()
  await f.propose()
  let failed = false
  f.save.mockImplementation(async () => {
    const raw = f.values.get('wiswork.presentation.existing-change.v1')
    if (!failed && raw && JSON.parse(raw).state === 'applied') {
      failed = true
      throw new Error('save_failed')
    }
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.text()).toBe('after')
  expect(f.records()[0]!.record.state).toBe('pending')
  const id = f.records()[0]!.record.changeId
  f.reopen()
  const inspection = JSON.parse(
    (await f.call('inspect_existing_presentation_change', { change_id: id })).output,
  )
  expect(inspection.status).toBe('already_applied')
  await f.call('resume_existing_presentation_change', { change_id: id })
  await f.confirm()
  expect(f.records()[0]!.record.state).toBe('applied')
  expect(f.editText).toHaveBeenCalledTimes(1)
})
it('does not write the host when the pending save fails and refuses undo after a manual target edit', async () => {
  const f = await fixture()
  await f.propose()
  f.save.mockRejectedValueOnce(new Error('save_failed'))
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editText).not.toHaveBeenCalled()
  await f.propose()
  await f.confirm()
  f.setText('manual')
  expect(
    (
      await f.call('undo_existing_presentation_change', {
        change_id: f.records()[0]!.record.changeId,
      })
    ).isError,
  ).toBe(true)
  expect(f.text()).toBe('manual')
})
it('binds visual review to a freshly captured unchanged page and preserves it as historical evidence', async () => {
  const f = await fixture()
  await f.propose()
  await f.confirm()
  const change_id = f.records()[0]!.record.changeId
  const capture = await f.call('capture_existing_presentation_change', { change_id })
  expect(capture.isError, capture.output).not.toBe(true)
  expect(capture.modelContent?.[0]?.type).toBe('image')
  const screenshot_digest = JSON.parse(capture.output).screenshotDigest
  const r = await f.call('record_existing_presentation_change_review', {
    change_id,
    screenshot_digest,
    status: 'pass',
    notes: '文字可读且位置合适',
  })
  expect(r.isError, r.output).not.toBe(true)
  expect(f.records()[0]!.record.review?.status).toBe('pass')
  f.setScreenshot(otherPng)
  expect(
    (
      await f.call('record_existing_presentation_change_review', {
        change_id,
        screenshot_digest,
        status: 'pass',
        notes: 'old',
      })
    ).isError,
  ).toBe(true)
  f.reopen()
  expect(
    (
      await f.call('record_existing_presentation_change_review', {
        change_id,
        screenshot_digest,
        status: 'pass',
        notes: 'old',
      })
    ).isError,
  ).toBe(true)
})
it('keeps proposals stale after clearing the session and rejects out-of-scope shapes', async () => {
  const f = await fixture(),
    baseline_id = await f.baseline()
  expect(
    (
      await f.call('edit_existing_presentation_text', {
        baseline_id,
        slide_id: 'other',
        shape_id: 'shape',
        text: 'no',
      })
    ).isError,
  ).toBe(true)
  await f.propose()
  const proposals = f.getRuntime().proposals as StructuredProposalController
  const id = proposals.pending()!.id
  f.getRuntime().clearSession()
  await expect(proposals.confirm(id)).rejects.toThrow('proposal_missing')
  expect(f.editText).not.toHaveBeenCalled()
})
it('recovers an interrupted undo after reopening without writing the host twice', async () => {
  const f = await fixture()
  await f.propose()
  await f.confirm()
  const change_id = f.records()[0]!.record.changeId
  await f.call('undo_existing_presentation_change', { change_id })
  let failed = false
  f.save.mockImplementation(async () => {
    const raw = f.values.get('wiswork.presentation.existing-change.v1')
    if (!failed && raw && JSON.parse(raw).state === 'undone') {
      failed = true
      throw new Error('save_failed')
    }
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.text()).toBe('before')
  expect(f.records()[0]!.record.state).toBe('undo_pending')
  f.reopen()
  await f.call('resume_existing_presentation_change', { change_id })
  await f.confirm()
  expect(f.records()[0]!.record.state).toBe('undone')
  expect(f.editText).toHaveBeenCalledTimes(2)
})
it('revalidates after QA invalidation waits and after pending settings save waits', async () => {
  const f = await fixture()
  await f.propose()
  f.invalidateQa.mockImplementationOnce(async () => {
    f.setText('manual during hook')
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editText).not.toHaveBeenCalled()
  f.setText('before')
  await f.propose()
  f.save.mockImplementationOnce(async () => {
    f.setSelection([])
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.editText).not.toHaveBeenCalled()
  expect(f.records()[0]!.record.state).toBe('pending')
})
it('blocks a new existing edit while a generated text transaction is unresolved', async () => {
  const f = await fixture(),
    b = f.binding(),
    documentId = await b.documentId()
  await b.writeTextChange(
    {
      version: 1,
      changeId: 'generated',
      documentId,
      projectId: 'project',
      requestId: 'request',
      artifactDigest: 'a'.repeat(64),
      pageId: 'page',
      hostSlideId: 'slide',
      shapeId: 'shape',
      before: 'before',
      after: 'generated',
      state: 'pending',
    },
    undefined,
  )
  await f.propose()
  await expect(f.confirm()).rejects.toThrow('presentation_change_history_pending')
  expect(f.editText).not.toHaveBeenCalled()
})
it('reopens two existing changes and undoes through the offline workbench in dependency order', async () => {
  const f = await fixture()
  await f.propose()
  await f.confirm()
  await f.propose('geometry')
  await f.confirm()
  f.reopen()
  const workbench = f.getRuntime().changes!
  await workbench.refresh()
  expect(workbench.snapshot().entries).toHaveLength(2)
  for (const entry of workbench.snapshot().entries) {
    await workbench.run(entry.id, 'undo')
    expect(f.getRuntime().proposals.pending(), JSON.stringify(workbench.snapshot())).toBeDefined()
    await f.confirm()
    await workbench.refresh()
  }
  expect(f.records().map((e) => e.record.state)).toEqual(['undone', 'undone'])
  expect(f.text()).toBe('before')
  expect(f.geometry().left).toBe(1)
})
it('rejects a changed screenshot before storing a review and invalidates a capture when another change is confirmed', async () => {
  const f = await fixture()
  await f.propose()
  await f.confirm()
  const change_id = f.records()[0]!.record.changeId
  const r = await f.call('capture_existing_presentation_change', { change_id }),
    screenshot_digest = JSON.parse(r.output).screenshotDigest
  f.setScreenshot(otherPng)
  expect(
    (
      await f.call('record_existing_presentation_change_review', {
        change_id,
        screenshot_digest,
        status: 'pass',
        notes: 'mismatch',
      })
    ).output,
  ).toBe('presentation_existing_qa_stale')
  expect(f.records()[0]!.record.review).toBeUndefined()
  f.setScreenshot(png)
  await f.propose('geometry')
  await f.confirm()
  expect(
    (
      await f.call('record_existing_presentation_change_review', {
        change_id,
        screenshot_digest,
        status: 'pass',
        notes: 'stale',
      })
    ).output,
  ).toBe('presentation_existing_qa_stale')
})
it('does not reuse a baseline after a successful edit and validates malformed tool input before proposing', async () => {
  const f = await fixture(),
    baseline_id = await f.baseline()
  for (const input of [
    { baseline_id, slide_id: 'slide', shape_id: 'shape', text: 'after', extra: true },
    { baseline_id, slide_id: 'slide', shape_id: 'shape', text: 'a'.repeat(12001) },
  ])
    expect((await f.call('edit_existing_presentation_text', input)).output).toBe(
      'invalid_tool_input',
    )
  await f.call('edit_existing_presentation_text', {
    baseline_id,
    slide_id: 'slide',
    shape_id: 'shape',
    text: 'after',
  })
  await f.confirm()
  expect(
    (
      await f.call('edit_existing_presentation_text', {
        baseline_id,
        slide_id: 'slide',
        shape_id: 'shape',
        text: 'again',
      })
    ).output,
  ).toBe('presentation_baseline_changed')
  expect(f.editText).toHaveBeenCalledTimes(1)
})
