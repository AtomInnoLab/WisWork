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
    otherText = 'other-before',
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
  const otherShape = () => ({ ...shape(), id: 'other-shape', text: otherText })
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
    async (id) => ({
      slideId: id,
      shapes: id === 'slide' ? [shape()] : id === 'other' ? [otherShape()] : [],
    }),
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
    .mockImplementation(async (slide, _shape, next, expected) => {
      if (slide === 'other') {
        if (otherText !== expected) throw new Error('office_concurrent_change')
        otherText = next
        return
      }
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
      shapes: slideId === 'other' ? [otherShape()] : [shape()],
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
    otherText: () => otherText,
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
it('applies and reverses an ordered text plus geometry batch through one durable savepoint', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and placement',
    preserved: ['Other objects'],
    validation: ['Readback', 'Page screenshot'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  expect(f.binding().listChangeHistory()).toHaveLength(0)
  await f.confirm()
  const record = f.binding().listChangeHistory()[0]
  expect(record).toMatchObject({ kind: 'existing_batch', record: { state: 'applied', cursor: 2 } })
  expect(f.text()).toBe('after')
  expect(f.geometry().left).toBe(30)
  expect(f.invalidateQa).toHaveBeenCalledWith(['slide'])
  const listed = await f.call('list_existing_presentation_changes')
  expect(JSON.parse(listed.output).changes).toMatchObject([{ kind: 'batch', operationCount: 2 }])
  const captured = await f.call('capture_existing_presentation_batch_page', {
    change_id: changeId,
    slide_id: 'slide',
  })
  expect(captured.isError, captured.output).not.toBe(true)
  expect(JSON.parse(captured.output)).toMatchObject({ hostSlideId: 'slide', qaPassed: false })
  f.reopen()
  const workbench = f.getRuntime().changes!
  await workbench.refresh()
  const row = workbench
    .snapshot()
    .entries.find((entry) => entry.id === `existing_batch:${changeId}`)
  expect(row).toMatchObject({
    source: 'existing_batch',
    state: 'applied',
    actions: ['inspect', 'undo'],
  })
  await workbench.run(row!.id, 'undo')
  expect(f.getRuntime().proposals.pending()).toBeDefined()
  await f.confirm()
  expect(f.text()).toBe('before')
  expect(f.geometry().left).toBe(1)
  expect(f.binding().listChangeHistory()[0].record).toMatchObject({ state: 'undone', cursor: 0 })
})
it('edits two native objects on two pages and invalidates both QA page scopes', async () => {
  const f = await fixture()
  const baseline = await f.call('read_presentation_baseline', { scope: 'deck' })
  expect(baseline.isError, baseline.output).not.toBe(true)
  const baseline_id = JSON.parse(baseline.output).baselineId as string
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update both titles',
    preserved: ['Other objects'],
    validation: ['Review both pages'],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      { slide_id: 'other', shape_id: 'other-shape', kind: 'text', text: 'other-after' },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  await f.confirm()
  expect(f.text()).toBe('after')
  expect(f.otherText()).toBe('other-after')
  expect(f.invalidateQa).toHaveBeenCalledWith(['slide', 'other'])
  for (const slide_id of ['slide', 'other']) {
    const shot = await f.call('capture_existing_presentation_batch_page', {
      change_id: changeId,
      slide_id,
    })
    expect(shot.isError, shot.output).not.toBe(true)
  }
  f.reopen()
  const undo = await f.call('undo_existing_presentation_batch', { change_id: changeId })
  expect(undo.isError, undo.output).not.toBe(true)
  await f.confirm()
  expect(f.text()).toBe('before')
  expect(f.otherText()).toBe('other-before')
})
it('resumes a batch after the first step was saved and the second host write failed', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and placement',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  f.editGeometry.mockRejectedValueOnce(new Error('office_write_failed'))
  await expect(f.confirm()).rejects.toThrow()
  const entry = f.binding().listChangeHistory()[0]
  expect(entry).toMatchObject({ kind: 'existing_batch', record: { state: 'applying', cursor: 1 } })
  expect(f.text()).toBe('after')
  expect(f.geometry().left).toBe(1)
  f.reopen()
  const inspected = await f.call('inspect_existing_presentation_batch', { change_id: changeId })
  expect(JSON.parse(inspected.output).values).toEqual(['after', 'before'])
  const resumed = await f.call('resume_existing_presentation_batch', { change_id: changeId })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.editText).toHaveBeenCalledTimes(1)
  expect(f.binding().listChangeHistory()[0].record).toMatchObject({ state: 'applied', cursor: 2 })
})
it('finalizes an interrupted first-step receipt without replaying the host write', async () => {
  const f = await fixture()
  const baseline_id = await f.baseline()
  const proposed = await f.call('edit_existing_presentation_batch', {
    baseline_id,
    intent: 'Update title and placement',
    preserved: [],
    validation: [],
    risk: 'medium',
    operations: [
      { slide_id: 'slide', shape_id: 'shape', kind: 'text', text: 'after' },
      {
        slide_id: 'slide',
        shape_id: 'shape',
        kind: 'geometry',
        geometry: { ...f.geometry(), left: 30 },
      },
    ],
  })
  expect(proposed.isError, proposed.output).not.toBe(true)
  const changeId = JSON.parse(proposed.output).changeId as string
  let failed = false
  f.save.mockImplementation(async () => {
    if (!failed && f.editText.mock.calls.length === 1) {
      failed = true
      throw new Error('receipt_failed')
    }
  })
  await expect(f.confirm()).rejects.toThrow()
  expect(f.binding().readExistingBatch(changeId)).toMatchObject({ state: 'applying', cursor: 0 })
  expect(f.text()).toBe('after')
  f.reopen()
  const inspected = await f.call('inspect_existing_presentation_batch', { change_id: changeId })
  expect(JSON.parse(inspected.output).values).toEqual(['after', 'before'])
  const resumed = await f.call('resume_existing_presentation_batch', { change_id: changeId })
  expect(resumed.isError, resumed.output).not.toBe(true)
  await f.confirm()
  expect(f.editText).toHaveBeenCalledTimes(1)
  expect(f.binding().readExistingBatch(changeId)).toMatchObject({ state: 'applied', cursor: 2 })
})
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
