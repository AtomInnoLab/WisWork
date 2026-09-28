import { createHash } from 'node:crypto'
import { expect, it, vi } from 'vitest'
import {
  createPresentationQaSkill,
  type PresentationQaRecord,
  validatePresentationQaRecord,
} from '../src/skills/powerpoint/presentation-qa.js'
import type { PresentationImportRecord } from '../src/skills/powerpoint/presentation-delivery.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
function setup() {
  const artifact = {
    documentId: 'doc',
    projectId: 'project',
    requestId: 'request',
    pptxBase64: 'UEsDBAAAAAA=',
    slideCount: 2,
    pages: [
      { id: 'first', title: 'First', sourceSlideId: '256#' },
      { id: 'second', title: 'Second', sourceSlideId: '257#' },
    ],
  }
  const artifactDigest = createHash('sha256').update(artifact.pptxBase64).digest('hex')
  const receipt = {
    state: 'pending' as const,
    documentId: 'doc',
    checkpoint: {
      version: 1 as const,
      artifactDigest,
      sourceSlideIds: ['256#', '257#'],
      baselineSlideIds: ['original'],
      completed: [{ sourceSlideId: '256#', slideId: 'host1' }],
    },
  }
  const inspectPage = vi.fn(async () => ({
    slideId: 'host1',
    slideWidth: 960,
    slideHeight: 540,
    shapes: [],
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: png },
  }))
  let record: PresentationQaRecord | undefined
  const readQa = () => record,
    writeQa = vi.fn(async (_key: string, value: PresentationQaRecord) => {
      record = structuredClone(value)
    })
  const vfs = new InMemoryVfs(),
    documentId = vi.fn(async () => 'doc')
  const options = {
    available: () => true,
    artifact: () => artifact,
    documentId,
    readReceipt: () => receipt,
    inspectPage,
    readQa,
    writeQa,
    vfs,
  }
  const skill = createPresentationQaSkill(options),
    capture = { id: 'capture', name: 'capture_presentation_page_qa', input: { page_id: 'first' } }
  return {
    artifact,
    receipt,
    inspectPage,
    readQa,
    writeQa,
    vfs,
    documentId,
    options,
    skill,
    capture,
  }
}
it('captures the exact imported page, publishes real screenshot and persists only metadata', async () => {
  const f = setup(),
    result = await f.skill.executeTool(f.capture)
  expect(result.isError).not.toBe(true)
  expect(f.inspectPage).toHaveBeenCalledWith('host1', undefined)
  expect(result.modelContent).toEqual([
    { type: 'image', image: { mime: 'image/png', base64: png } },
  ])
  expect(JSON.stringify(f.readQa())).not.toContain(png)
  expect(f.readQa()?.pages[0]).toMatchObject({
    hostSlideId: 'host1',
    visual: { status: 'needs_review' },
    structure: { status: 'passed' },
  })
  expect(f.vfs.list('/home/user')).toContain('/home/user/generated/qa-project-first.png')
  expect(validatePresentationQaRecord(f.readQa())).toBe(true)
})
it('records an agent review only after a live capture and unchanged recapture', async () => {
  const f = setup()
  await f.skill.executeTool(f.capture)
  const input = {
    page_id: 'first',
    screenshot_digest: f.readQa()!.pages[0]!.screenshotDigest,
    outcome: 'pass',
    notes: 'Reviewed screenshot: title is legible.',
  }
  expect(
    await createPresentationQaSkill(f.options).executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input,
    }),
  ).toMatchObject({ isError: true, output: 'presentation_qa_capture_required' })
  expect(
    await f.skill.executeTool({ id: 'review', name: 'record_presentation_page_review', input }),
  ).not.toHaveProperty('isError', true)
  expect(f.readQa()?.pages[0]?.visual).toMatchObject({ status: 'pass', reviewer: 'agent' })
  expect(f.inspectPage).toHaveBeenCalledTimes(2)
  const read = await f.skill.executeTool({ id: 'read', name: 'read_presentation_qa', input: {} })
  expect(JSON.parse(read.output)).toHaveProperty('needs_recapture', true)
})
it('rejects pages not yet imported and changed structure before review', async () => {
  const f = setup()
  expect(await f.skill.executeTool({ ...f.capture, input: { page_id: 'second' } })).toMatchObject({
    isError: true,
    output: 'presentation_qa_page_not_imported',
  })
  await f.skill.executeTool(f.capture)
  f.inspectPage.mockResolvedValue({ ...(await f.inspectPage()), slideWidth: 961 })
  expect(
    await f.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: 'first',
        screenshot_digest: f.readQa()!.pages[0]!.screenshotDigest,
        outcome: 'pass',
        notes: 'okay',
      },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_qa_stale' })
  expect(f.readQa()?.pages[0]?.visual.status).toBe('needs_review')
})
it('publishes neither screenshot nor approval after save failure or lifecycle clear', async () => {
  const f = setup()
  f.writeQa.mockRejectedValue(new Error('save_failed'))
  expect(await f.skill.executeTool(f.capture)).toMatchObject({ isError: true })
  expect(f.vfs.list('/home/user')).toEqual([])
  const g = setup(),
    original = g.inspectPage.getMockImplementation()!
  g.inspectPage.mockImplementation(async () => {
    g.skill.clear()
    return original()
  })
  expect(await g.skill.executeTool(g.capture)).toMatchObject({ isError: true, output: 'cancelled' })
  expect(g.writeQa).not.toHaveBeenCalled()
})
it('rejects foreign document changes, forged exact-page replies and invalid input without passing', async () => {
  const f = setup(),
    original = f.inspectPage.getMockImplementation()!
  f.inspectPage.mockImplementation(async () => {
    f.documentId.mockResolvedValue('foreign')
    return original()
  })
  expect(await f.skill.executeTool(f.capture)).toMatchObject({
    isError: true,
    output: 'presentation_document_changed',
  })
  expect(f.writeQa).not.toHaveBeenCalled()
  const g = setup()
  g.inspectPage.mockResolvedValue({ ...(await g.inspectPage()), slideId: 'wrong' })
  expect(await g.skill.executeTool(g.capture)).toMatchObject({
    isError: true,
    output: 'presentation_qa_capture_invalid',
  })
  expect(
    await g.skill.executeTool({ ...g.capture, input: { page_id: 'first', unexpected: true } }),
  ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
})
it('bounds metadata and rejects unknown fields, false structural passes and malformed reviews', async () => {
  const f = setup()
  await f.skill.executeTool(f.capture)
  const record = f.readQa()!,
    page = record.pages[0]!
  for (const changed of [
    { ...record, unexpected: true },
    { ...record, pages: [page, page] },
    { ...record, pages: [{ ...page, screenshotBytes: 2 * 1024 * 1024 + 1 }] },
    { ...record, pages: [{ ...page, capturedAt: 'tomorrow' }] },
    { ...record, pages: [{ ...page, structure: { ...page.structure, shapeCount: 101 } }] },
    { ...record, pages: [{ ...page, structure: { ...page.structure, overflowCount: 401 } }] },
    { ...record, pages: [{ ...page, structure: { ...page.structure, overlapCount: 1 } }] },
    {
      ...record,
      pages: [
        {
          ...page,
          visual: { status: 'pass', reviewer: 'human', notes: 'OK', reviewedAt: page.capturedAt },
        },
      ],
    },
    { ...record, pages: [{ ...page, visual: { status: 'needs_review', notes: 'claimed' } }] },
  ])
    expect(validatePresentationQaRecord(changed)).toBe(false)
  expect(
    validatePresentationQaRecord({
      ...record,
      pages: Array.from({ length: 32 }, (_, i) => ({
        ...page,
        pageId: `page${i}`,
        hostSlideId: `host${i}`,
        visual: {
          status: 'needs_changes',
          reviewer: 'agent',
          notes: 'x'.repeat(2000),
          reviewedAt: page.capturedAt,
        },
      })),
    }),
  ).toBe(false)
})
it('reports truncation as incomplete and review changes do not mark the page passed', async () => {
  const f = setup()
  f.inspectPage.mockResolvedValue({ ...(await f.inspectPage()), shapesTruncated: true })
  await f.skill.executeTool(f.capture)
  expect(f.readQa()?.pages[0]?.structure.status).toBe('incomplete')
  const old = f.readQa()!.pages[0]!
  f.inspectPage.mockResolvedValue({
    ...(await f.inspectPage()),
    screenshot: { mime: 'image/png', base64: png.replace('P8/x8A', 'P8/x9A') },
  })
  expect(
    await f.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: 'first',
        screenshot_digest: old.screenshotDigest,
        outcome: 'pass',
        notes: 'Looks fine',
      },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_qa_stale' })
  expect(f.readQa()?.pages[0]?.visual.status).toBe('needs_review')
})
it('serializes QA work and blocks cancellation before screenshot publication', async () => {
  const f = setup(),
    original = f.inspectPage.getMockImplementation()!
  let resolve!: () => void
  const waiting = new Promise<void>((done) => {
    resolve = done
  })
  f.inspectPage.mockImplementation(async () => {
    await waiting
    return original()
  })
  const controller = new AbortController(),
    pending = f.skill.executeTool(f.capture, controller.signal)
  expect(
    await f.skill.executeTool({ id: 'read', name: 'read_presentation_qa', input: {} }),
  ).toMatchObject({ isError: true, output: 'presentation_qa_busy' })
  controller.abort()
  resolve()
  expect(await pending).toMatchObject({ isError: true, output: 'cancelled' })
  expect(f.writeQa).not.toHaveBeenCalled()
  expect(f.vfs.list('/home/user')).toEqual([])
})
it('requires exact import digest and treats missing historical reports as empty', async () => {
  const f = setup()
  expect(
    JSON.parse(
      (await f.skill.executeTool({ id: 'read', name: 'read_presentation_qa', input: {} })).output,
    ),
  ).toMatchObject({ record: null, needs_recapture: true })
  f.receipt.checkpoint.artifactDigest = 'a'.repeat(64)
  expect(await f.skill.executeTool(f.capture)).toMatchObject({
    isError: true,
    output: 'presentation_qa_page_not_imported',
  })
  expect(f.inspectPage).not.toHaveBeenCalled()
})
it('does not allow review when a capture could not publish its image to the model', async () => {
  const f = setup(),
    options = { ...f.options, vfs: new InMemoryVfs({ maxTotalBytes: 1 }) },
    skill = createPresentationQaSkill(options)
  const result = await skill.executeTool(f.capture)
  expect(result).toMatchObject({ isError: true })
  expect(result.modelContent).toBeUndefined()
  const page = f.readQa()!.pages[0]!
  expect(page.visual.status).toBe('needs_review')
  expect(
    await skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: 'first',
        screenshot_digest: page.screenshotDigest,
        outcome: 'pass',
        notes: 'okay',
      },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_qa_capture_required' })
})
it('leaves the page waiting for a screenshot after host capture failure without replacing a prior review', async () => {
  const f = setup()
  await f.skill.executeTool(f.capture)
  const before = structuredClone(f.readQa()!)
  f.inspectPage.mockRejectedValue(
    Object.assign(new Error('office_read_failed'), { code: 'office_screenshot_unavailable' }),
  )
  const result = await f.skill.executeTool(f.capture)
  expect(result).toMatchObject({
    output: expect.stringContaining('waiting_screenshot'),
    mutated: false,
  })
  expect(result.modelContent).toBeUndefined()
  expect(f.readQa()).toEqual(before)
  expect(f.writeQa).toHaveBeenCalledTimes(1)
  expect(
    await f.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: 'first',
        screenshot_digest: before.pages[0]!.screenshotDigest,
        outcome: 'pass',
        notes: 'Reviewed',
      },
    }),
  ).toMatchObject({ output: expect.stringContaining('waiting_screenshot') })
  expect(f.readQa()).toEqual(before)
})
it('keeps structural and unsupported API errors distinct from missing screenshots', async () => {
  const f = setup()
  f.inspectPage.mockRejectedValueOnce(new Error('office_read_failed'))
  expect(await f.skill.executeTool(f.capture)).toMatchObject({
    output: 'office_read_failed',
    isError: true,
  })
  f.inspectPage.mockRejectedValueOnce(new Error('office_api_unsupported'))
  expect(await f.skill.executeTool(f.capture)).toMatchObject({
    output: 'office_api_unsupported',
    isError: true,
  })
  f.inspectPage.mockRejectedValueOnce(new Error('cancelled'))
  expect(await f.skill.executeTool(f.capture)).toMatchObject({ output: 'cancelled', isError: true })
  expect(f.writeQa).not.toHaveBeenCalled()
})
it('does not report a stale import as merely waiting for a screenshot', async () => {
  const f = setup()
  f.inspectPage.mockImplementationOnce(async () => {
    f.receipt.checkpoint.artifactDigest = 'a'.repeat(64)
    throw Object.assign(new Error('office_read_failed'), { code: 'office_screenshot_unavailable' })
  })
  expect(await f.skill.executeTool(f.capture)).toMatchObject({
    output: 'presentation_qa_stale',
    isError: true,
  })
  expect(f.writeQa).not.toHaveBeenCalled()
})
it('invalidates prior visual review after recapture', async () => {
  const f = setup()
  await f.skill.executeTool(f.capture)
  const old = f.readQa()!.pages[0]!
  await f.skill.executeTool({
    id: 'review',
    name: 'record_presentation_page_review',
    input: {
      page_id: 'first',
      screenshot_digest: old.screenshotDigest,
      outcome: 'needs_changes',
      notes: 'Title needs more contrast',
    },
  })
  expect(f.readQa()!.pages[0]!.visual.status).toBe('needs_changes')
  await f.skill.executeTool(f.capture)
  expect(f.readQa()!.pages[0]!.visual).toEqual({ status: 'needs_review' })
  expect(f.vfs.list('/home/user')).toContain('/home/user/generated/project.qa.json')
})
it('accepts only an explicit true recheck flag and preserves the old review as history until recapture', async () => {
  const f = setup()
  await f.skill.executeTool(f.capture)
  const page = f.readQa()!.pages[0]!
  await f.skill.executeTool({
    id: 'review',
    name: 'record_presentation_page_review',
    input: {
      page_id: 'first',
      screenshot_digest: page.screenshotDigest,
      outcome: 'pass',
      notes: 'The title is legible',
    },
  })
  const old = structuredClone(f.readQa()!)
  const flagged = {
    ...old,
    pages: old.pages.map((p) => ({ ...p, recheckRequired: true as const })),
  }
  expect(validatePresentationQaRecord(old)).toBe(true)
  expect(validatePresentationQaRecord(flagged)).toBe(true)
  for (const value of [false, null, 'true', 1])
    expect(
      validatePresentationQaRecord({
        ...old,
        pages: old.pages.map((p) => ({ ...p, recheckRequired: value })),
      }),
    ).toBe(false)
  await f.writeQa('project/request', flagged)
  const historical = await f.skill.executeTool({
    id: 'read',
    name: 'read_presentation_qa',
    input: {},
  })
  expect(JSON.parse(historical.output).record.pages[0]).toMatchObject({
    recheckRequired: true,
    visual: { status: 'pass', reviewer: 'agent' },
  })
  expect(
    await f.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: 'first',
        screenshot_digest: page.screenshotDigest,
        outcome: 'pass',
        notes: 'Reuse old review',
      },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_qa_capture_required' })
  await f.skill.executeTool(f.capture)
  expect(f.readQa()!.pages[0]!.recheckRequired).toBeUndefined()
  expect(f.readQa()!.pages[0]!.visual.status).toBe('needs_review')
})
it('holds the mutation lock across session clear and invalidates earlier live captures', async () => {
  const f = setup()
  await f.skill.executeTool(f.capture)
  const screenshotDigest = f.readQa()!.pages[0]!.screenshotDigest
  f.skill.beginMutation()
  expect(() => f.skill.beginMutation()).toThrow('presentation_qa_busy')
  f.skill.clear()
  for (const call of [
    f.capture,
    { id: 'read', name: 'read_presentation_qa', input: {} },
    {
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: 'first',
        screenshot_digest: screenshotDigest,
        outcome: 'pass',
        notes: 'Old image',
      },
    },
  ]) {
    expect(await f.skill.executeTool(call)).toMatchObject({
      isError: true,
      output: 'presentation_qa_busy',
    })
  }
  f.skill.endMutation()
  expect(
    await f.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: 'first',
        screenshot_digest: screenshotDigest,
        outcome: 'pass',
        notes: 'Old image',
      },
    }),
  ).toMatchObject({ isError: true, output: 'presentation_qa_capture_required' })
  expect(await f.skill.executeTool(f.capture)).not.toHaveProperty('isError', true)
})
it('rejects mutation while a QA operation is in flight without cancelling that operation', async () => {
  const f = setup()
  const pending = f.skill.executeTool(f.capture)
  expect(() => f.skill.beginMutation()).toThrow('presentation_qa_busy')
  expect(await pending).not.toHaveProperty('isError', true)
  f.skill.beginMutation()
  f.skill.endMutation()
})

it('invalidates only target live screenshots and permits review of an unrelated captured page', async () => {
  const f = setup()
  f.artifact.slideCount = 3
  f.artifact.pages.push({ id: 'third', title: 'Third', sourceSlideId: '258#' })
  f.receipt.checkpoint.sourceSlideIds.push('258#')
  f.receipt.checkpoint.completed.push({ sourceSlideId: '257#', slideId: 'host2' })
  f.inspectPage.mockImplementation(async (...args: unknown[]) => ({
    slideId: String(args[0]),
    slideWidth: 960,
    slideHeight: 540,
    shapes: [],
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: png },
  }))
  await f.skill.executeTool(f.capture)
  await f.skill.executeTool({ ...f.capture, input: { page_id: 'second' } })
  const digest = f.readQa()!.pages[0]!.screenshotDigest
  f.skill.beginMutation(['host1'])
  f.skill.endMutation()
  const review = (page_id: string) => ({
    id: 'review',
    name: 'record_presentation_page_review',
    input: {
      page_id,
      screenshot_digest: digest,
      outcome: 'pass',
      notes: 'Reviewed current screenshot',
    },
  })
  expect(await f.skill.executeTool(review('first'))).toMatchObject({
    isError: true,
    output: 'presentation_qa_capture_required',
  })
  expect(await f.skill.executeTool(review('second'))).not.toHaveProperty('isError', true)
})
it('rejects invalid mutation scopes without losing live evidence or acquiring a lock', async () => {
  const f = setup()
  await f.skill.executeTool(f.capture)
  for (const scope of [
    ['host1', 'host1'],
    [''],
    ['bad\n'],
    ['bad\x7f'],
    ['x'.repeat(257)],
    Array.from({ length: 101 }, (_, i) => String(i)),
    null,
  ])
    expect(() => f.skill.beginMutation(scope as string[])).toThrow('invalid_tool_input')
  const screenshot_digest = f.readQa()!.pages[0]!.screenshotDigest
  expect(
    await f.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: { page_id: 'first', screenshot_digest, outcome: 'pass', notes: 'Unchanged' },
    }),
  ).not.toHaveProperty('isError', true)
})
it('preserves live evidence for an unmatched maximum-sized scope', async () => {
  const f = setup()
  await f.skill.executeTool(f.capture)
  f.skill.beginMutation(Array.from({ length: 100 }, (_, i) => String(i).padStart(256, 'x')))
  expect(await f.skill.executeTool(f.capture)).toMatchObject({
    isError: true,
    output: 'presentation_qa_busy',
  })
  f.skill.endMutation()
  expect(
    await f.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: {
        page_id: 'first',
        screenshot_digest: f.readQa()!.pages[0]!.screenshotDigest,
        outcome: 'pass',
        notes: 'Unchanged page',
      },
    }),
  ).not.toHaveProperty('isError', true)
})

function productionSetup() {
  const f = setup()
  const artifact = {
    ...f.artifact,
    pptxBase64: '',
    pagePptxBase64: ['UEsDBAAAAAA=', 'UEsDBAEAAAA='],
    planRevision: 1,
    pages: f.artifact.pages.map((p) => ({ ...p, sourceSlideId: '256#' })),
  }
  const artifactDigest = createHash('sha256')
    .update(
      JSON.stringify({
        documentId: artifact.documentId,
        projectId: artifact.projectId,
        requestId: artifact.requestId,
        planRevision: artifact.planRevision,
        pages: artifact.pages,
        pagePptxBase64: artifact.pagePptxBase64,
      }),
    )
    .digest('hex')
  const receipt: PresentationImportRecord = {
    ...f.receipt,
    state: 'complete',
    slideIds: ['host1', 'host2'],
    checkpoint: {
      ...f.receipt.checkpoint,
      version: 2 as const,
      artifactDigest,
      pageIds: ['first', 'second'],
      sourceSlideIds: ['256#', '256#'],
      completed: [
        { sourceSlideId: '256#', slideId: 'host1' },
        { sourceSlideId: '256#', slideId: 'host2' },
      ],
    },
  }
  const records = new Map<string, PresentationQaRecord>()
  const options = {
    ...f.options,
    artifact: () => artifact,
    readReceipt: (key: string) => (key === 'production/project/request' ? receipt : undefined),
    readQa: (key: string) => records.get(key),
    writeQa: vi.fn(async (key: string, value: PresentationQaRecord) => {
      records.set(key, structuredClone(value))
    }),
  }
  f.inspectPage.mockImplementation(async () => ({
    slideId: 'host2',
    slideWidth: 960,
    slideHeight: 540,
    shapes: [],
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: png },
  }))
  return {
    ...f,
    artifact,
    receipt,
    records,
    options,
    skill: createPresentationQaSkill(options),
    capture: { ...f.capture, input: { page_id: 'second' } },
  }
}
it('captures the second production business page with repeated source IDs and isolated QA', async () => {
  const f = productionSetup()
  expect((await f.skill.executeTool(f.capture)).isError).not.toBe(true)
  expect(f.inspectPage).toHaveBeenCalledWith('host2', undefined)
  expect(f.records.get('production/project/request')).toMatchObject({
    source: 'production',
    pages: [{ pageId: 'second', hostSlideId: 'host2' }],
  })
  expect(f.records.has('project/request')).toBe(false)
})
it('rejects changed production bytes or page order and uncompleted pages', async () => {
  for (const change of ['bytes', 'order', 'pending'] as const) {
    const f = productionSetup()
    if (change === 'bytes') f.artifact.pagePptxBase64[1] = 'UEsDBAIAAAA='
    if (change === 'order') f.artifact.pages.reverse()
    if (change === 'pending') {
      f.receipt.state = 'pending'
      delete f.receipt.slideIds
      f.receipt.checkpoint!.completed.pop()
    }
    expect((await f.skill.executeTool(f.capture)).isError).toBe(true)
    expect(f.inspectPage).not.toHaveBeenCalled()
  }
})
it('rejects production changes during awaited capture without publishing QA', async () => {
  const f = productionSetup(),
    original = f.inspectPage.getMockImplementation()!
  f.inspectPage.mockImplementation(async () => {
    const result = await original()
    f.artifact.planRevision++
    return result
  })
  expect((await f.skill.executeTool(f.capture)).isError).toBe(true)
  expect(f.options.writeQa).not.toHaveBeenCalled()
})

it('reviews a live production capture and rejects a mixed-source stored record', async () => {
  const f = productionSetup()
  const captured = await f.skill.executeTool(f.capture)
  const record = f.records.get('production/project/request')!
  const result = await f.skill.executeTool({
    id: 'review',
    name: 'record_presentation_page_review',
    input: {
      page_id: 'second',
      screenshot_digest: record.pages[0]!.screenshotDigest,
      outcome: 'pass',
      notes: '布局清晰',
    },
  })
  expect(result.isError).not.toBe(true)
  expect(f.inspectPage).toHaveBeenCalledTimes(2)
  expect(f.records.get('production/project/request')?.pages[0]?.visual.status).toBe('pass')
  expect(JSON.parse(captured.output).path).toMatch(/qa-[a-f0-9]{64}\.png$/)
  const contaminated = structuredClone(record)
  delete contaminated.source
  f.records.set('production/project/request', contaminated)
  expect(await f.skill.executeTool(f.capture)).toMatchObject({
    isError: true,
    output: 'presentation_qa_state_invalid',
  })
})
it('allows only confirmed production prefix pages while the next page is uncertain', async () => {
  const f = productionSetup()
  f.receipt.state = 'pending'
  delete f.receipt.slideIds
  f.receipt.checkpoint!.completed.pop()
  f.receipt.checkpoint!.inFlight = { sourceSlideId: '256#' }
  expect(await f.skill.executeTool(f.capture)).toMatchObject({
    isError: true,
    output: 'presentation_qa_page_not_imported',
  })
  const original = f.inspectPage.getMockImplementation()!
  f.inspectPage.mockImplementation(async () => ({ ...(await original()), slideId: 'host1' }))
  expect(
    (await f.skill.executeTool({ ...f.capture, input: { page_id: 'first' } })).isError,
  ).not.toBe(true)
  expect(f.inspectPage).toHaveBeenCalledWith('host1', undefined)
})
it('keeps production screenshot and metadata paths distinct between requests', async () => {
  const paths = []
  for (const request of ['request', 'other']) {
    const f = productionSetup()
    f.artifact.requestId = request
    // The artifact and receipt are rebuilt together, as when a separate production is selected.
    f.receipt.checkpoint!.artifactDigest = createHash('sha256')
      .update(
        JSON.stringify({
          documentId: f.artifact.documentId,
          projectId: f.artifact.projectId,
          requestId: request,
          planRevision: f.artifact.planRevision,
          pages: f.artifact.pages,
          pagePptxBase64: f.artifact.pagePptxBase64,
        }),
      )
      .digest('hex')
    f.options.readReceipt = () => f.receipt
    const result = await createPresentationQaSkill(f.options).executeTool(f.capture)
    expect(result.isError).not.toBe(true)
    paths.push(JSON.parse(result.output).path)
    expect(f.vfs.list('/home/user').filter((path) => path.endsWith('.json'))).toHaveLength(1)
  }
  expect(new Set(paths).size).toBe(2)
})

it('locks an explicitly empty mutation scope without discarding any live capture', async () => {
  const f = setup()
  await f.skill.executeTool(f.capture)
  const screenshot_digest = f.readQa()!.pages[0]!.screenshotDigest
  f.skill.beginMutation([])
  expect((await f.skill.executeTool(f.capture)).output).toBe('presentation_qa_busy')
  f.skill.endMutation()
  expect(
    await f.skill.executeTool({
      id: 'review',
      name: 'record_presentation_page_review',
      input: { page_id: 'first', screenshot_digest, outcome: 'pass', notes: 'Unchanged page' },
    }),
  ).not.toHaveProperty('isError', true)
})
