import { createHash } from 'node:crypto'
import { expect, it, vi } from 'vitest'
import {
  createPresentationQaSkill,
  type PresentationQaRecord,
  validatePresentationQaRecord,
} from '../src/skills/powerpoint/presentation-qa.js'
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
