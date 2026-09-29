import { afterEach, expect, it, vi } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document.js'
import { createPresentationQaSkill } from '../src/skills/powerpoint/presentation-qa.js'
import { createPresentationProductionDeliverySkill } from '../src/skills/powerpoint/presentation-page-delivery.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
import { createOfficeDiagnostics } from '../src/diagnostics/office-diagnostics.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import type {
  CompiledPresentationArtifact,
  PresentationDeliveryOptions,
} from '../src/skills/powerpoint/presentation-delivery.js'
import type { PresentationImportProgress } from '../src/skills/powerpoint/presentation-page-delivery.js'
import type { PresentationProjectStatus } from '../src/skills/powerpoint/presentation-project.js'

const png =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII='
afterEach(() => {
  vi.useRealTimers()
  vi.restoreAllMocks()
})
async function fixture() {
  const slides = [
    { id: 'one', title: 'One' },
    { id: 'two', title: 'Two' },
  ]
  const pages = await Promise.all(
    slides.map(async (slide) => {
      const compiled = await compilePresentationDeck({
        version: 1,
        id: slide.id,
        title: slide.title,
        style: {
          fontFace: 'Noto Sans CJK SC',
          background: 'FFFFFF',
          textColor: '173248',
          accentColor: '087D83',
        },
        assets: [],
        claims: [],
        slides: [
          {
            ...slide,
            claimIds: [],
            elements: [
              {
                kind: 'text',
                id: 'title',
                x: 1,
                y: 1,
                w: 10,
                h: 1,
                text: slide.title,
                fontSize: 24,
              },
            ],
          },
        ],
      })
      return Buffer.from(compiled.bytes).toString('base64')
    }),
  )
  const values = new Map<string, unknown>()
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: unknown) => {
      values.set(key, value)
    },
    save: async () => {},
    location: () => '',
  }
  let binding = createPresentationDocumentBinding(settings)
  const documentId = await binding.documentId()
  const artifact: CompiledPresentationArtifact = {
    documentId,
    projectId: 'qa-project',
    requestId: 'run',
    planRevision: 1,
    pptxBase64: '',
    pagePptxBase64: pages,
    slideCount: 2,
    pages: slides.map((slide) => ({ ...slide, sourceSlideId: '256#' })),
  }
  const host = ['original']
  const proposals = createStructuredProposalController()
  const options: PresentationDeliveryOptions = {
    available: () => true,
    artifact: () => artifact,
    proposals,
    documentId: () => binding.documentId(),
    readReceipt: (key) => binding.readReceipt(key),
    writeReceipt: (key, value) => binding.writeReceipt(key, value),
    adapter: {
      available: () => true,
      snapshot: async () => ({ slideIds: [...host], fingerprint: JSON.stringify(host) }),
      insert: vi.fn(),
      insertPage: async () => {
        const id = `host-${host.length}`
        host.push(id)
        return { slideIds: [id] }
      },
      verify: async () => true,
    },
  }
  const delivery = createPresentationProductionDeliverySkill(options)
  expect(
    (
      await delivery.executeTool({
        id: 'import',
        name: 'import_presentation_production',
        input: {},
      })
    ).isError,
  ).not.toBe(true)
  await proposals.confirm(proposals.pending()!.id)
  const importedResult = await delivery.executeTool({
    id: 'import-status',
    name: 'read_presentation_production_import_status',
    input: {},
  })
  expect(importedResult.isError).not.toBe(true)
  const imported = JSON.parse(importedResult.output) as PresentationImportProgress
  const project: PresentationProjectStatus = {
    projectId: artifact.projectId,
    title: 'QA source',
    status: 'compiled',
    slideCount: 2,
    slides,
    history: [],
    production: {
      projectId: artifact.projectId,
      requestId: artifact.requestId,
      planRevision: 1,
      status: 'compiled',
      total: 2,
      compiledCount: 2,
      pages: slides.map((slide) => ({ ...slide, state: 'compiled', attempt: 1 })),
    },
  }
  const inspectionHost = vi.fn(async (id: string) => ({
    slideId: id,
    slideWidth: 960,
    slideHeight: 540,
    shapes: [],
    shapesTruncated: false,
    overflows: [],
    overlaps: [],
    overlapsTruncated: false,
    screenshot: { mime: 'image/png' as const, base64: png },
  }))
  const skill = createPresentationQaSkill({
    available: () => true,
    artifact: () => artifact,
    documentId: () => binding.documentId(),
    readReceipt: (key) => binding.readReceipt(key),
    readQa: (key) => binding.readQa(key),
    writeQa: (key, value) => binding.writeQa(key, value),
    vfs: new InMemoryVfs(),
    inspectPage: inspectionHost,
    readQaAttempts: (key) => binding.readQaAttempts(key),
    writeQaAttempt: (key, value) => binding.writeQaAttempt(key, value),
  })
  const key = `production/${artifact.projectId}/${artifact.requestId}`
  const read = () => binding.readQa(key)!
  const capture = (pageId: string) =>
    skill.executeTool({
      id: `capture-${pageId}`,
      name: 'capture_presentation_page_qa',
      input: { page_id: pageId },
    })
  const review = (pageId: string) =>
    skill.executeTool({
      id: `review-${pageId}`,
      name: 'record_presentation_page_review',
      input: {
        page_id: pageId,
        screenshot_digest: read().pages.find((page) => page.pageId === pageId)!.screenshotDigest,
        outcome: 'pass',
        notes:
          'Synthetic screenshot fixture only; this does not verify professional content or a real host.',
      },
    })
  const workflow = () =>
    presentationWorkflowSummary(project, imported, read(), undefined, undefined, {
      attempts: binding.readQaAttempts(key),
    })!
  const rows = () => workflow().timeline.filter((event) => event.scope === 'saved_page_qa')
  return {
    capture,
    artifact,
    key,
    close: (
      expected: import('../src/skills/powerpoint/presentation-qa-attempts.js').PresentationQaAttempt,
    ) => skill.closeAttempt(expected),
    inspectionHost,
    attempts: () => binding.readQaAttempts(key),
    review,
    read,
    rows,
    workflow,
    binding: () => binding,
    reopen: () => {
      binding = createPresentationDocumentBinding(settings)
    },
    host,
  }
}
it('replays actual partial QA, first scoped invalidation and fresh recapture through durable settings', async () => {
  const f = await fixture()
  const base = Date.now()
  vi.useFakeTimers({ toFake: ['Date'] })
  vi.setSystemTime(base)
  expect((await f.capture('one')).isError).not.toBe(true)
  vi.setSystemTime(base + 1000)
  expect((await f.review('one')).isError).not.toBe(true)
  const first = structuredClone(f.read().pages[0]!)
  expect(f.rows()).toHaveLength(1)
  expect(f.rows()[0]!.records?.map((event) => event.at)).toEqual([
    first.capturedAt,
    first.visual.reviewedAt,
  ])
  expect(f.workflow().stages.find((stage) => stage.name === '交付核验')!.status).not.toBe(
    'recorded',
  )
  expect(f.workflow().stages.find((stage) => stage.name === '页面审查')!.status).toBe('attention')
  expect((await f.capture('two')).isError).not.toBe(true)
  vi.setSystemTime(base + 2000)
  await f.binding().invalidateQa(['host-1'])
  f.reopen()
  const stale = f.read().pages.find((page) => page.pageId === 'one')!
  expect(stale).toMatchObject({
    recheckRequired: true,
    invalidatedAt: new Date(base + 2000).toISOString(),
    screenshotDigest: first.screenshotDigest,
    visual: first.visual,
  })
  expect(f.read().pages.find((page) => page.pageId === 'two')!.recheckRequired).toBeUndefined()
  const staleRow = f.rows().find((row) => row.type === 'qa.evidence.invalidated')!
  expect(staleRow.records?.map((event) => event.at)).toEqual([
    first.capturedAt,
    first.visual.reviewedAt,
    stale.invalidatedAt,
  ])
  expect(f.rows()).toEqual(f.rows())
  vi.setSystemTime(base + 3000)
  await f.binding().invalidateQa(['host-1'])
  expect(f.read().pages.find((page) => page.pageId === 'one')!.invalidatedAt).toBe(
    stale.invalidatedAt,
  )
  expect(await f.review('one')).toMatchObject({
    isError: true,
    output: 'presentation_qa_capture_required',
  })
  expect((await f.capture('one')).isError).not.toBe(true)
  f.reopen()
  const fresh = f.read().pages.find((page) => page.pageId === 'one')!
  expect(fresh.recheckRequired).toBeUndefined()
  expect(fresh.invalidatedAt).toBeUndefined()
  expect(fresh.visual.status).toBe('needs_review')
  expect(f.rows().some((row) => row.id === staleRow.id)).toBe(false)
  expect(f.host).toEqual(['original', 'host-1', 'host-2'])
})

it('reopens actual waiting screenshot attempts and retries only when explicitly requested', async () => {
  const f = await fixture()
  f.inspectionHost.mockRejectedValueOnce(
    Object.assign(new Error('private host information'), { code: 'office_screenshot_unavailable' }),
  )
  const waiting = await f.capture('one')
  expect(JSON.parse(waiting.output).status).toBe('waiting_screenshot')
  const attempt = structuredClone(f.attempts()[0]!)
  expect(attempt.status).toBe('waiting')
  expect(attempt.errorCode).toBe('screenshot_unavailable')
  expect(f.read()).toBeUndefined()
  f.reopen()
  expect(f.attempts()).toEqual([attempt])
  expect(f.inspectionHost).toHaveBeenCalledTimes(1)
  const timeline = f.workflow().timeline.filter((row) => row.scope === 'page_qa_attempt')
  expect(timeline).toHaveLength(1)
  expect(timeline[0].records!.map((item) => item.at)).toEqual([
    attempt.startedAt,
    attempt.finishedAt,
  ])
  expect(JSON.stringify(timeline)).not.toContain('private host information')
  expect((await f.capture('two')).isError).not.toBe(true)
  expect((await f.capture('one')).isError).not.toBe(true)
  f.reopen()
  expect(f.attempts().map((item) => item.status)).toEqual(['waiting', 'recorded', 'recorded'])
  expect(f.attempts()[0]).toEqual(attempt)
  expect(f.read().pages.every((page) => page.visual.status === 'needs_review')).toBe(true)
  expect(f.workflow().stages.find((stage) => stage.name === '页面审查')).toMatchObject({
    status: 'working',
  })
  expect(f.host).toEqual(['original', 'host-1', 'host-2'])
  expect(f.inspectionHost).toHaveBeenCalledTimes(3)
})

it('explicitly closes one of 64 actual durable unresolved attempts then accepts a fresh screenshot without replaying old ones', async () => {
  const f = await fixture()
  const digest = f.binding().readReceipt(f.key)!.checkpoint!.artifactDigest
  for (let index = 0; index < 64; index++)
    await f.binding().writeQaAttempt(f.key, {
      version: 1,
      id: `12345678-1234-4234-8234-${index.toString(16).padStart(12, '0')}`,
      source: 'production',
      documentId: f.artifact.documentId,
      projectId: f.artifact.projectId,
      requestId: f.artifact.requestId,
      artifactDigest: digest,
      pageId: 'one',
      hostSlideId: 'host-1',
      startedAt: '2026-09-29T00:00:00.000Z',
      status: 'started',
    })
  f.reopen()
  const unresolved = structuredClone(f.attempts())
  expect((await f.capture('two')).output).toBe('presentation_qa_attempt_history_full')
  expect(f.inspectionHost).not.toHaveBeenCalled()
  const closed = await f.close(unresolved[0])
  expect(closed).toMatchObject({
    status: 'closed',
    errorCode: 'explicitly_closed',
    id: unresolved[0].id,
  })
  f.reopen()
  expect(await f.close(unresolved[0])).toEqual(closed)
  expect(f.attempts()).toHaveLength(64)
  expect(f.attempts().filter((item) => item.status === 'started')).toEqual(unresolved.slice(1))
  expect(f.inspectionHost).not.toHaveBeenCalled()
  expect(f.workflow().timeline.filter((row) => row.type === 'qa.attempt.closed')).toHaveLength(1)
  expect((await f.capture('two')).isError).not.toBe(true)
  f.reopen()
  expect(f.attempts()).toHaveLength(64)
  expect(f.attempts().filter((item) => item.status === 'started')).toEqual(unresolved.slice(1))
  expect(f.attempts().filter((item) => item.status === 'recorded')).toHaveLength(1)
  expect(f.inspectionHost).toHaveBeenCalledTimes(1)
  expect(f.host).toEqual(['original', 'host-1', 'host-2'])
  expect(f.read().pages).toHaveLength(1)
  expect(f.read().pages[0].visual.status).toBe('needs_review')
})

it('exports reopened waiting and explicit-close records only with requested local diagnostic context', async () => {
  const f = await fixture()
  f.inspectionHost.mockRejectedValueOnce(
    Object.assign(new Error('private host error'), { code: 'Timeout' }),
  )
  expect(JSON.parse((await f.capture('one')).output).status).toBe('waiting_screenshot')
  const { finishedAt: _finish, errorCode: _error, ...pending } = f.attempts()[0]
  const started = {
    ...pending,
    status: 'started' as const,
    id: '12345678-1234-4234-8234-123456789abc',
  }
  await f.binding().writeQaAttempt(f.key, started)
  const closed = await f.close(started)
  expect((await f.capture('two')).isError).not.toBe(true)
  f.reopen()
  const history = structuredClone(f.attempts())
  const provider = vi.fn(() => f.attempts())
  const send = vi.fn()
  const diagnostics = createOfficeDiagnostics({
    host: 'powerpoint',
    build: 'integration-build',
    remoteEnabled: true,
    remoteSamplePercent: 100,
    send,
  })
  diagnostics.setTool('capture_presentation_page_qa', {
    project_id: f.artifact.projectId,
    page_id: 'one',
  })
  diagnostics.record({ phase: 'verify', errorCode: 'office_read_failed' })
  const standard = JSON.parse(diagnostics.exportJson({ screenshotAttempts: provider }))
  expect(provider).not.toHaveBeenCalled()
  expect(standard).not.toHaveProperty('local_presentation_qa_attempts')
  const local = JSON.parse(
    diagnostics.exportJson({ includeLocalContext: true, screenshotAttempts: provider }),
  )
  expect(local.local_presentation_qa_attempts).toMatchObject({
    scope: 'retained_visible_presentation_task',
    status: 'available',
    record_count: 3,
    unresolved_count: 0,
  })
  expect(local.local_presentation_qa_attempts.attempts).toEqual(history)
  expect(
    local.local_presentation_qa_attempts.attempts.find(
      (item: { id: string }) => item.id === closed.id,
    ),
  ).toEqual(closed)
  expect(JSON.stringify(local)).not.toContain('private host error')
  expect(send).toHaveBeenCalledTimes(1)
  expect(JSON.stringify(send.mock.calls)).not.toContain('local_presentation_qa_attempts')
  diagnostics.clear()
  const cleared = JSON.parse(
    diagnostics.exportJson({ includeLocalContext: true, screenshotAttempts: provider }),
  )
  expect(cleared.events).toEqual([])
  expect(cleared.local_presentation_qa_attempts.attempts).toEqual(history)
  expect(f.inspectionHost).toHaveBeenCalledTimes(2)
  expect(f.host).toEqual(['original', 'host-1', 'host-2'])
})
