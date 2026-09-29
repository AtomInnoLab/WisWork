import { afterEach, expect, it, vi } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document.js'
import { createPresentationQaSkill } from '../src/skills/powerpoint/presentation-qa.js'
import { createPresentationProductionDeliverySkill } from '../src/skills/powerpoint/presentation-page-delivery.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
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
  const skill = createPresentationQaSkill({
    available: () => true,
    artifact: () => artifact,
    documentId: () => binding.documentId(),
    readReceipt: (key) => binding.readReceipt(key),
    readQa: (key) => binding.readQa(key),
    writeQa: (key, value) => binding.writeQa(key, value),
    vfs: new InMemoryVfs(),
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
  const workflow = () => presentationWorkflowSummary(project, imported, read())!
  const rows = () => workflow().timeline.filter((event) => event.scope === 'saved_page_qa')
  return {
    capture,
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
