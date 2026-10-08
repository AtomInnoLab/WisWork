import { afterEach, expect, it, vi } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document.js'
import { createPresentationProductionDeliverySkill } from '../src/skills/powerpoint/presentation-page-delivery.js'
import type { PresentationImportProgress } from '../src/skills/powerpoint/presentation-page-delivery.js'
import type {
  CompiledPresentationArtifact,
  PresentationDeliveryOptions,
} from '../src/skills/powerpoint/presentation-delivery.js'
import type { PresentationProjectStatus } from '../src/skills/powerpoint/presentation-project.js'

afterEach(() => vi.restoreAllMocks())
async function fixture() {
  const compiled = await compilePresentationDeck({
    version: 1,
    id: 'import-timing',
    title: 'Native source',
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
        id: 'page1',
        title: 'One native page',
        claimIds: [],
        elements: [
          {
            kind: 'text',
            id: 'title',
            x: 1,
            y: 1,
            w: 10,
            h: 1,
            text: 'One native page',
            fontSize: 24,
          },
        ],
      },
    ],
  })
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
    projectId: 'project',
    requestId: 'run',
    planRevision: 1,
    pptxBase64: '',
    pagePptxBase64: [Buffer.from(compiled.bytes).toString('base64')],
    slideCount: 1,
    pages: [{ id: 'page1', title: 'One native page', sourceSlideId: '256#' }],
  }
  const project: PresentationProjectStatus = {
    projectId: 'project',
    title: 'Native source',
    status: 'compiled',
    slideCount: 1,
    slides: [{ id: 'page1', title: 'One native page' }],
    history: [],
    production: {
      projectId: 'project',
      requestId: 'run',
      planRevision: 1,
      status: 'compiled',
      total: 1,
      compiledCount: 1,
      pages: [{ id: 'page1', title: 'One native page', state: 'compiled', attempt: 1 }],
    },
  }
  const host = ['original']
  const proposals = createStructuredProposalController()
  const adapter = {
    available: () => true,
    snapshot: async () => ({ slideIds: [...host], fingerprint: JSON.stringify(host) }),
    insert: vi.fn(),
    insertPage: vi.fn(async () => {
      host.push('new-slide')
      return { slideIds: ['new-slide'] }
    }),
    verify: async () => true,
    exportPage: async () => artifact.pagePptxBase64![0]!,
  }
  let failCompletion = false
  const options: PresentationDeliveryOptions = {
    adapter,
    proposals,
    available: () => true,
    artifact: () => artifact,
    documentId: () => binding.documentId(),
    readReceipt: (key) => binding.readReceipt(key),
    writeReceipt: async (key, value) => {
      if (failCompletion && value?.state === 'complete') {
        failCompletion = false
        throw new Error('receipt_write_failed')
      }
      await binding.writeReceipt(key, value)
    },
  }
  const skill = () => createPresentationProductionDeliverySkill(options)
  const call = {
    id: 'import',
    name: 'import_presentation_production',
    input: { project_id: 'project' },
  }
  const confirm = async () => {
    expect((await skill().executeTool(call)).isError).not.toBe(true)
    await proposals.confirm(proposals.pending()!.id)
  }
  const read = async () => {
    const result = await skill().executeTool({
      id: 'status',
      name: 'read_presentation_production_import_status',
      input: {},
    })
    expect(result.isError).not.toBe(true)
    return JSON.parse(result.output) as PresentationImportProgress
  }
  const rows = (progress: PresentationImportProgress) =>
    presentationWorkflowSummary(project, progress, undefined)!.timeline.filter(
      (row) => row.scope === 'host_page_import',
    )
  return {
    options,
    project,
    adapter,
    host,
    binding: () => binding,
    artifact,
    confirm,
    read,
    rows,
    skill,
    call,
    reopen: () => {
      binding = createPresentationDocumentBinding(settings)
    },
    loseCompletion: () => {
      failCompletion = true
    },
  }
}
it('reopens actual native-page receipts with original start and record times and unchanged QA boundary', async () => {
  const f = await fixture()
  let now = Date.parse('2026-09-29T01:00:00.000Z')
  vi.spyOn(Date, 'now').mockImplementation(() => now)
  f.adapter.insertPage.mockImplementationOnce(async () => {
    f.host.push('new-slide')
    now += 1500
    return { slideIds: ['new-slide'] }
  })
  await f.confirm()
  f.reopen()
  const progress = await f.read()
  expect(progress.pages[0]).toMatchObject({
    startedAt: '2026-09-29T01:00:00.000Z',
    completedAt: '2026-09-29T01:00:01.500Z',
  })
  const rows = f.rows(progress)
  expect(rows).toHaveLength(1)
  expect(rows[0]).toMatchObject({
    type: 'host.import.recorded',
    records: [{ at: progress.pages[0]!.startedAt }, { at: progress.pages[0]!.completedAt }],
  })
  expect(f.rows(await f.read())).toEqual(rows)
  const summary = presentationWorkflowSummary(f.project, progress, undefined)!
  expect(summary.pages[0]!.qa).not.toContain('通过')
  expect((await f.skill().executeTool(f.call)).output).toContain('already_imported')
  expect(f.adapter.insertPage).toHaveBeenCalledOnce()
  expect(f.rows({ ...progress, requestId: 'other-run' })).toEqual([])
})
it('keeps original uncertain start through exact-package read-only reconciliation without replay', async () => {
  const f = await fixture()
  let now = Date.parse('2026-09-29T02:00:00.000Z')
  vi.spyOn(Date, 'now').mockImplementation(() => now)
  f.loseCompletion()
  await expect(f.confirm()).rejects.toThrow('receipt_write_failed')
  f.reopen()
  const pending = await f.read()
  expect(pending.status).toBe('uncertain')
  expect(f.rows(pending)[0]).toMatchObject({
    type: 'host.import.uncertain',
    records: [{ at: '2026-09-29T02:00:00.000Z' }],
  })
  now += 60_000
  const result = await f
    .skill()
    .executeTool({ id: 'reconcile', name: 'reconcile_presentation_production_import', input: {} })
  expect(result.isError).not.toBe(true)
  f.reopen()
  const recorded = await f.read()
  expect(recorded.pages[0]).toMatchObject({
    startedAt: pending.pages[0]!.startedAt,
    completedAt: '2026-09-29T02:01:00.000Z',
  })
  expect(f.rows(recorded)[0]!.records?.map((record) => record.at)).toEqual([
    pending.pages[0]!.startedAt,
    recorded.pages[0]!.completedAt,
  ])
  expect(f.adapter.insertPage).toHaveBeenCalledOnce()
  expect(f.host).toEqual(['original', 'new-slide'])
  const legacy = structuredClone(f.binding().readReceipt('production/project/run')!)
  delete legacy.checkpoint!.completed[0]!.startedAt
  await f.binding().writeReceipt('production/project/run', legacy)
  f.reopen()
  const old = await f.read()
  expect(old.pages[0]!.startedAt).toBeUndefined()
  expect(f.rows(old)[0]!.records).toHaveLength(1)
  expect(f.rows(old)[0]!.records![0]!.at).toBe(recorded.pages[0]!.completedAt)
})
