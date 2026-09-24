import { expect, it } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
import type { PresentationProjectStatus } from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationImportProgress } from '../src/skills/powerpoint/presentation-page-delivery.js'
import type { PresentationQaRecord } from '../src/skills/powerpoint/presentation-qa.js'
import { deliveryReportFixture } from './presentation-delivery-fixture.js'

const plan = benchmarkPlan()
const project: PresentationProjectStatus = {
  projectId: plan.projectId, title: plan.title, status: 'planned',
  plan: { revision: 1, value: plan }, slideCount: plan.slides.length,
  slides: plan.slides.map(({ id, title }) => ({ id, title })), history: [],
}
const production: NonNullable<PresentationProjectStatus['production']> = {
  projectId: plan.projectId, requestId: 'run1', planRevision: 1,
  status: 'compiled', total: plan.slides.length, compiledCount: plan.slides.length,
  pages: plan.slides.map(({ id, title }) => ({ id, title, state: 'compiled', attempt: 1 })),
}
const imported: PresentationImportProgress = {
  source: 'production', projectId: plan.projectId, requestId: production.requestId,
  total: production.total, completed: production.total, status: 'complete',
  pages: production.pages.map(({ id, title }) => ({ id, title, state: 'complete' })),
}
const qa: PresentationQaRecord = {
  version: 1, source: 'production', projectId: plan.projectId, requestId: production.requestId,
  documentId: 'doc', artifactDigest: 'a'.repeat(64),
  pages: production.pages.map(({ id, title }) => ({
    pageId: id, title, hostSlideId: `host-${id}`, capturedAt: '2026-09-24T00:00:00.000Z',
    screenshotDigest: 'b'.repeat(64), screenshotBytes: 128,
    structure: { status: 'passed', shapeCount: 1, overflowCount: 0, overlapCount: 0, shapesTruncated: false, overlapsTruncated: false },
    visual: { status: 'pass', reviewer: 'agent', reviewedAt: '2026-09-24T00:01:00.000Z' },
  })),
}

it('walks the saved plan through production, import and QA without claiming delivery', () => {
  expect(presentationWorkflowSummary(project, undefined, undefined)?.nextAction).toContain('启动逐页生产')
  const partial = { ...production, status: 'partial' as const, compiledCount: production.total - 1,
    pages: production.pages.map((page, index) => index === 0 ? { ...page, state: 'failed' as const } : page) }
  expect(presentationWorkflowSummary({ ...project, production: partial }, undefined, undefined)?.nextAction).toContain('修复失败页')
  expect(presentationWorkflowSummary({ ...project, production: partial }, undefined, undefined)?.nextTool).toBe('start_job')
  expect(presentationWorkflowSummary({ ...project, production: partial, jobsUnavailable: true }, undefined, undefined)?.nextTool).toBe('run_pages')
  expect(presentationWorkflowSummary({ ...project, production: partial, productionJob: { state: 'paused' } as NonNullable<PresentationProjectStatus['productionJob']> }, undefined, undefined)?.nextTool).toBe('resume_job')
  expect(presentationWorkflowSummary({ ...project, production: partial }, undefined, undefined)?.pages[0]?.nextAction).toContain('重试此页')
  expect(presentationWorkflowSummary({ ...project, production: partial }, undefined, undefined)?.attention).toEqual([
    expect.objectContaining({ id: 'failed-pages', text: expect.stringContaining('1 页编译失败') }),
  ])
  expect(presentationWorkflowSummary({ ...project, production }, undefined, undefined)?.nextAction).toContain('逐页导入')
  expect(presentationWorkflowSummary({ ...project, production }, undefined, undefined)?.nextTool).toBe('prepare_import')
  expect(presentationWorkflowSummary({ ...project, production: { ...production, revision: { parentRequestId: 'parent', pageId: production.pages[0]!.id, parentInputDigest: 'a'.repeat(64) } } }, undefined, undefined)?.nextAction).toContain('单页修订')
  expect(presentationWorkflowSummary({ ...project, production }, imported, undefined)?.nextAction).toContain('采集页面截图')
  const complete = presentationWorkflowSummary({ ...project, production }, imported, qa)!
  expect(complete.nextAction).toContain('读取当前任务的内容证据')
  expect(complete.nextTool).toBe('read_report')
  expect(complete.stages.at(-1)?.detail).toContain('尚不能')
  expect(complete.pages.every((page) => page.qa === '历史结构与视觉通过')).toBe(true)
})

it('uses only the selected request report and keeps its open issues visible', async () => {
  const report = await deliveryReportFixture()
  const selected = { ...project, production: { ...production, requestId: report.requestId } }
  const exactImport = { ...imported, requestId: report.requestId }
  const exactQa = { ...qa, requestId: report.requestId }
  expect(presentationWorkflowSummary(selected, exactImport, exactQa, report)?.nextAction).toContain('待处理问题')
  expect(presentationWorkflowSummary(selected, exactImport, exactQa, report)?.attention.map((item) => item.id)).toContain('content-issues')
  expect(presentationWorkflowSummary(selected, exactImport, exactQa, { ...report, requestId: 'old' })?.nextAction).toContain('读取当前任务')
})

it('refuses stale import and QA records and requests recheck after a page edit', () => {
  const selected = { ...project, production }
  const staleImport = { ...imported, pages: imported.pages.map((page, index) => index === 0 ? { ...page, id: 'old' } : page) }
  expect(presentationWorkflowSummary(selected, staleImport, qa)?.nextAction).toContain('逐页导入')
  expect(presentationWorkflowSummary(selected, { ...imported, requestId: 'old' }, qa)?.nextAction).toContain('逐页导入')
  const staleQa = { ...qa, requestId: 'old' }
  expect(presentationWorkflowSummary(selected, imported, staleQa)?.nextAction).toContain('采集页面截图')
  const recheck = { ...qa, pages: qa.pages.map((page, index) => index === 0 ? { ...page, recheckRequired: true as const } : page) }
  expect(presentationWorkflowSummary(selected, imported, recheck)?.nextAction).toContain('重审受影响页')
  expect(presentationWorkflowSummary(selected, imported, recheck)?.pages[0]?.nextAction).toBe('重审此页')
  const uncertain = { ...imported, status: 'uncertain' as const, completed: 0,
    pages: imported.pages.map((page, index) => index === 0 ? { ...page, state: 'uncertain' as const } : { ...page, state: 'pending' as const }) }
  expect(presentationWorkflowSummary(selected, uncertain, qa)?.pages[0]?.nextAction).toContain('检查宿主页')
  expect(presentationWorkflowSummary(selected, uncertain, recheck)?.attention.map((item) => item.id)).toEqual([
    'uncertain-import', 'qa-recheck',
  ])
})

it('rebuilds recovery events from saved records and isolates the selected request', () => {
  const event = { type: 'page.compiled' as const, pageId: production.pages[0]!.id,
    attempt: 1, sequence: 4, createdAt: '2026-09-24T00:00:00.000Z' }
  const job = { projectId: plan.projectId, requestId: production.requestId, events: [event] } as
    NonNullable<PresentationProjectStatus['productionJob']>
  const selected = { ...project, production, productionJob: job }
  const first = presentationWorkflowSummary(selected, imported, qa)!
  const replayed = presentationWorkflowSummary(selected, imported, qa)!
  expect(replayed.timeline).toEqual(first.timeline)
  expect(first.timeline.map((item) => item.id)).toEqual(['plan', 'production', 'job-4', 'import', 'qa'])
  expect(first.timeline[2]).toMatchObject({ at: event.createdAt, text: expect.stringContaining(event.pageId) })
  const stale = presentationWorkflowSummary({ ...selected, productionJob: { ...job, requestId: 'old' } },
    { ...imported, requestId: 'old' }, { ...qa, requestId: 'old' })!
  expect(stale.timeline.map((item) => item.id)).toEqual(['plan', 'production'])
})

it('replays durable plan revisions without inventing research approval', () => {
  const revisions = [1, 2].map((revision) => ({ revision,
    inputDigest: String(revision).repeat(64), createdAt: `2026-09-24T00:0${revision}:00.000Z` }))
  const summary = presentationWorkflowSummary({ ...project,
    plan: { ...project.plan!, revision: 2, revisions } }, undefined, undefined)!
  expect(summary.timeline.map((event) => event.id)).toEqual(['plan-1', 'plan-2'])
  expect(summary.timeline[1]).toMatchObject({ at: revisions[1]!.createdAt, text: '已保存计划第 2 版' })
  expect(JSON.stringify(summary.timeline)).not.toContain('批准')
})
