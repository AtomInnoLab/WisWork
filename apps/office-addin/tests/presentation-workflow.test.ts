import { expect, it } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { PRESENTATION_DOMAIN_PROFILES } from '@wiswork/pptx-engine/presentation-plan'
import {
  presentationWorkflowSummary,
  presentationProductionEventRows,
} from '../src/agent/presentation-workflow.js'
import type { PresentationProjectStatus } from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationImportProgress } from '../src/skills/powerpoint/presentation-page-delivery.js'
import type { PresentationQaRecord } from '../src/skills/powerpoint/presentation-qa.js'
import { deliveryReportFixture } from './presentation-delivery-fixture.js'

const plan = benchmarkPlan()
it('replays only explicit plan and style acceptance decisions with their saved identity and historical scope', () => {
  const acceptance = {
    decisionId: 'decision',
    planRevision: 1,
    planDigest: 'a'.repeat(64),
    styleDigest: 'b'.repeat(64),
    acceptedAt: '2026-09-29T00:00:00.000Z',
  }
  const saved: PresentationProjectStatus = {
    projectId: plan.projectId,
    title: plan.title,
    status: 'planned',
    slideCount: plan.slides.length,
    slides: plan.slides.map(({ id, title }) => ({ id, title })),
    history: [],
    plan: { revision: 2, value: plan },
    planAcceptance: {
      version: 1,
      projectId: plan.projectId,
      documentId: 'doc',
      records: [acceptance],
    },
  }
  const summary = presentationWorkflowSummary(saved, undefined, undefined)!
  const events = summary.timeline.filter((event) => event.scope === 'explicit_user_decision')
  expect(events.map((event) => event.type)).toEqual(['plan.approved', 'style.approved'])
  expect(events[0]).toMatchObject({
    at: acceptance.acceptedAt,
    text: expect.stringContaining('历史决定'),
  })
  expect(
    presentationWorkflowSummary(saved, undefined, undefined)!.timeline.filter(
      (event) => event.scope === 'explicit_user_decision',
    ),
  ).toEqual(events)
  expect(
    presentationWorkflowSummary(
      { ...saved, planAcceptanceUnavailable: true },
      undefined,
      undefined,
    )!.timeline.some((event) => event.scope === 'explicit_user_decision'),
  ).toBe(false)
})
it('folds durable source-excerpt research by frozen plan and keeps retry detail without claiming source truth', () => {
  const runs = [
    {
      id: 'one',
      sequence: 1,
      scope: 'source_excerpt_audit' as const,
      planRevision: 1,
      planDigest: 'a'.repeat(64),
      state: 'completed' as const,
      startedAt: '2026-09-29T00:00:00.000Z',
      finishedAt: '2026-09-29T00:01:00.000Z',
      sourceCount: 1,
      foundCount: 1,
    },
    {
      id: 'two',
      sequence: 3,
      scope: 'source_excerpt_audit' as const,
      planRevision: 1,
      planDigest: 'a'.repeat(64),
      state: 'failed' as const,
      startedAt: '2026-09-29T00:02:00.000Z',
      finishedAt: '2026-09-29T00:03:00.000Z',
      sourceCount: 1,
      error: 'aborted' as const,
    },
  ]
  const value = {
    ...project,
    sourceAuditHistory: {
      version: 1 as const,
      projectId: project.projectId,
      documentId: 'doc',
      revision: 4,
      runs,
    },
  }
  const workflow = presentationWorkflowSummary(value, undefined, undefined)!
  const rows = workflow.timeline.filter((event) => event.type?.startsWith('source_audit.'))
  expect(rows).toHaveLength(1)
  expect(rows[0]).toMatchObject({
    type: 'source_audit.failed',
    scope: 'source_excerpt_audit',
    records: [{ id: 'one' }, { id: 'two' }],
  })
  expect(rows[0]!.text).toContain('2 次')
  expect(rows[0]!.text).toContain('未核验')
  const next = {
    ...runs[0]!,
    id: 'three',
    sequence: 5,
    startedAt: '2026-09-29T00:04:00.000Z',
    finishedAt: '2026-09-29T00:05:00.000Z',
  }
  value.sourceAuditHistory.runs.push(next)
  value.sourceAuditHistory.revision = 6
  const after = presentationWorkflowSummary(value, undefined, undefined)!.timeline.filter((event) =>
    event.type?.startsWith('source_audit.'),
  )
  expect(after[0]!.id).toBe(rows[0]!.id)
  expect(after[0]!.text).toContain('3 次')
  expect(after[0]!.type).toBe('source_audit.completed')
})
it('an unfinished durable research read does not pretend to be a live background task', () => {
  const value = {
    ...project,
    sourceAuditHistory: {
      version: 1 as const,
      projectId: project.projectId,
      documentId: 'doc',
      revision: 1,
      runs: [
        {
          id: 'one',
          sequence: 1,
          scope: 'source_excerpt_audit' as const,
          planRevision: 1,
          planDigest: 'a'.repeat(64),
          state: 'running' as const,
          startedAt: '2026-09-29T00:00:00.000Z',
          sourceCount: 1,
        },
      ],
    },
  }
  const workflow = presentationWorkflowSummary(value, undefined, undefined)!
  expect(workflow.timeline.find((event) => event.type === 'source_audit.started')?.text).toContain(
    '不能证明仍在后台执行',
  )
  expect(workflow.attention.some((item) => item.id === 'source-audit-unfinished')).toBe(true)
})
const project: PresentationProjectStatus = {
  projectId: plan.projectId,
  title: plan.title,
  status: 'planned',
  plan: { revision: 1, value: plan },
  slideCount: plan.slides.length,
  slides: plan.slides.map(({ id, title }) => ({ id, title })),
  history: [],
}
const production: NonNullable<PresentationProjectStatus['production']> = {
  projectId: plan.projectId,
  requestId: 'run1',
  planRevision: 1,
  status: 'compiled',
  total: plan.slides.length,
  compiledCount: plan.slides.length,
  pages: plan.slides.map(({ id, title }) => ({ id, title, state: 'compiled', attempt: 1 })),
}
const imported: PresentationImportProgress = {
  source: 'production',
  projectId: plan.projectId,
  requestId: production.requestId,
  total: production.total,
  completed: production.total,
  status: 'complete',
  pages: production.pages.map(({ id, title }) => ({
    id,
    title,
    state: 'complete',
    slideId: `host-${id}`,
  })),
}
const qa: PresentationQaRecord = {
  version: 1,
  source: 'production',
  projectId: plan.projectId,
  requestId: production.requestId,
  documentId: 'doc',
  artifactDigest: 'a'.repeat(64),
  pages: production.pages.map(({ id, title }) => ({
    pageId: id,
    title,
    hostSlideId: `host-${id}`,
    capturedAt: '2026-09-24T00:00:00.000Z',
    screenshotDigest: 'b'.repeat(64),
    screenshotBytes: 128,
    structure: {
      status: 'passed',
      shapeCount: 1,
      overflowCount: 0,
      overlapCount: 0,
      shapesTruncated: false,
      overlapsTruncated: false,
    },
    visual: {
      status: 'pass',
      reviewer: 'agent',
      notes: '历史Agent视觉意见',
      reviewedAt: '2026-09-24T00:01:00.000Z',
    },
  })),
}

it('does not count duplicate QA pages or a report missing a production page as current evidence', async () => {
  const selected = { ...project, production }
  const duplicateQa = structuredClone(qa)
  duplicateQa.pages[1]!.pageId = duplicateQa.pages[0]!.pageId
  const qaSummary = presentationWorkflowSummary(selected, imported, duplicateQa)!
  expect(qaSummary.stages.find((stage) => stage.name === '页面审查')).toMatchObject({
    status: 'attention',
    detail: expect.stringContaining('无法与当前页任务匹配'),
  })
  expect(qaSummary.pages[1]!.qa).toBe('无当前任务 QA 记录')

  const movedImport = structuredClone(imported)
  movedImport.pages[0]!.slideId = 'host-replaced-page'
  const movedSummary = presentationWorkflowSummary(selected, movedImport, qa)!
  expect(movedSummary.stages.find((stage) => stage.name === '页面审查')?.status).toBe('attention')
  expect(movedSummary.pages[0]!.qa).toBe('无当前任务 QA 记录')
  expect(movedSummary.pages[1]!.qa).toBe('历史结构与视觉通过')
  expect(movedSummary.stages.find((stage) => stage.name === '页面审查')?.detail).toContain(
    '7/8 页通过',
  )

  const uncertainImport = structuredClone(imported)
  uncertainImport.status = 'uncertain'
  uncertainImport.completed--
  uncertainImport.pages[0]!.state = 'uncertain'
  const uncertainSummary = presentationWorkflowSummary(selected, uncertainImport, qa)!
  expect(uncertainSummary.stages.find((stage) => stage.name === '页面审查')?.status).toBe(
    'attention',
  )
  expect(uncertainSummary.pages[0]!.qa).toBe('无当前任务 QA 记录')
  expect(uncertainSummary.pages[1]!.qa).toBe('历史结构与视觉通过')

  const report = await deliveryReportFixture()
  const reportSelected = { ...project, production: { ...production, requestId: report.requestId } }
  const incompleteReport = structuredClone(report)
  incompleteReport.pages[1]!.pageId = incompleteReport.pages[0]!.pageId
  const reportSummary = presentationWorkflowSummary(
    reportSelected,
    { ...imported, requestId: report.requestId },
    { ...qa, requestId: report.requestId },
    incompleteReport,
  )!
  expect(reportSummary.stages.find((stage) => stage.name === '交付核验')).toMatchObject({
    status: 'attention',
    detail: expect.stringContaining('报告页与当前页任务不匹配'),
  })
  expect(reportSummary.pages[1]!.evidence).toBe('无当前任务内容报告')
})

it('shows document-bound source preparation without claiming source truth', () => {
  const withSource = structuredClone(project)
  const attachmentId = 'a'.repeat(64)
  withSource.plan!.value.sources[0]!.uri = `attachment:${attachmentId}`
  withSource.sourcePreparation = [
    { sourceId: withSource.plan!.value.sources[0]!.id, attachmentId, status: 'missing' },
  ]
  const missing = presentationWorkflowSummary(withSource, undefined, undefined)!
  expect(missing.stages[0]!.detail).toContain('0/1 份摘录已匹配原文')
  expect(missing.attention).toContainEqual(expect.objectContaining({ id: 'source-preparation' }))
  expect(missing.nextAction).toContain('补齐计划引用的资料')
  expect(missing.timeline).toContainEqual(expect.objectContaining({ id: 'source-preparation' }))
  withSource.sourcePreparation[0]!.status = 'excerpt_mismatch'
  const mismatch = presentationWorkflowSummary(withSource, undefined, undefined)!
  expect(mismatch.attention.find((item) => item.id === 'source-preparation')?.text).toContain(
    '摘录不在原文中',
  )
  expect(mismatch.nextAction).toContain('修订计划摘录')
  withSource.sourcePreparation[0]!.status = 'ready'
  const legacy = presentationWorkflowSummary(withSource, undefined, undefined)!
  expect(legacy.attention.find((item) => item.id === 'source-preparation')?.text).toContain(
    '旧版 PC 未核对摘录',
  )
  withSource.sourcePreparation[0]!.status = 'excerpt_matched'
  const ready = presentationWorkflowSummary(withSource, undefined, undefined)!
  expect(ready.attention.some((item) => item.id === 'source-preparation')).toBe(false)
  expect(ready.nextAction).toContain('启动逐页生产')
  expect(ready.stages[0]!.detail).toContain('真实性仍需核验')
  delete withSource.sourcePreparation
  withSource.sourcePreparationUnavailable = true
  expect(presentationWorkflowSummary(withSource, undefined, undefined)?.attention).toContainEqual(
    expect.objectContaining({ id: 'source-preparation-unavailable' }),
  )
  expect(presentationWorkflowSummary(withSource, undefined, undefined)?.nextAction).toContain(
    '资料状态',
  )
})

it('shows the chosen domain sections only for the matching plan revision', () => {
  const domainPlan = benchmarkPlan()
  domainPlan.domain = 'research'
  domainPlan.slides.forEach((slide, index) => {
    slide.domainSection = PRESENTATION_DOMAIN_PROFILES.research.sections[index % 5]
  })
  const planned = { ...project, plan: { revision: 2, value: domainPlan } }
  const current = presentationWorkflowSummary(planned, undefined, undefined)!
  expect(current.stages[1]!.detail).toContain('研究报告结构')
  expect(current.pages[0]!.section).toBe('研究问题')
  const oldTask = presentationWorkflowSummary({ ...planned, production }, undefined, undefined)!
  expect(oldTask.pages[0]!.section).toBeUndefined()
})

it('walks the saved plan through production, import and QA without claiming delivery', () => {
  expect(
    presentationWorkflowSummary(project, undefined, undefined)?.stages.map((stage) => stage.status),
  ).toEqual(['recorded', 'recorded', 'pending', 'pending', 'pending', 'pending'])
  expect(presentationWorkflowSummary(project, undefined, undefined)?.nextAction).toContain(
    '启动逐页生产',
  )
  const partial = {
    ...production,
    status: 'partial' as const,
    compiledCount: production.total - 1,
    pages: production.pages.map((page, index) =>
      index === 0 ? { ...page, state: 'failed' as const } : page,
    ),
  }
  expect(
    presentationWorkflowSummary({ ...project, production: partial }, undefined, undefined)
      ?.nextAction,
  ).toContain('修复失败页')
  expect(
    presentationWorkflowSummary({ ...project, production: partial }, undefined, undefined)
      ?.nextTool,
  ).toBe('start_job')
  expect(
    presentationWorkflowSummary(
      { ...project, production: partial, jobsUnavailable: true },
      undefined,
      undefined,
    )?.nextTool,
  ).toBe('run_pages')
  expect(
    presentationWorkflowSummary(
      {
        ...project,
        production: partial,
        productionJob: { state: 'paused' } as NonNullable<
          PresentationProjectStatus['productionJob']
        >,
      },
      undefined,
      undefined,
    )?.nextTool,
  ).toBe('resume_job')
  expect(
    presentationWorkflowSummary({ ...project, production: partial }, undefined, undefined)?.pages[0]
      ?.nextAction,
  ).toContain('重试此页')
  expect(
    presentationWorkflowSummary({ ...project, production: partial }, undefined, undefined)
      ?.attention,
  ).toEqual([
    expect.objectContaining({ id: 'failed-pages', text: expect.stringContaining('1 页编译失败') }),
  ])
  expect(
    presentationWorkflowSummary({ ...project, production }, undefined, undefined)?.nextAction,
  ).toContain('逐页导入')
  expect(
    presentationWorkflowSummary({ ...project, production }, undefined, undefined)?.nextTool,
  ).toBe('prepare_import')
  expect(
    presentationWorkflowSummary(
      {
        ...project,
        production: {
          ...production,
          revision: {
            parentRequestId: 'parent',
            pageId: production.pages[0]!.id,
            parentInputDigest: 'a'.repeat(64),
          },
        },
      },
      undefined,
      undefined,
    )?.nextAction,
  ).toContain('单页修订')
  expect(
    presentationWorkflowSummary({ ...project, production }, imported, undefined)?.nextAction,
  ).toContain('采集页面截图')
  const complete = presentationWorkflowSummary({ ...project, production }, imported, qa)!
  expect(complete.nextAction).toContain('读取当前任务的内容证据')
  expect(complete.nextTool).toBe('read_report')
  expect(complete.stages.at(-1)?.detail).toContain('尚不能')
  expect(complete.stages.map((stage) => stage.status)).toEqual([
    'recorded',
    'recorded',
    'recorded',
    'recorded',
    'recorded',
    'pending',
  ])
  expect(complete.pages.every((page) => page.qa === '历史结构与视觉通过')).toBe(true)
})
it('identifies a fallback visual review as pending PowerPoint host appearance', () => {
  const fallback = structuredClone(qa)
  fallback.pages[0]!.screenshotRenderer = 'libreoffice'
  const summary = presentationWorkflowSummary({ ...project, production }, imported, fallback)!
  expect(summary.pages[0]!.qa).toContain('备用预览')
  expect(summary.pages[0]!.nextAction).toContain('宿主外观')
  expect(summary.stages.find((stage) => stage.name === '页面审查')?.status).toBe('working')
  expect(summary.nextAction).toContain('PowerPoint')
  expect(summary.nextTool).toBeUndefined()
  expect(summary.stages.find((stage) => stage.name === '页面审查')?.detail).toContain(
    '1 页使用备用预览',
  )
  expect(summary.timeline.find((event) => event.scope === 'saved_page_qa')?.text).toContain(
    'LibreOffice 备用预览',
  )
  const mixed = structuredClone(fallback)
  mixed.pages[1]!.visual.status = 'needs_changes'
  const withFailure = presentationWorkflowSummary({ ...project, production }, imported, mixed)!
  expect(withFailure.stages.find((stage) => stage.name === '页面审查')?.status).toBe('attention')
  expect(withFailure.nextAction).toContain('处理未通过')
})

it('marks durable phase problems for attention without claiming delivery completion', async () => {
  const failed = {
    ...production,
    compiledCount: production.total - 1,
    pages: production.pages.map((page, index) =>
      index === 0 ? { ...page, state: 'failed' as const } : page,
    ),
  }
  expect(
    presentationWorkflowSummary({ ...project, production: failed }, undefined, undefined)?.stages[2]
      ?.status,
  ).toBe('attention')
  const uncertain = {
    ...imported,
    status: 'uncertain' as const,
    completed: imported.total - 1,
    pages: imported.pages.map((page, index) =>
      index === 0 ? { ...page, state: 'uncertain' as const } : page,
    ),
  }
  expect(
    presentationWorkflowSummary({ ...project, production }, uncertain, qa)?.stages[3]?.status,
  ).toBe('attention')
  const report = await deliveryReportFixture()
  expect(
    presentationWorkflowSummary({ ...project, production }, imported, qa, report)?.stages[5]
      ?.status,
  ).not.toBe('recorded')
})

it('uses only the selected request report and keeps its open issues visible', async () => {
  const report = await deliveryReportFixture()
  const selected = { ...project, production: { ...production, requestId: report.requestId } }
  const exactImport = { ...imported, requestId: report.requestId }
  const exactQa = { ...qa, requestId: report.requestId }
  expect(presentationWorkflowSummary(selected, exactImport, exactQa, report)?.nextAction).toContain(
    '待处理问题',
  )
  expect(
    presentationWorkflowSummary(selected, exactImport, exactQa, report)?.attention.map(
      (item) => item.id,
    ),
  ).toContain('content-issues')
  expect(
    presentationWorkflowSummary(selected, exactImport, exactQa, { ...report, requestId: 'old' })
      ?.nextAction,
  ).toContain('读取当前任务')
  const summary = presentationWorkflowSummary(selected, exactImport, exactQa, report)!
  const issuePage = report.pages.find((page) =>
    page.issues.some((issue) => issue.disposition.state === 'open'),
  )!
  expect(summary.pages.find((page) => page.id === issuePage.pageId)).toMatchObject({
    evidence: expect.stringContaining('证据问题待处理'),
    nextAction: expect.stringContaining('此页内容证据问题'),
  })
  expect(
    presentationWorkflowSummary(selected, exactImport, exactQa, {
      ...report,
      requestId: 'old',
    })?.pages.find((page) => page.id === issuePage.pageId)?.evidence,
  ).toContain('无当前任务')
})

it('replays saved content issue decisions without treating explanations as verification', async () => {
  const report = await deliveryReportFixture()
  const issue = report.pages.flatMap((page) => page.issues)[0]!
  const action = {
    actionId: 'decision1',
    issueId: issue.id,
    issueDigest: issue.digest,
    state: 'explained' as const,
    note: 'Needs source review',
    sequence: 1,
    createdAt: '2026-09-24T00:03:00.000Z',
  }
  const selected = { ...project, production: { ...production, requestId: report.requestId } }
  const withAction = {
    ...report,
    issueLedger: { ...report.issueLedger, revision: 1, actions: [action] },
  }
  const timeline = presentationWorkflowSummary(selected, undefined, undefined, withAction)!.timeline
  expect(timeline.find((event) => event.id === 'issue-decision1')).toMatchObject({
    at: action.createdAt,
    text: expect.stringContaining('解释'),
  })
  expect(timeline.find((event) => event.id === 'issue-decision1')?.text).toContain('未验证')
  const changed = {
    ...withAction,
    issueLedger: {
      ...withAction.issueLedger,
      actions: [{ ...action, issueDigest: 'f'.repeat(64) }],
    },
  }
  expect(
    presentationWorkflowSummary(selected, undefined, undefined, changed)?.timeline.find(
      (event) => event.id === 'issue-decision1',
    )?.text,
  ).toContain('证据已变化')
  expect(
    presentationWorkflowSummary(selected, undefined, undefined, {
      ...withAction,
      requestId: 'old',
    })?.timeline.some((event) => event.id === 'issue-decision1'),
  ).toBe(false)
})

it('replays scoped agent evidence judgments as historical research events', async () => {
  const report = await deliveryReportFixture()
  const selected = { ...project, production: { ...production, requestId: report.requestId } }
  const source = report.plan.sources[0]!
  const claim = report.plan.claims.find((item) => item.sourceIds.includes(source.id))!
  const page = report.plan.slides.find((item) => item.claimIds.includes(claim.id))!
  const review = {
    version: 1 as const,
    projectId: report.projectId,
    requestId: report.requestId,
    reviewId: 'evidence1',
    planRevision: report.planRevision,
    inputDigest: report.inputDigest,
    planDigest: report.planDigest,
    pageId: page.id,
    claimId: claim.id,
    sourceId: source.id,
    attachmentId: 'a'.repeat(64),
    offset: 0,
    maxChars: 100,
    evidenceDigest: 'b'.repeat(64),
    outcome: 'supported' as const,
    notes: 'Scoped judgment',
    reviewer: 'agent' as const,
    createdAt: '2026-09-24T00:04:00.000Z',
    checks: {
      support: 'agent_reviewed' as const,
      sourceAuthority: 'not_verified' as const,
      timeliness: 'not_verified' as const,
      host: 'not_checked' as const,
    },
  }
  const withReview = { ...report, reviews: [review] }
  const event = presentationWorkflowSummary(
    selected,
    undefined,
    undefined,
    withReview,
  )?.timeline.find((item) => item.id === 'evidence-evidence1')
  expect(event).toMatchObject({ at: review.createdAt, text: expect.stringContaining('Agent') })
  expect(event?.text).toContain('来源真实性仍需核验')
  expect(
    presentationWorkflowSummary(selected, undefined, undefined, {
      ...withReview,
      requestId: 'old',
    })?.timeline.some((item) => item.id === 'evidence-evidence1'),
  ).toBe(false)
})

it('refuses stale import and QA records and requests recheck after a page edit', () => {
  const selected = { ...project, production }
  const staleImport = {
    ...imported,
    pages: imported.pages.map((page, index) => (index === 0 ? { ...page, id: 'old' } : page)),
  }
  expect(presentationWorkflowSummary(selected, staleImport, qa)?.nextAction).toContain('逐页导入')
  expect(
    presentationWorkflowSummary(selected, { ...imported, requestId: 'old' }, qa)?.nextAction,
  ).toContain('逐页导入')
  const staleQa = { ...qa, requestId: 'old' }
  expect(presentationWorkflowSummary(selected, imported, staleQa)?.nextAction).toContain(
    '采集页面截图',
  )
  const recheck = {
    ...qa,
    pages: qa.pages.map((page, index) =>
      index === 0 ? { ...page, recheckRequired: true as const } : page,
    ),
  }
  expect(presentationWorkflowSummary(selected, imported, recheck)?.nextAction).toContain(
    '重审受影响页',
  )
  expect(presentationWorkflowSummary(selected, imported, recheck)?.pages[0]?.nextAction).toBe(
    '重审此页',
  )
  const uncertain = {
    ...imported,
    status: 'uncertain' as const,
    completed: 0,
    pages: imported.pages.map((page, index) =>
      index === 0
        ? { ...page, state: 'uncertain' as const }
        : { ...page, state: 'pending' as const },
    ),
  }
  expect(presentationWorkflowSummary(selected, uncertain, qa)?.pages[0]?.nextAction).toContain(
    '检查宿主页',
  )
  const timedUncertain = {
    ...uncertain,
    pages: uncertain.pages.map((page, index) =>
      index === 0 ? { ...page, startedAt: '2026-09-24T00:02:00.000Z' } : page,
    ),
  }
  expect(
    presentationWorkflowSummary(selected, timedUncertain, qa)?.timeline.find(
      (item) => item.scope === 'host_page_import' && item.type === 'host.import.uncertain',
    ),
  ).toMatchObject({
    at: '2026-09-24T00:02:00.000Z',
    text: expect.stringContaining('待核查'),
  })
  expect(
    presentationWorkflowSummary(selected, uncertain, recheck)?.attention.map((item) => item.id),
  ).toEqual(['uncertain-import', 'qa-recheck'])
})

it('rebuilds recovery events from saved records and isolates the selected request', () => {
  const event = {
    type: 'page.compiled' as const,
    pageId: production.pages[0]!.id,
    attempt: 1,
    sequence: 4,
    createdAt: '2026-09-24T00:00:00.000Z',
  }
  const job = {
    projectId: plan.projectId,
    requestId: production.requestId,
    events: [event],
  } as NonNullable<PresentationProjectStatus['productionJob']>
  const selected = { ...project, production, productionJob: job }
  const first = presentationWorkflowSummary(selected, imported, qa)!
  const replayed = presentationWorkflowSummary(selected, imported, qa)!
  expect(replayed.timeline).toEqual(first.timeline)
  expect(first.timeline.map((item) => item.id)).toEqual(
    expect.arrayContaining([
      'plan',
      'production',
      `job-page:${JSON.stringify([project.projectId, production.requestId, event.pageId])}`,
      'import',
      'qa',
    ]),
  )
  expect(first.timeline.filter((item) => item.scope === 'saved_page_qa')).toHaveLength(
    qa.pages.length,
  )
  expect(
    first.timeline.filter(
      (item) =>
        item.scope === 'saved_page_qa' &&
        item.records?.some((record) => record.type === 'qa.visual.recorded'),
    ),
  ).toHaveLength(qa.pages.length)
  expect(
    first.timeline.find(
      (item) =>
        item.id ===
        `job-page:${JSON.stringify([project.projectId, production.requestId, event.pageId])}`,
    ),
  ).toMatchObject({
    at: event.createdAt,
    text: expect.stringContaining(production.pages[0]!.title),
  })
  const withHistory = presentationWorkflowSummary(
    {
      ...selected,
      productionTasks: [
        {
          requestId: production.requestId,
          sequence: 2,
          planRevision: 1,
          status: 'compiled',
          compiledCount: production.total,
          total: production.total,
          lastEvent: { type: 'run.completed', createdAt: '2026-09-24T00:00:01.000Z' },
        },
        {
          requestId: 'older',
          sequence: 1,
          planRevision: 1,
          status: 'partial',
          compiledCount: 1,
          total: production.total,
          jobState: 'paused',
          lastEvent: { type: 'run.paused', createdAt: '2026-09-23T00:00:00.000Z' },
        },
      ],
    },
    imported,
    qa,
  )!
  expect(withHistory.timeline.find((item) => item.id === 'task-older-latest')).toMatchObject({
    text: expect.stringContaining('已暂停制作'),
    at: '2026-09-23T00:00:00.000Z',
  })
  expect(
    withHistory.timeline.some((item) => item.id === `task-${production.requestId}-latest`),
  ).toBe(false)
  expect(first.timeline.filter((item) => item.at).map((item) => item.at)).toEqual(
    first.timeline
      .filter((item) => item.at)
      .map((item) => item.at)
      .sort(),
  )
  const stale = presentationWorkflowSummary(
    { ...selected, productionJob: { ...job, requestId: 'old' } },
    { ...imported, requestId: 'old' },
    { ...qa, requestId: 'old' },
  )!
  expect(stale.timeline.map((item) => item.id)).toEqual(['plan', 'production'])
})

it('replays page import and historical QA times only for the current request', () => {
  const timed = {
    ...imported,
    pages: imported.pages.map((page, index) => ({
      ...page,
      slideId: `host-${page.id}`,
      completedAt: `2026-09-24T00:0${index}:00.000Z`,
    })),
  }
  const summary = presentationWorkflowSummary({ ...project, production }, timed, qa)!
  expect(summary.timeline.filter((item) => item.scope === 'host_page_import')).toHaveLength(
    imported.pages.length,
  )
  expect(
    summary.timeline.find(
      (item) => item.scope === 'host_page_import' && item.at === timed.pages[0]!.completedAt,
    ),
  ).toMatchObject({
    at: timed.pages[0]!.completedAt,
    text: expect.stringContaining('不代表视觉'),
  })
  expect(summary.timeline.find((item) => item.scope === 'saved_page_qa')?.text).toContain(
    '历史视觉复核',
  )
  expect(
    presentationWorkflowSummary(
      { ...project, production },
      { ...timed, requestId: 'old' },
      { ...qa, requestId: 'old' },
    )?.timeline.some(
      (event) => event.scope === 'host_page_import' || event.scope === 'saved_page_qa',
    ),
  ).toBe(false)
})

it('replays durable plan revisions without inventing research approval', () => {
  const revisions = [1, 2].map((revision) => ({
    revision,
    inputDigest: String(revision).repeat(64),
    createdAt: `2026-09-24T00:0${revision}:00.000Z`,
  }))
  const summary = presentationWorkflowSummary(
    { ...project, plan: { ...project.plan!, revision: 2, revisions } },
    undefined,
    undefined,
  )!
  expect(summary.timeline.map((event) => event.id)).toEqual(['plan-1', 'plan-2'])
  expect(summary.timeline[1]).toMatchObject({
    at: revisions[1]!.createdAt,
    text: expect.stringContaining('已保存计划第 2 版'),
  })
  expect(JSON.stringify(summary.timeline)).not.toContain('批准')
})

it('describes changed saved plan sections even when source counts stay the same', () => {
  const snapshot = {
    sourceCount: 1,
    claimCount: 2,
    slideCount: 3,
    sourcesDigest: 'a'.repeat(64),
    claimsDigest: 'b'.repeat(64),
    slidesDigest: 'c'.repeat(64),
    styleDigest: 'd'.repeat(64),
  }
  const revisions = [
    { revision: 1, inputDigest: '1'.repeat(64), createdAt: '2026-09-24T00:01:00.000Z', snapshot },
    {
      revision: 2,
      inputDigest: '2'.repeat(64),
      createdAt: '2026-09-24T00:02:00.000Z',
      snapshot: { ...snapshot, sourcesDigest: 'e'.repeat(64), styleDigest: 'f'.repeat(64) },
    },
  ]
  const summary = presentationWorkflowSummary(
    { ...project, plan: { ...project.plan!, revision: 2, revisions } },
    undefined,
    undefined,
  )!
  expect(summary.timeline.find((event) => event.id === 'plan-1')?.text).toContain('登记 1 份资料')
  expect(summary.timeline.find((event) => event.id === 'plan-2')?.text).toContain(
    '已登记资料、样式规范有变化',
  )
  expect(summary.timeline.find((event) => event.id === 'plan-2')?.text).toContain(
    '来源真实性仍需核验',
  )
  expect(summary.timeline.find((event) => event.id === 'plan-1')).toMatchObject({
    type: 'plan.proposed',
    scope: 'saved_plan',
  })
  expect(summary.timeline.find((event) => event.id === 'plan-2')).toMatchObject({
    type: 'plan.revised',
    scope: 'saved_plan',
  })
  expect(summary.timeline.find((event) => event.id === 'style-2')).toMatchObject({
    type: 'style.proposed',
    scope: 'saved_style',
  })
  expect(summary.timeline.some((event) => String(event.type).endsWith('.approved'))).toBe(false)
  const truncated = presentationWorkflowSummary(
    {
      ...project,
      plan: {
        ...project.plan!,
        revision: 5,
        revisions: [
          { ...revisions[0]!, revision: 3 },
          { ...revisions[1]!, revision: 5 },
        ],
      },
    },
    undefined,
    undefined,
  )!
  expect(truncated.timeline.find((event) => event.id === 'plan-5')?.text).not.toContain('有变化')
  expect(truncated.timeline.find((event) => event.id === 'plan-history')?.text).toContain(
    '不推断缺失版本',
  )
})

it('warns when the selected production task still uses an older saved plan', () => {
  const selected = { ...project, plan: { ...project.plan!, revision: 2 }, production }
  const summary = presentationWorkflowSummary(selected, undefined, undefined)!
  expect(summary.attention).toEqual(
    expect.arrayContaining([
      expect.objectContaining({
        id: 'plan-revision',
        text: expect.stringContaining('计划第 1 版'),
      }),
    ]),
  )
  expect(summary.attention.find((item) => item.id === 'plan-revision')?.text).toContain('第 2 版')
  expect(summary.nextTool).toBeUndefined()
  expect(summary.nextAction).toContain('选择继续旧任务或按新计划重新生产')
})

it('labels old QA as historical after a brand or style revision', () => {
  const snapshots = [
    {
      revision: 1,
      inputDigest: '1'.repeat(64),
      createdAt: '2026-09-24T00:01:00.000Z',
      snapshot: {
        sourceCount: 0,
        claimCount: 0,
        slideCount: plan.slides.length,
        sourcesDigest: 'a'.repeat(64),
        claimsDigest: 'b'.repeat(64),
        slidesDigest: 'c'.repeat(64),
        styleDigest: 'd'.repeat(64),
      },
    },
    {
      revision: 2,
      inputDigest: '2'.repeat(64),
      createdAt: '2026-09-24T00:02:00.000Z',
      snapshot: {
        sourceCount: 0,
        claimCount: 0,
        slideCount: plan.slides.length,
        sourcesDigest: 'a'.repeat(64),
        claimsDigest: 'b'.repeat(64),
        slidesDigest: 'c'.repeat(64),
        styleDigest: 'e'.repeat(64),
      },
    },
  ]
  const selected = {
    ...project,
    plan: { ...project.plan!, revision: 2, revisions: snapshots },
    production,
  }
  const summary = presentationWorkflowSummary(selected, imported, qa)!
  expect(summary.timeline.map((event) => event.id)).toEqual(
    expect.arrayContaining(['research-1', 'style-1', 'style-2']),
  )
  expect(summary.timeline.map((event) => event.id)).not.toContain('research-2')
  expect(summary.attention.map((item) => item.id)).toContain('style-revision')
  expect(summary.pages[0]?.qa).toBe('旧样式版本历史通过')
  expect(summary.pages[0]?.nextAction).toBe('先确认继续旧计划或选择新任务')
  expect(summary.nextTool).toBeUndefined()
})

function eventProject(): PresentationProjectStatus {
  const pageId = production.pages[0]!.id
  const events = [
    { sequence: 1, createdAt: '2026-09-24T00:00:01.000Z', type: 'run.started' as const },
    {
      sequence: 2,
      createdAt: '2026-09-24T00:00:02.000Z',
      type: 'page.started' as const,
      pageId,
      attempt: 1,
    },
    {
      sequence: 3,
      createdAt: '2026-09-24T00:00:03.000Z',
      type: 'page.failed' as const,
      pageId,
      attempt: 1,
      error: 'compile_failed' as const,
    },
    { sequence: 4, createdAt: '2026-09-24T00:00:04.000Z', type: 'run.paused' as const },
    {
      sequence: 5,
      createdAt: '2026-09-24T00:00:05.000Z',
      type: 'page.started' as const,
      pageId,
      attempt: 2,
    },
    {
      sequence: 6,
      createdAt: '2026-09-24T00:00:06.000Z',
      type: 'page.compiled' as const,
      pageId,
      attempt: 2,
    },
    { sequence: 7, createdAt: '2026-09-24T00:00:07.000Z', type: 'run.completed' as const },
  ]
  return {
    ...project,
    production,
    productionJob: {
      projectId: project.projectId,
      requestId: production.requestId,
      revision: 7,
      events,
    } as NonNullable<PresentationProjectStatus['productionJob']>,
  }
}

it('folds retries by stable request/page identity at the latest event position, preserving raw history', () => {
  const selected = eventProject()
  const original = structuredClone(selected)
  const grouped = presentationProductionEventRows(selected)!
  expect(grouped.rows).toHaveLength(4)
  const row = grouped.rows.find((row) => row.pageId)!
  expect(row.text).toContain(production.pages[0]!.title)
  expect(row.text).toContain('第 2 次')
  expect(row.text).toContain('未导入验收')
  expect(row.at).toBe('2026-09-24T00:00:06.000Z')
  expect(row.attempts).toHaveLength(4)
  expect(row.attempts?.map((event) => event.at)).toEqual(
    selected
      .productionJob!.events.filter((event) => 'pageId' in event)
      .map((event) => event.createdAt),
  )
  expect(grouped.rows.map((row) => row.at)).toEqual([
    '2026-09-24T00:00:01.000Z',
    '2026-09-24T00:00:04.000Z',
    '2026-09-24T00:00:06.000Z',
    '2026-09-24T00:00:07.000Z',
  ])
  expect(presentationProductionEventRows(structuredClone(selected))).toEqual(grouped)
  const beforeRetry = {
    ...selected,
    productionJob: {
      ...selected.productionJob!,
      events: selected.productionJob!.events.slice(0, 3),
      revision: 3,
    },
  }
  expect(presentationProductionEventRows(beforeRetry)!.rows.find((event) => event.pageId)!.id).toBe(
    row.id,
  )
  expect(selected).toEqual(original)
  const timeline = presentationWorkflowSummary(selected, undefined, undefined)!.timeline
  expect(timeline.filter((event) => event.id === row.id)).toHaveLength(1)
  expect(timeline.filter((event) => event.id.startsWith('job-')).length).toBe(4)
})

it('bounds visible event history and never invents failure totals when earlier attempts were truncated', () => {
  const selected = eventProject()
  selected.productionJob!.revision = 130
  selected.productionJob!.events = Array.from({ length: 30 }, (_, index) => ({
    sequence: index + 101,
    createdAt: `2026-09-24T00:00:${String(index).padStart(2, '0')}.000Z`,
    type: 'page.failed' as const,
    pageId: production.pages[0]!.id,
    attempt: index + 1,
    error: 'compile_failed' as const,
  }))
  const grouped = presentationProductionEventRows(selected)!
  expect(grouped.retainedEventCount).toBe(20)
  expect(grouped.truncated).toBe(true)
  expect(grouped.rows).toHaveLength(1)
  expect(grouped.rows[0]!.attempts).toHaveLength(20)
  expect(grouped.rows[0]!.text).toContain('第 30 次')
  expect(grouped.rows[0]!.text).not.toContain('失败 30 次')
  expect(
    presentationWorkflowSummary(selected, undefined, undefined)!.timeline.some((row) =>
      row.text.includes('更早历史已截断'),
    ),
  ).toBe(true)
})

it('isolates event grouping by project/request and uses only matching title metadata', () => {
  const selected = eventProject()
  expect(
    presentationProductionEventRows({
      ...selected,
      productionJob: { ...selected.productionJob!, requestId: 'other' },
    }),
  ).toBeUndefined()
  expect(
    presentationProductionEventRows({
      ...selected,
      productionJob: { ...selected.productionJob!, projectId: 'other' },
    }),
  ).toBeUndefined()
  const changedRequest = {
    ...selected,
    production: { ...production, requestId: 'other' },
    productionJob: { ...selected.productionJob!, requestId: 'other' },
  }
  expect(
    presentationProductionEventRows(changedRequest)!.rows.find((row) => row.pageId)!.id,
  ).not.toBe(presentationProductionEventRows(selected)!.rows.find((row) => row.pageId)!.id)
  const noPage = { ...selected, production: { ...production, pages: [] } }
  expect(presentationProductionEventRows(noPage)!.rows.find((row) => row.pageId)?.text).toContain(
    plan.slides[0]!.title,
  )
  const oldPlan = { ...noPage, plan: { ...project.plan!, revision: 2 } }
  expect(presentationProductionEventRows(oldPlan)!.rows.find((row) => row.pageId)?.text).toContain(
    production.pages[0]!.id,
  )
  expect(
    presentationProductionEventRows(oldPlan)!.rows.find((row) => row.pageId)?.text,
  ).not.toContain(plan.slides[0]!.title)
})

it('keeps distinct page groups through interleaved attempts and lifecycle checkpoints', () => {
  const selected = eventProject()
  const first = production.pages[0]!.id
  const second = production.pages[1]!.id
  selected.productionJob!.events = [
    {
      sequence: 1,
      createdAt: '2026-09-24T00:00:01.000Z',
      type: 'page.started',
      pageId: first,
      attempt: 1,
    },
    {
      sequence: 2,
      createdAt: '2026-09-24T00:00:02.000Z',
      type: 'page.started',
      pageId: second,
      attempt: 1,
    },
    {
      sequence: 3,
      createdAt: '2026-09-24T00:00:03.000Z',
      type: 'page.failed',
      pageId: first,
      attempt: 1,
      error: 'compile_failed',
    },
    { sequence: 4, createdAt: '2026-09-24T00:00:04.000Z', type: 'run.interrupted' },
    {
      sequence: 5,
      createdAt: '2026-09-24T00:00:05.000Z',
      type: 'page.compiled',
      pageId: second,
      attempt: 1,
    },
    {
      sequence: 6,
      createdAt: '2026-09-24T00:00:06.000Z',
      type: 'page.started',
      pageId: first,
      attempt: 2,
    },
  ]
  selected.productionJob!.revision = 6
  const grouped = presentationProductionEventRows(selected)!
  expect(grouped.rows.map((row) => row.pageId ?? 'run')).toEqual(['run', second, first])
  expect(grouped.rows.filter((row) => row.pageId)).toHaveLength(2)
  expect(grouped.rows.find((row) => row.pageId === first)?.text).toContain('开始编译')
  expect(grouped.rows.find((row) => row.pageId === first)?.text).not.toContain('编译完成')
  expect(grouped.rows.find((row) => row.pageId === first)?.attempts).toHaveLength(3)
  expect(grouped.rows.find((row) => row.pageId === second)?.attempts).toHaveLength(2)
})

it('retains bounded saved run failure reasons alongside page grouping', () => {
  const selected = eventProject()
  selected.productionJob!.events.push({
    sequence: 8,
    createdAt: '2026-09-24T00:00:08.000Z',
    type: 'run.failed',
    error: 'invalid_state',
  })
  selected.productionJob!.revision = 8
  expect(presentationProductionEventRows(selected)!.rows.at(-1)).toMatchObject({
    text: '页任务失败 · 任务状态异常',
    at: '2026-09-24T00:00:08.000Z',
  })
})
it('folds asset resolution by task and page without claiming verification or live work', () => {
  const assetHistory = {
    version: 1 as const,
    scope: 'production_asset_resolution' as const,
    projectId: project.projectId,
    documentId: 'doc',
    requestId: production.requestId,
    inputDigest: 'a'.repeat(64),
    planDigest: 'b'.repeat(64),
    revision: 2,
    events: [
      {
        type: 'asset.fetching' as const,
        pageId: production.pages[0]!.id,
        assetId: 'image',
        attempt: 1,
        sequence: 1,
        createdAt: '2026-09-29T00:00:00.000Z',
      },
      {
        type: 'asset.ready' as const,
        pageId: production.pages[0]!.id,
        assetId: 'image',
        attempt: 1,
        sequence: 2,
        createdAt: '2026-09-29T00:01:00.000Z',
      },
    ],
  }
  const summary = presentationWorkflowSummary(
    { ...project, production, assetHistory },
    undefined,
    undefined,
  )!
  const events = summary.timeline.filter((event) => event.scope === 'production_asset_resolution')
  expect(events).toHaveLength(1)
  expect(events[0]).toMatchObject({
    type: 'asset.ready',
    at: assetHistory.events[1]!.createdAt,
    text: expect.stringContaining('未核验'),
  })
  expect(
    presentationWorkflowSummary(
      { ...project, production, assetHistory: { ...assetHistory, requestId: 'foreign' } },
      undefined,
      undefined,
    )!.timeline.some((event) => event.scope === 'production_asset_resolution'),
  ).toBe(false)
  expect(
    presentationWorkflowSummary(
      {
        ...project,
        production,
        assetHistory: { ...assetHistory, revision: 1, events: assetHistory.events.slice(0, 1) },
      },
      undefined,
      undefined,
    )!.timeline.find((event) => event.type === 'asset.fetching')!.text,
  ).toContain('不能证明仍在执行')
})
