import { expect, it } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
import type { PresentationProjectStatus } from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationImportProgress } from '../src/skills/powerpoint/presentation-page-delivery.js'
import type { PresentationQaRecord } from '../src/skills/powerpoint/presentation-qa.js'
import { deliveryReportFixture } from './presentation-delivery-fixture.js'

const plan = benchmarkPlan()
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
  pages: production.pages.map(({ id, title }) => ({ id, title, state: 'complete' })),
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
    visual: { status: 'pass', reviewer: 'agent', reviewedAt: '2026-09-24T00:01:00.000Z' },
  })),
}

it('walks the saved plan through production, import and QA without claiming delivery', () => {
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
  expect(complete.pages.every((page) => page.qa === '历史结构与视觉通过')).toBe(true)
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
      (item) => item.id === `import-uncertain-${uncertain.pages[0]!.id}`,
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
    expect.arrayContaining(['plan', 'production', 'job-4', 'import', 'qa']),
  )
  expect(first.timeline.filter((item) => item.id.startsWith('capture-'))).toHaveLength(
    qa.pages.length,
  )
  expect(first.timeline.filter((item) => item.id.startsWith('review-'))).toHaveLength(
    qa.pages.length,
  )
  expect(first.timeline.find((item) => item.id === 'job-4')).toMatchObject({
    at: event.createdAt,
    text: expect.stringContaining(event.pageId),
  })
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
      completedAt: `2026-09-24T00:0${index}:00.000Z`,
    })),
  }
  const summary = presentationWorkflowSummary({ ...project, production }, timed, qa)!
  expect(summary.timeline.filter((item) => item.id.startsWith('import-'))).toHaveLength(
    imported.pages.length,
  )
  expect(
    summary.timeline.find((item) => item.id === `import-${imported.pages[0]!.id}`),
  ).toMatchObject({
    at: timed.pages[0]!.completedAt,
    text: expect.stringContaining('尚未完成视觉验收'),
  })
  expect(
    summary.timeline.find((item) => item.id === `review-${qa.pages[0]!.pageId}`)?.text,
  ).toContain('历史视觉复核')
  expect(
    presentationWorkflowSummary(
      { ...project, production },
      { ...timed, requestId: 'old' },
      { ...qa, requestId: 'old' },
    )?.timeline.some((event) => /^(import-|capture-|review-)/.test(event.id)),
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
    { revision: 1, inputDigest: '1'.repeat(64), createdAt: '2026-09-24T00:01:00.000Z', snapshot: {
      sourceCount: 0, claimCount: 0, slideCount: plan.slides.length,
      sourcesDigest: 'a'.repeat(64), claimsDigest: 'b'.repeat(64), slidesDigest: 'c'.repeat(64), styleDigest: 'd'.repeat(64),
    } },
    { revision: 2, inputDigest: '2'.repeat(64), createdAt: '2026-09-24T00:02:00.000Z', snapshot: {
      sourceCount: 0, claimCount: 0, slideCount: plan.slides.length,
      sourcesDigest: 'a'.repeat(64), claimsDigest: 'b'.repeat(64), slidesDigest: 'c'.repeat(64), styleDigest: 'e'.repeat(64),
    } },
  ]
  const selected = { ...project, plan: { ...project.plan!, revision: 2, revisions: snapshots }, production }
  const summary = presentationWorkflowSummary(selected, imported, qa)!
  expect(summary.attention.map((item) => item.id)).toContain('style-revision')
  expect(summary.pages[0]?.qa).toBe('旧样式版本历史通过')
  expect(summary.pages[0]?.nextAction).toBe('先确认继续旧计划或选择新任务')
  expect(summary.nextTool).toBeUndefined()
})
