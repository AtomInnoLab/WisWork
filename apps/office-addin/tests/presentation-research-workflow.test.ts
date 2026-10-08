import { expect, it } from 'vitest'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
import type { PresentationProjectStatus } from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationResearchSummary } from '@wiswork/project-store/presentation-research'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { researchSummary } from './presentation-research-fixture.js'
function fixture(summary = researchSummary()) {
  const project: PresentationProjectStatus & {
    researchSummary?: PresentationResearchSummary
    researchHistoryUnavailable?: boolean
  } = {
    projectId: 'research',
    title: '资料项目',
    status: 'planned',
    slideCount: 0,
    slides: [],
    history: [],
    researchSummary: summary,
  }
  return project
}
const view = (project: PresentationProjectStatus) =>
  presentationWorkflowSummary(project, undefined, undefined)!
it('renders standalone archived research with exact original start and end events and never treats completion as verification', () => {
  const project = fixture(),
    record = project.researchSummary!.records[0]!
  const first = view(project),
    rows = first.timeline.filter((row) => row.scope === 'research_ledger')
  expect(rows).toHaveLength(1)
  expect(rows[0]).toMatchObject({
    id: `research-ledger:${JSON.stringify(['doc', 'research', record.id, record.sequence, record.draftDigest])}`,
    type: 'research.completed',
    at: record.finishedAt,
    recordsLabel: '研究整理记录',
  })
  expect(rows[0]!.records?.map((event) => event.at)).toEqual([record.startedAt, record.finishedAt])
  expect(rows[0]!.text).toContain('来源 1 · 结论 2 · 冲突 1')
  expect(rows[0]!.text).toContain('不代表主张支持、来源权威性或时效通过')
  expect(rows[0]!.text).toContain('历史研究')
  expect(view(project).timeline.filter((row) => row.scope === 'research_ledger')).toEqual(rows)
  expect(first.nextAction).toBe(view({ ...project, researchSummary: undefined }).nextAction)
  expect(first.nextTool).toBeUndefined()
  expect(first.stages).toEqual(view({ ...project, researchSummary: undefined }).stages)
  expect(first.pages).toEqual(view({ ...project, researchSummary: undefined }).pages)
  project.researchSummary!.documentId = 'other-doc'
  expect(view(project).timeline.find((row) => row.scope === 'research_ledger')!.id).not.toBe(
    rows[0]!.id,
  )
  project.researchSummary!.documentId = 'doc'
  expect(first.attention.some((item) => item.id === 'research-conflicts')).toBe(true)
})
it('uses stable identity through terminal transition without manufacturing a finish or claiming running activity', () => {
  const project = fixture(),
    record = project.researchSummary!.records[0]!
  record.state = 'running'
  delete record.finishedAt
  const running = view(project).timeline.find((row) => row.scope === 'research_ledger')!
  expect(running.type).toBe('research.started')
  expect(running.records).toHaveLength(1)
  expect(running.at).toBe(record.startedAt)
  expect(running.text).toContain('缺少结束回执')
  expect(running.text).toContain('不能证明仍在执行或已经中断')
  expect(view(project).attention.some((item) => item.id === 'research-unfinished')).toBe(true)
  record.state = 'failed'
  record.finishedAt = '2026-09-29T00:03:00.000Z'
  record.error = 'source_unavailable'
  const failed = view(project).timeline.filter((row) => row.scope === 'research_ledger')
  expect(failed).toHaveLength(1)
  expect(failed[0]!.id).toBe(running.id)
  expect(failed[0]).toMatchObject({ type: 'research.failed', at: record.finishedAt })
  expect(failed[0]!.records?.map((event) => event.at)).toEqual([
    record.startedAt,
    record.finishedAt,
  ])
  expect(failed[0]!.text).toContain('原文资料暂不可用')
  expect(failed[0]!.text).not.toContain('source_unavailable')
})
it('marks only exact current plan binding and does not replace it with recent research or infer missing-window terminal state', () => {
  const project = fixture(),
    record = project.researchSummary!.records[0]!,
    plan = benchmarkPlan()
  plan.projectId = 'research'
  plan.research = {
    ledgerId: record.id,
    sequence: record.sequence,
    draftDigest: record.draftDigest,
    sources: [],
    claims: [],
  }
  project.plan = { revision: 1, value: plan }
  expect(view(project).timeline.find((row) => row.scope === 'research_ledger')!.text).toContain(
    '当前计划精确绑定',
  )
  for (const patch of [{ sequence: 2 }, { draftDigest: 'f'.repeat(64) }, { ledgerId: 'missing' }]) {
    plan.research = {
      ledgerId: record.id,
      sequence: record.sequence,
      draftDigest: record.draftDigest,
      sources: [],
      claims: [],
      ...patch,
    }
    const summary = view(project)
    expect(summary.timeline.find((row) => row.scope === 'research_ledger')!.text).toContain(
      '历史研究',
    )
    expect(summary.attention.find((item) => item.id === 'research-binding-window')?.text).toContain(
      '不能据此推断删除或完成',
    )
    expect(summary.nextTool).toBe(view({ ...project, researchSummary: undefined }).nextTool)
  }
})
it('preserves V2 deletion gaps, displays the retained window and handles unreadable or legacy history', () => {
  const project = fixture(),
    summary = project.researchSummary!,
    record = summary.records[0]!
  Object.assign(summary, {
    version: 2,
    lastSequence: 90,
    revision: 230,
    totalRecords: 40,
    records: Array.from({ length: 32 }, (_, index) => ({
      ...record,
      id: `ledger${index}`,
      sequence: index * 2 + 1,
    })),
  })
  const shown = view(project)
  expect(shown.timeline.filter((row) => row.scope === 'research_ledger')).toHaveLength(32)
  expect(shown.attention.find((item) => item.id === 'research-history-window')?.text).toContain(
    '现存 40 条；仅展示最近 32 条',
  )
  expect(
    shown.timeline
      .filter((row) => row.scope === 'research_ledger')
      .every((row) => !row.text.includes('研究 #2 ·')),
  ).toBe(true)
  project.researchHistoryUnavailable = true
  expect(view(project).timeline.some((row) => row.scope === 'research_ledger')).toBe(false)
  expect(view(project).attention.some((item) => item.id === 'research-history-unavailable')).toBe(
    true,
  )
  delete project.researchSummary
  delete project.researchHistoryUnavailable
  expect(view(project).timeline.some((row) => row.scope === 'research_ledger')).toBe(false)
  project.researchSummary = { ...researchSummary(), projectId: 'other' }
  expect(view(project).timeline.some((row) => row.scope === 'research_ledger')).toBe(false)
})
it('labels literal source excerpt audits separately from standalone research events', () => {
  const project = fixture()
  project.sourceAuditHistory = {
    version: 1,
    documentId: 'doc',
    projectId: 'research',
    revision: 1,
    runs: [
      {
        id: 'audit',
        sequence: 1,
        scope: 'source_excerpt_audit',
        planRevision: 1,
        planDigest: 'a'.repeat(64),
        state: 'running',
        startedAt: '2026-09-29T00:00:00.000Z',
        sourceCount: 1,
      },
    ],
  }
  const audit = view(project).timeline.find((row) => row.scope === 'source_excerpt_audit')!
  expect(audit.type).toBe('source_audit.started')
  expect(audit.id).toContain('source-audit:')
  expect(
    view(project)
      .timeline.filter((row) => row.type?.startsWith('research.'))
      .every((row) => row.scope === 'research_ledger'),
  ).toBe(true)
})
