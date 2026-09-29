// @vitest-environment jsdom
import { researchSummary } from './presentation-research-fixture.js'
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationWorkflowCard } from '../src/agent/presentation-workflow-card.js'
import type { PresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationImportProgressController } from '../src/agent/presentation-import-progress.js'

const roots: ReturnType<typeof createRoot>[] = []
afterEach(async () => {
  for (const root of roots.splice(0)) await act(async () => root.unmount())
})

it('replays the saved project stage and updates from the import record', async () => {
  const { benchmarkPlan } =
    await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
  const plan = benchmarkPlan()
  let projectSnapshot: ReturnType<PresentationProjectController['snapshot']> = {
    phase: 'idle',
    project: {
      projectId: plan.projectId,
      title: plan.title,
      status: 'planned',
      plan: { revision: 1, value: plan },
      slideCount: plan.slides.length,
      slides: plan.slides.map(({ id, title }) => ({ id, title })),
      history: [],
    },
  }
  const projectListeners = new Set<() => void>()
  const importListeners = new Set<() => void>()
  let importRevision = 0
  const importRecord: { current?: ReturnType<PresentationImportProgressController['read']> } = {}
  const project = {
    snapshot: () => projectSnapshot,
    subscribe: (listener: () => void) => {
      projectListeners.add(listener)
      return () => {
        projectListeners.delete(listener)
      }
    },
    prepareProduction: vi.fn(async () => {}),
    readDeliveryReport: vi.fn(async () => {}),
    startProductionJob: vi.fn(async () => {}),
    resumeProductionJob: vi.fn(async () => {}),
    runProduction: vi.fn(async () => {}),
  } as unknown as PresentationProjectController
  const imported = {
    revision: () => importRevision,
    subscribe: (listener: () => void) => {
      importListeners.add(listener)
      return () => {
        importListeners.delete(listener)
      }
    },
    read: () => importRecord.current,
  } satisfies PresentationImportProgressController
  const node = document.createElement('div')
  const root = createRoot(node)
  roots.push(root)
  await act(async () =>
    root.render(React.createElement(PresentationWorkflowCard, { project, imported })),
  )
  expect(node.textContent).toContain('启动逐页生产')
  expect(node.textContent).toContain('真实性仍需核验')
  expect(node.textContent).toContain('目标与资料 · 已有记录')
  expect(node.textContent).toContain('逐页制作 · 待开始')
  expect(node.textContent).toContain('恢复记录 · 1 项')
  projectSnapshot = {
    ...projectSnapshot,
    project: {
      ...projectSnapshot.project!,
      production: {
        projectId: plan.projectId,
        requestId: 'run1',
        planRevision: 1,
        status: 'partial',
        total: plan.slides.length,
        compiledCount: plan.slides.length - 1,
        pages: plan.slides.map(({ id, title }, index) => ({
          id,
          title,
          state: index === 0 ? ('failed' as const) : ('compiled' as const),
          attempt: 1,
        })),
      },
    },
  }
  await act(async () => projectListeners.forEach((listener) => listener()))
  expect(node.querySelector('[aria-label="待处理问题"]')?.textContent).toContain('1 页编译失败')
  expect(node.textContent).toContain('逐页制作 · 需处理')
  await act(async () =>
    Array.from(node.querySelectorAll('button'))
      .find((button) => button.textContent === '后台制作剩余页面')!
      .click(),
  )
  expect(project.startProductionJob).toHaveBeenCalledWith('run1')
  projectSnapshot = {
    ...projectSnapshot,
    project: {
      ...projectSnapshot.project!,
      production: {
        projectId: plan.projectId,
        requestId: 'run1',
        planRevision: 1,
        status: 'compiled',
        total: plan.slides.length,
        compiledCount: plan.slides.length,
        pages: plan.slides.map(({ id, title }) => ({ id, title, state: 'compiled', attempt: 1 })),
      },
    },
  }
  await act(async () => projectListeners.forEach((listener) => listener()))
  expect(node.querySelectorAll('ol')[0]?.querySelectorAll('li')).toHaveLength(6)
  expect(node.textContent).toContain('逐页制作 · 已有记录')
  expect(node.textContent).toContain('逐页状态')
  expect(node.textContent).toContain('确认逐页导入')
  expect(node.textContent).toContain('恢复记录 · 2 项')
  await act(async () =>
    Array.from(node.querySelectorAll('button'))
      .find((button) => button.textContent === '准备逐页导入')!
      .click(),
  )
  expect(project.prepareProduction).toHaveBeenCalledOnce()
  importRecord.current = {
    source: 'production',
    projectId: plan.projectId,
    requestId: 'run1',
    total: plan.slides.length,
    completed: plan.slides.length,
    status: 'complete',
    pages: plan.slides.map(({ id, title }) => ({ id, title, state: 'complete' })),
  }
  importRevision++
  await act(async () => importListeners.forEach((listener) => listener()))
  expect(node.textContent).toContain('采集页面截图')
  expect(node.textContent).toContain('恢复记录 · 3 项')
  expect(node.textContent).toContain('保存重开')
})

it('shows original research start and end in folded records and updates in place after an interrupted record ends', async () => {
  const summary = researchSummary()
  const record = summary.records[0]!
  record.state = 'running'
  delete record.finishedAt
  summary.revision = 1
  let snapshot: ReturnType<PresentationProjectController['snapshot']> = {
    phase: 'idle',
    project: {
      projectId: 'research',
      title: '研究项目',
      status: 'planned',
      slideCount: 0,
      slides: [],
      history: [],
      researchSummary: summary,
    },
  }
  const listeners = new Set<() => void>()
  const project = {
    snapshot: () => snapshot,
    subscribe: (listener: () => void) => {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
  } as unknown as PresentationProjectController
  const node = document.createElement('div')
  const root = createRoot(node)
  roots.push(root)
  await act(async () => root.render(React.createElement(PresentationWorkflowCard, { project })))
  const findRecords = () =>
    Array.from(node.querySelectorAll('details')).find((item) =>
      item.querySelector(':scope > summary')?.textContent?.includes('研究整理记录'),
    )!
  expect(findRecords()).toBeDefined()
  expect(findRecords().open).toBe(false)
  expect(findRecords().querySelectorAll('time')).toHaveLength(1)
  expect(findRecords().querySelector('time')?.dateTime).toBe(record.startedAt)
  expect(node.querySelector('[aria-label="待处理问题"]')).not.toBeNull()
  const ended = structuredClone(summary)
  ended.revision = 2
  ended.records[0]!.state = 'failed'
  ended.records[0]!.error = 'aborted'
  ended.records[0]!.finishedAt = '2026-09-29T00:00:03.000Z'
  snapshot = { ...snapshot, project: { ...snapshot.project!, researchSummary: ended } }
  await act(async () => listeners.forEach((listener) => listener()))
  expect(
    Array.from(node.querySelectorAll('summary')).filter((item) =>
      item.textContent?.includes('研究整理记录'),
    ),
  ).toHaveLength(1)
  expect(findRecords().querySelectorAll('time')).toHaveLength(2)
  expect(findRecords().querySelectorAll('time')[1]!.dateTime).toBe(ended.records[0]!.finishedAt)
  expect(findRecords().open).toBe(false)
  await act(async () => findRecords().querySelector('summary')!.click())
  expect(findRecords().open).toBe(true)
  expect(findRecords().querySelectorAll('time')).toHaveLength(2)
  expect(findRecords().textContent).toContain(record.startedAt)
  expect(findRecords().textContent).toContain(ended.records[0]!.finishedAt!)
})
