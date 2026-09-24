// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationWorkflowCard } from '../src/agent/presentation-workflow-card.js'
import type { PresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationImportProgressController } from '../src/agent/presentation-import-progress.js'

const roots: ReturnType<typeof createRoot>[] = []
afterEach(async () => { for (const root of roots.splice(0)) await act(async () => root.unmount()) })

it('replays the saved project stage and updates from the import record', async () => {
  const { benchmarkPlan } = await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
  const plan = benchmarkPlan()
  let projectSnapshot: ReturnType<PresentationProjectController['snapshot']> = {
    phase: 'idle',
    project: { projectId: plan.projectId, title: plan.title, status: 'planned',
      plan: { revision: 1, value: plan }, slideCount: plan.slides.length,
      slides: plan.slides.map(({ id, title }) => ({ id, title })), history: [] },
  }
  const projectListeners = new Set<() => void>()
  const importListeners = new Set<() => void>()
  let importRevision = 0
  const importRecord: { current?: ReturnType<PresentationImportProgressController['read']> } = {}
  const project = {
    snapshot: () => projectSnapshot,
    subscribe: (listener: () => void) => { projectListeners.add(listener); return () => { projectListeners.delete(listener) } },
    prepareProduction: vi.fn(async () => {}),
    readDeliveryReport: vi.fn(async () => {}),
    startProductionJob: vi.fn(async () => {}),
    resumeProductionJob: vi.fn(async () => {}),
    runProduction: vi.fn(async () => {}),
  } as unknown as PresentationProjectController
  const imported = {
    revision: () => importRevision,
    subscribe: (listener: () => void) => { importListeners.add(listener); return () => { importListeners.delete(listener) } },
    read: () => importRecord.current,
  } satisfies PresentationImportProgressController
  const node = document.createElement('div')
  const root = createRoot(node)
  roots.push(root)
  await act(async () => root.render(React.createElement(PresentationWorkflowCard, { project, imported })))
  expect(node.textContent).toContain('启动逐页生产')
  expect(node.textContent).toContain('真实性仍需核验')
  expect(node.textContent).toContain('恢复记录 · 1 项')
  projectSnapshot = { ...projectSnapshot, project: { ...projectSnapshot.project!, production: {
    projectId: plan.projectId, requestId: 'run1', planRevision: 1,
    status: 'partial', total: plan.slides.length, compiledCount: plan.slides.length - 1,
    pages: plan.slides.map(({ id, title }, index) => ({ id, title,
      state: index === 0 ? 'failed' as const : 'compiled' as const, attempt: 1 })),
  } } }
  await act(async () => projectListeners.forEach((listener) => listener()))
  await act(async () => (Array.from(node.querySelectorAll('button')).find((button) => button.textContent === '后台制作剩余页面')!).click())
  expect(project.startProductionJob).toHaveBeenCalledWith('run1')
  projectSnapshot = { ...projectSnapshot, project: { ...projectSnapshot.project!, production: {
    projectId: plan.projectId, requestId: 'run1', planRevision: 1,
    status: 'compiled', total: plan.slides.length, compiledCount: plan.slides.length,
    pages: plan.slides.map(({ id, title }) => ({ id, title, state: 'compiled', attempt: 1 })),
  } } }
  await act(async () => projectListeners.forEach((listener) => listener()))
  expect(node.querySelectorAll('ol')[0]?.querySelectorAll('li')).toHaveLength(6)
  expect(node.textContent).toContain('逐页状态')
  expect(node.textContent).toContain('确认逐页导入')
  expect(node.textContent).toContain('恢复记录 · 2 项')
  await act(async () => (Array.from(node.querySelectorAll('button')).find((button) => button.textContent === '准备逐页导入')!).click())
  expect(project.prepareProduction).toHaveBeenCalledOnce()
  importRecord.current = { source: 'production', projectId: plan.projectId, requestId: 'run1',
    total: plan.slides.length, completed: plan.slides.length, status: 'complete',
    pages: plan.slides.map(({ id, title }) => ({ id, title, state: 'complete' })) }
  importRevision++
  await act(async () => importListeners.forEach((listener) => listener()))
  expect(node.textContent).toContain('采集页面截图')
  expect(node.textContent).toContain('恢复记录 · 3 项')
  expect(node.textContent).toContain('保存重开')
})
