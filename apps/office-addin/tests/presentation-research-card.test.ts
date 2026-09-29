// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationResearchCard } from '../src/agent/presentation-research-card.js'
import type {
  PresentationResearchController,
  PresentationResearchSnapshot,
} from '../src/agent/presentation-research.js'
import { researchRecord, researchSummary } from './presentation-research-fixture.js'
const roots: ReturnType<typeof createRoot>[] = []
afterEach(async () => {
  for (const root of roots.splice(0)) await act(async () => root.unmount())
})
async function mount(snapshot: PresentationResearchSnapshot) {
  const listeners = new Set<() => void>()
  const controller: PresentationResearchController = {
    snapshot: () => snapshot,
    subscribe: (listener) => {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    refresh: vi.fn(async () => {}),
    selectProject: vi.fn(async () => {}),
    read: vi.fn(async () => {}),
    export: vi.fn(async () => {}),
    cancel: vi.fn(),
    clear: vi.fn(),
  }
  const node = document.createElement('div')
  const root = createRoot(node)
  roots.push(root)
  await act(async () =>
    root.render(React.createElement(PresentationResearchCard, { controller, disabled: false })),
  )
  return {
    node,
    controller,
    update: async (next: PresentationResearchSnapshot) => {
      snapshot = next
      await act(async () => listeners.forEach((l) => l()))
    },
  }
}
it('shows research before a plan, with folded facts/conflict/source gaps and read/export actions', async () => {
  const f = await mount({
    phase: 'idle',
    available: true,
    projectId: 'research',
    summary: researchSummary(),
    record: researchRecord(),
  })
  expect(f.node.textContent).toContain('资料研究账本')
  expect(f.node.querySelector('details')?.open).toBe(false)
  expect(f.node.textContent).toContain('销售增长')
  expect(f.node.textContent).toContain('销售下降')
  expect(f.node.textContent).toContain('事实主张')
  expect(f.node.textContent).toContain('判断')
  expect(f.node.textContent).toContain('冲突')
  expect(f.node.textContent).toContain('原文缺失')
  expect(f.node.textContent).toContain('不代表来源权威性或时效通过')
  expect(f.node.querySelector('a')?.href).toBe('https://example.com/report')
  const button = (label: string) =>
    Array.from(f.node.querySelectorAll('button')).find((b) => b.textContent === label)!
  await act(async () => button('读取完整研究记录').click())
  expect(f.controller.read).toHaveBeenCalledWith('ledger1')
  await act(async () => button('导出研究 JSON 与 Markdown').click())
  expect(f.controller.export).toHaveBeenCalledWith('ledger1')
})
it('hides old PCs and never labels a running record as active background research', async () => {
  const f = await mount({ phase: 'idle', available: false })
  expect(f.node.textContent).toBe('')
  const summary = researchSummary()
  const { finishedAt: _end, ...running } = summary.records[0]!
  await f.update({
    phase: 'idle',
    available: true,
    projectId: 'research',
    summary: { ...summary, revision: 1, records: [{ ...running, state: 'running' }] },
  })
  expect(f.node.textContent).toContain('缺少结束回执')
  expect(f.node.textContent).toContain('不表示后台仍在运行')
})
