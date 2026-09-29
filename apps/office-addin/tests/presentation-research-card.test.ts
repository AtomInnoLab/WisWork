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
    deleteRecord: vi.fn(async () => {}),
    retryDelete: vi.fn(async () => {}),
    checkDeleteStatus: vi.fn(async () => {}),
    abandonRecord: vi.fn(async () => {}),
    retryAbandon: vi.fn(async () => {}),
    checkAbandonStatus: vi.fn(async () => {}),
    forgetAbandon: vi.fn(async () => {}),
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

it('requires inline confirmation for finished archive cleanup and isolates document/project changes', async () => {
  const summary = researchSummary()
  const f = await mount({
    phase: 'idle',
    available: true,
    cleanupAvailable: true,
    projectId: 'research',
    summary,
  })
  const button = (label: string) =>
    Array.from(f.node.querySelectorAll('button')).find((item) => item.textContent === label)!
  await act(async () => button('清理本机研究归档').click())
  expect(f.controller.deleteRecord).not.toHaveBeenCalled()
  expect(f.node.textContent).toContain('只删除本机这条已结束且未被引用的研究归档')
  expect(f.node.textContent).toContain('原附件、PowerPoint 文稿、交付包与导出副本仍保留')
  await act(async () => button('取消清理').click())
  expect(f.controller.deleteRecord).not.toHaveBeenCalled()
  await act(async () => button('清理本机研究归档').click())
  await f.update({
    phase: 'idle',
    available: true,
    cleanupAvailable: true,
    projectId: 'other',
    summary: { ...summary, projectId: 'other' },
  })
  expect(f.node.textContent).not.toContain('确认清理此归档')
  expect(f.controller.deleteRecord).not.toHaveBeenCalled()
  await f.update({
    phase: 'idle',
    available: true,
    cleanupAvailable: true,
    projectId: 'research',
    summary,
  })
  await act(async () => button('清理本机研究归档').click())
  await act(async () => button('确认清理此归档').click())
  expect(f.controller.deleteRecord).toHaveBeenCalledWith('ledger1', summary.records[0]!.draftDigest)
  const { finishedAt: _end, ...running } = summary.records[0]!
  await f.update({
    phase: 'idle',
    available: true,
    cleanupAvailable: true,
    projectId: 'research',
    summary: { ...summary, revision: 1, records: [{ ...running, state: 'running' }] },
  })
  expect(f.node.textContent).toContain('未收到结束回执，不能清理此归档')
  expect(f.node.textContent).not.toContain('清理本机研究归档')
  await f.update({
    phase: 'idle',
    available: true,
    cleanupAvailable: false,
    projectId: 'research',
    summary,
  })
  expect(f.node.textContent).not.toContain('清理本机研究归档')
})
it('offers read-only receipt recovery and an explicit retry without claiming cancellation undoes cleanup', async () => {
  const summary = researchSummary()
  const f = await mount({
    phase: 'idle',
    available: true,
    cleanupAvailable: true,
    projectId: 'research',
    summary,
    deleteAttempt: {
      documentId: 'doc',
      projectId: 'research',
      ledgerId: 'ledger1',
      sequence: 1,
      draftDigest: summary.records[0]!.draftDigest,
      deleteId: 'delete1',
      expectedRevision: 2,
    },
  })
  const button = (label: string) =>
    Array.from(f.node.querySelectorAll('button')).find((item) => item.textContent === label)!
  await act(async () => button('读取删除回执').click())
  expect(f.controller.checkDeleteStatus).toHaveBeenCalledOnce()
  expect(f.controller.retryDelete).not.toHaveBeenCalled()
  await act(async () => button('重试同一次清理').click())
  expect(f.controller.retryDelete).not.toHaveBeenCalled()
  await act(async () => button('确认重试同一次清理').click())
  expect(f.controller.retryDelete).toHaveBeenCalledOnce()
  await f.update({ ...f.controller.snapshot(), phase: 'deleting' })
  expect(f.node.textContent).toContain('停止等待不承诺撤销已提交的清理')
  expect(button('读取删除回执').disabled).toBe(true)
})
it('requires inline confirmation for running recovery and isolates confirmation on project change', async () => {
  const summary = researchSummary()
  summary.records[0]!.state = 'running'
  delete summary.records[0]!.finishedAt
  const f = await mount({
    phase: 'idle',
    available: true,
    recoveryAvailable: true,
    projectId: 'research',
    summary,
  })
  const click = async (text: string) => {
    const button = Array.from(f.node.querySelectorAll('button')).find(
      (item) => item.textContent === text,
    )!
    expect(button).toBeDefined()
    await act(async () => button.click())
  }
  await click('结束未完成研究')
  expect(f.controller.abandonRecord).not.toHaveBeenCalled()
  expect(f.node.textContent).toContain('活跃研究正常完成时不会覆盖其结果')
  await click('取消结束')
  expect(f.controller.abandonRecord).not.toHaveBeenCalled()
  await click('结束未完成研究')
  await click('确认结束此研究')
  expect(f.controller.abandonRecord).toHaveBeenCalledWith(
    'ledger1',
    summary.records[0]!.draftDigest,
  )
  await click('结束未完成研究')
  await f.update({
    phase: 'idle',
    available: true,
    recoveryAvailable: true,
    projectId: 'other',
    summary: { ...summary, projectId: 'other' },
  })
  expect(f.node.textContent).not.toContain('确认结束此研究')
  await f.update({ phase: 'idle', available: true, projectId: 'research', summary })
  expect(f.node.textContent).not.toContain('结束未完成研究')
})
it('pending recovery queries only and requires confirmation for retry or local forgetting', async () => {
  const summary = researchSummary()
  const f = await mount({
    phase: 'idle',
    available: true,
    recoveryAvailable: true,
    projectId: 'research',
    summary,
    abandonAttempt: {
      documentId: 'doc',
      projectId: 'research',
      ledgerId: 'ledger1',
      sequence: 1,
      draftDigest: summary.records[0]!.draftDigest,
      expectedRevision: 1,
    },
  })
  const click = async (text: string) => {
    const button = Array.from(f.node.querySelectorAll('button')).find(
      (item) => item.textContent === text,
    )!
    expect(button).toBeDefined()
    await act(async () => button.click())
  }
  await click('读取结束状态')
  expect(f.controller.checkAbandonStatus).toHaveBeenCalledOnce()
  expect(f.controller.retryAbandon).not.toHaveBeenCalled()
  await click('重试同一次结束')
  expect(f.controller.retryAbandon).not.toHaveBeenCalled()
  await click('确认重试同一次结束')
  expect(f.controller.retryAbandon).toHaveBeenCalledOnce()
  await click('仅忘记本机恢复身份')
  expect(f.controller.forgetAbandon).not.toHaveBeenCalled()
  expect(f.node.textContent).toContain('不修改 PC 研究记录')
  await click('确认忘记恢复身份')
  expect(f.controller.forgetAbandon).toHaveBeenCalledOnce()
})
