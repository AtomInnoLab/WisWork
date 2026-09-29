// @vitest-environment jsdom
import type { PresentationAcquisitionHistory } from '@wiswork/project-store/presentation-acquisition'
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationAcquisitionHistoryCard } from '../src/agent/presentation-acquisition-history.js'
const roots: ReturnType<typeof createRoot>[] = []
afterEach(async () => {
  for (const root of roots.splice(0)) await act(async () => root.unmount())
})
const history = (): PresentationAcquisitionHistory => ({
  version: 1,
  scope: 'remote_material_acquisition',
  documentId: 'doc1',
  revision: 3,
  totalAttempts: 2,
  records: [1, 2].map((attempt) => ({
    id: `attempt${attempt}`,
    attempt,
    kind: 'webpage',
    source: 'https://example.com/page',
    sourceUrlHash: 'a'.repeat(64),
    startedAt: '2026-09-29T00:00:00.000Z',
    ...(attempt === 1
      ? {
          state: 'rejected' as const,
          finishedAt: '2026-09-29T00:00:01.000Z',
          error: 'remote_webpage_unavailable' as const,
        }
      : { state: 'fetching' as const }),
  })),
})
async function mount(
  read = vi.fn(async (): Promise<PresentationAcquisitionHistory | undefined> => history()),
) {
  const node = document.createElement('div')
  const root = createRoot(node)
  roots.push(root)
  await act(async () =>
    root.render(React.createElement(PresentationAcquisitionHistoryCard, { read, refreshKey: 0 })),
  )
  return { node, root, read }
}
it('shows collapsed grouped durable history, each attempt, safe status and refresh', async () => {
  const { node, read } = await mount()
  const details = node.querySelector('details')!
  expect(details).not.toBeNull()
  expect(details.open).toBe(false)
  expect(node.querySelectorAll('[data-acquisition-source]')).toHaveLength(1)
  expect(node.textContent).toContain('尝试 #1')
  expect(node.textContent).toContain('尝试 #2')
  expect(node.textContent).toContain('缺少完成回执')
  expect(node.textContent).toContain('不表示后台仍在运行')
  expect(node.textContent).toContain('不代表事实、许可或 QA 验收通过')
  expect(node.textContent).toContain(
    '记录时间是获取操作时间；缓存复用也会记录，不代表来源被重新抓取或更新。',
  )
  expect(node.textContent).toContain('2026-09-29T00:00:00.000Z')
  await act(async () => (node.querySelector('button') as HTMLButtonElement).click())
  expect(read).toHaveBeenCalledTimes(2)
})
it('hides unsupported history and rejects stale results after runtime change', async () => {
  const { node, root } = await mount(vi.fn(async () => undefined))
  expect(node.textContent).toBe('')
  let resolve!: (value: PresentationAcquisitionHistory | undefined) => void
  await act(async () =>
    root.render(
      React.createElement(PresentationAcquisitionHistoryCard, {
        read: () =>
          new Promise<PresentationAcquisitionHistory | undefined>((done) => {
            resolve = done
          }),
      }),
    ),
  )
  await act(async () =>
    root.render(
      React.createElement(PresentationAcquisitionHistoryCard, { read: async () => undefined }),
    ),
  )
  await act(async () => resolve(history()))
  expect(node.textContent).toBe('')
})
it('uses safe read errors and marks the retained 64-attempt window', async () => {
  const value = { ...history(), totalAttempts: 65 }
  const { node, root } = await mount(vi.fn(async () => value))
  expect(node.textContent).toContain('最近')
  expect(node.textContent).toContain('65')
  await act(async () =>
    root.render(
      React.createElement(PresentationAcquisitionHistoryCard, {
        read: async () => {
          throw new Error('private-token')
        },
      }),
    ),
  )
  expect(node.textContent).toContain('获取历史暂时无法读取')
  expect(node.textContent).not.toContain('private-token')
})
