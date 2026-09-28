// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { PresentationChangesCard } from '../src/agent/presentation-changes-card.js'
import type { PresentationChangesController } from '../src/agent/presentation-changes.js'
it('dispatches actual actions, escapes text and disables busy controls', async () => {
  const run = vi.fn(),
    refresh = vi.fn()
  const controller: PresentationChangesController = {
    snapshot: () => ({
      phase: 'idle',
      projectId: 'project',
      requestId: 'request',
      entries: [
        {
          id: 'id',
          source: 'existing',
          review: {
            screenshotDigest: 'a'.repeat(64),
            capturedAt: '2026-09-24T00:00:00Z',
            reviewedAt: '2026-09-24T00:01:00Z',
            status: 'pass',
            notes: '',
          },
          kind: 'text',
          pageId: 'page',
          state: 'pending',
          before: '<script>unsafe</script>',
          after: 'Changed',
          actions: ['inspect', 'resume'],
        },
      ],
    }),
    subscribe: () => () => {},
    run,
    refresh,
    clear: vi.fn(),
  }
  const container = document.createElement('div'),
    root = createRoot(container)
  try {
    await act(async () => root.render(React.createElement(PresentationChangesCard, { controller })))
    expect(container.querySelector('script')).toBeNull()
    expect(container.textContent).toContain('<script>unsafe</script>')
    expect(container.textContent).toContain('不完整历史')
    expect(container.textContent).toContain('现稿 · 页面')
    expect(container.textContent).toContain('历史截图复核：通过')
    expect(container.textContent).toContain('此结果不代表当前页面 QA 通过')
    await act(async () =>
      Array.from(container.querySelectorAll('button'))
        .find((b) => b.textContent === '检查')!
        .click(),
    )
    expect(run).toHaveBeenCalledWith('id', 'inspect')
    await act(async () =>
      Array.from(container.querySelectorAll('button'))
        .find((b) => b.textContent === '继续')!
        .click(),
    )
    expect(run).toHaveBeenCalledWith('id', 'resume')
    await act(async () =>
      Array.from(container.querySelectorAll('button'))
        .find((b) => b.textContent === '刷新保存点')!
        .click(),
    )
    expect(refresh).toHaveBeenCalledOnce()
    await act(async () =>
      root.render(React.createElement(PresentationChangesCard, { controller, disabled: true })),
    )
    expect(Array.from(container.querySelectorAll('button')).every((b) => b.disabled)).toBe(true)
  } finally {
    await act(async () => root.unmount())
  }
})
it('brings the interrupted call savepoint forward and exposes its host inspection action', async () => {
  const refresh = vi.fn(async () => undefined)
  const run = vi.fn(async () => undefined)
  const controller: PresentationChangesController = {
    snapshot: () => ({
      phase: 'idle',
      entries: [
        {
          id: 'other',
          kind: 'text',
          pageId: 'other',
          state: 'applied',
          before: '',
          after: '',
          actions: ['inspect'],
        },
        {
          id: 'pending',
          kind: 'text',
          pageId: 'target',
          state: 'pending',
          before: '',
          after: '',
          actions: ['inspect', 'resume'],
          origin: { agentRunId: 'run-1', toolCallId: 'call-1' },
        },
      ],
    }),
    subscribe: () => () => {},
    refresh,
    run,
    clear: vi.fn(),
  }
  const container = document.createElement('div')
  const root = createRoot(container)
  try {
    await act(async () =>
      root.render(
        React.createElement(PresentationChangesCard, {
          controller,
          interruptedChange: { agentRunId: 'run-1', toolCallId: 'call-1' },
        }),
      ),
    )
    expect(refresh).toHaveBeenCalledOnce()
    const rows = Array.from(container.querySelectorAll('ol > li'))
    expect(rows[0]?.textContent).toContain('上次中断的修改')
    expect(rows[0]?.textContent).toContain('target')
    expect(container.textContent).toContain('请逐项点击“检查”')
    await act(async () =>
      rows[0]?.querySelector<HTMLButtonElement>('button[aria-label="检查 target"]')?.click(),
    )
    expect(run).toHaveBeenCalledWith('pending', 'inspect')
  } finally {
    await act(async () => root.unmount())
  }
})
it('offers actual undo/commit/discard buttons and reacts to controller updates', async () => {
  let listener = () => {},
    acting = false
  const run = vi.fn(async () => {}),
    controller: PresentationChangesController = {
      snapshot: () => ({
        phase: acting ? 'acting' : 'idle',
        entries: [
          {
            id: 'page-change',
            kind: 'page',
            pageId: 'page',
            state: 'staged',
            before: 'old',
            after: 'new',
            actions: ['undo', 'commit', 'discard'],
          },
        ],
      }),
      subscribe: (next) => {
        listener = next
        return () => {}
      },
      run,
      refresh: vi.fn(),
      clear: vi.fn(),
    }
  const container = document.createElement('div'),
    root = createRoot(container)
  try {
    await act(async () => root.render(React.createElement(PresentationChangesCard, { controller })))
    expect(container.textContent).toContain('非视觉 diff')
    for (const [label, action] of [
      ['撤销', 'undo'],
      ['提交替换', 'commit'],
      ['丢弃替换', 'discard'],
    ]) {
      await act(async () =>
        Array.from(container.querySelectorAll('button'))
          .find((b) => b.textContent === label)!
          .click(),
      )
      expect(run).toHaveBeenLastCalledWith('page-change', action)
    }
    await act(async () => {
      acting = true
      listener()
    })
    expect(Array.from(container.querySelectorAll('button')).every((b) => b.disabled)).toBe(true)
  } finally {
    await act(async () => root.unmount())
  }
})
it('shows persisted batch page reviews as historical evidence', async () => {
  const controller: PresentationChangesController = {
    snapshot: () => ({
      phase: 'idle',
      entries: [
        {
          id: 'existing_batch:one',
          source: 'existing_batch',
          kind: 'text',
          pageId: 'slide',
          state: 'applied',
          before: 'old',
          after: 'new',
          actions: ['inspect', 'undo'],
          affectedPageCount: 2,
          reviews: [
            {
              hostSlideId: 'slide',
              screenshotDigest: 'a'.repeat(64),
              capturedAt: '2026-09-24T00:00:00.000Z',
              reviewedAt: '2026-09-24T00:01:00.000Z',
              status: 'pass',
              notes: '',
            },
          ],
        },
      ],
    }),
    subscribe: () => () => {},
    run: vi.fn(),
    refresh: vi.fn(),
    clear: vi.fn(),
  }
  const container = document.createElement('div'),
    root = createRoot(container)
  try {
    await act(async () => root.render(React.createElement(PresentationChangesCard, { controller })))
    expect(container.textContent).toContain('现稿批量')
    expect(container.textContent).toContain('历史截图复核：1/2 个受影响页面')
    expect(container.textContent).toContain('slide：通过')
    expect(container.textContent).toContain('不代表当前页面 QA 通过')
  } finally {
    await act(async () => root.unmount())
  }
})

it('shows partial batch progress and both recovery choices', async () => {
  const run = vi.fn()
  const controller: PresentationChangesController = {
    snapshot: () => ({
      phase: 'idle',
      entries: [
        {
          id: 'existing_batch:partial',
          source: 'existing_batch',
          kind: 'text',
          pageId: 'slide',
          state: 'applying',
          before: 'old',
          after: 'new',
          cursor: 1,
          operationCount: 2,
          actions: ['inspect', 'resume', 'undo'],
        },
      ],
    }),
    subscribe: () => () => {},
    run,
    refresh: vi.fn(),
    clear: vi.fn(),
  }
  const container = document.createElement('div')
  const root = createRoot(container)
  try {
    await act(async () => root.render(React.createElement(PresentationChangesCard, { controller })))
    expect(container.textContent).toContain('已持久记录 1/2 步')
    expect(container.textContent).toContain('宿主可能已有未记录的写入')
    for (const [label, action] of [
      ['继续', 'resume'],
      ['撤销', 'undo'],
    ]) {
      await act(async () =>
        container.querySelector<HTMLButtonElement>(`[aria-label="${label} slide"]`)!.click(),
      )
      expect(run).toHaveBeenLastCalledWith('existing_batch:partial', action)
    }
  } finally {
    await act(async () => root.unmount())
  }
})

it('routes the saved batch reapply button by its exact history ID', async () => {
  const run = vi.fn()
  const controller: PresentationChangesController = {
    snapshot: () => ({
      phase: 'idle',
      entries: [
        {
          id: 'existing_batch:redo',
          source: 'existing_batch',
          kind: 'text',
          pageId: 'slide',
          state: 'undone',
          before: 'old',
          after: 'new',
          actions: ['inspect', 'reapply', 'release'],
        },
      ],
    }),
    subscribe: () => () => {},
    run,
    refresh: vi.fn(),
    clear: vi.fn(),
  }
  const container = document.createElement('div')
  const root = createRoot(container)
  try {
    await act(async () => root.render(React.createElement(PresentationChangesCard, { controller })))
    await act(async () =>
      container.querySelector<HTMLButtonElement>('[aria-label="重新应用 slide"]')!.click(),
    )
    expect(run).toHaveBeenCalledWith('existing_batch:redo', 'reapply')
    expect(container.textContent).toContain('不代表当前页面 QA 通过')
  } finally {
    await act(async () => root.unmount())
  }
})

it('separates captured, pending, passed and failed existing-page evidence', async () => {
  const capturedAt = '2026-09-24T00:00:00.000Z'
  const controller: PresentationChangesController = {
    snapshot: () => ({
      phase: 'idle',
      entries: [
        {
          id: 'one',
          source: 'existing_page',
          kind: 'page',
          pageId: 'old',
          state: 'staged',
          before: 'old',
          after: 'new',
          actions: [],
          visualPageIds: ['old', 'new'],
          visualCaptures: [{ hostSlideId: 'old', screenshotDigest: 'a'.repeat(64), capturedAt }],
        },
        {
          id: 'two',
          source: 'existing_image',
          kind: 'image',
          pageId: 'slide',
          state: 'complete',
          before: 'old',
          after: 'new',
          actions: [],
          visualPageIds: ['slide'],
          visualReviews: [
            {
              hostSlideId: 'slide',
              screenshotDigest: 'b'.repeat(64),
              capturedAt,
              reviewedAt: capturedAt,
              status: 'fail',
              notes: 'crop',
            },
          ],
        },
      ],
    }),
    subscribe: () => () => {},
    run: vi.fn(),
    refresh: vi.fn(),
    clear: vi.fn(),
  }
  const container = document.createElement('div'),
    root = createRoot(container)
  try {
    await act(async () => root.render(React.createElement(PresentationChangesCard, { controller })))
    expect(container.textContent).toContain('old：已采集 · 待判断')
    expect(container.textContent).toContain('new：待采集')
    expect(container.textContent).toContain('slide：未通过')
  } finally {
    await act(async () => root.unmount())
  }
})
