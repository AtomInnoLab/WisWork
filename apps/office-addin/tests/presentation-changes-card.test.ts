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
