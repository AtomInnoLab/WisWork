// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, expect, it } from 'vitest'
import { PresentationStageCard } from '../src/agent/presentation-stage-card.js'
import type { PresentationStageGroup } from '../src/agent/presentation-stage-timeline.js'

const roots: ReturnType<typeof createRoot>[] = []
afterEach(async () => {
  for (const root of roots.splice(0)) await act(async () => root.unmount())
})

const group = (state: 'running' | 'complete' | 'error'): PresentationStageGroup => ({
  kind: 'stage',
  id: 'event-1',
  stage: 'production',
  running: state === 'running' ? 1 : 0,
  ended: state === 'complete' ? 1 : 0,
  failed: state === 'error' ? 1 : 0,
  events: [
    {
      id: 'event-1',
      kind: 'tool',
      callId: 'call-1',
      name: 'run_presentation_production',
      summary: 'Operation',
      state,
    },
  ],
})

async function mount(value: PresentationStageGroup) {
  const container = document.createElement('div')
  const root = createRoot(container)
  roots.push(root)
  const render = async (next: PresentationStageGroup) => {
    await act(async () => root.render(React.createElement(PresentationStageCard, { group: next })))
  }
  await render(value)
  return { container, render }
}

it('shows the stage and progress while keeping internal operations collapsed', async () => {
  const { container } = await mount(group('running'))
  expect(container.querySelector('strong')?.textContent).toBe('逐页制作')
  expect(container.querySelector('[role="status"]')?.textContent).toContain('进行中 1')
  const details = container.querySelector('details')
  expect(details).not.toBeNull()
  expect(details?.open).toBe(false)
  expect(details?.querySelector('summary')?.textContent).toContain('内部操作')
  expect(details?.textContent).toContain('run_presentation_production')
  expect(container.textContent).not.toContain('整套制作完成')
})

it('keeps failed operations and the recovery action visible outside the collapsed details', async () => {
  const { container } = await mount(group('error'))
  const alert = container.querySelector('[role="alert"]')
  expect(alert).not.toBeNull()
  expect(alert?.closest('details')).toBeNull()
  expect(alert?.textContent).toContain('恢复记录')
  expect(container.querySelector('[role="status"]')?.textContent).toContain('未完成 1')
})

it('updates the same stage without closing the details the reader opened', async () => {
  const { container, render } = await mount(group('running'))
  const details = container.querySelector('details')
  expect(details).not.toBeNull()
  details!.open = true
  await render(group('complete'))
  expect(container.querySelector('details')).toBe(details)
  expect(details!.open).toBe(true)
  expect(container.querySelector('[role="status"]')?.textContent).toContain('已结束 1')
  expect(container.textContent).toContain('项目记录')
})

it('does not treat a previous failed operation as a blocker while the stage continues', async () => {
  const failed = group('error')
  const current = group('running').events[0]!
  const { container } = await mount({
    ...failed,
    running: 1,
    events: [...failed.events, { ...current, id: 'event-2', callId: 'call-2' }],
  })
  expect(container.textContent).toContain('当前操作仍在进行')
  expect(container.querySelector('[role="alert"]')?.textContent).toContain('记录')
  expect(container.textContent).not.toContain('再决定下一步')
})
