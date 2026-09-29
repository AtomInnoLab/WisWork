// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
import { PresentationWorkflowCard } from '../src/agent/presentation-workflow-card.js'
import { PresentationQaCard } from '../src/agent/presentation-qa-card.js'
import type {
  PresentationProjectStatus,
  PresentationProjectController,
} from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationQaController } from '../src/agent/presentation-qa-card.js'
import type { PresentationQaAttempt } from '../src/skills/powerpoint/presentation-qa-attempts.js'

const startedAt = '2026-09-29T00:00:00.000Z',
  finishedAt = '2026-09-29T00:01:00.000Z'
const project: PresentationProjectStatus = {
  projectId: 'project',
  title: 'Project',
  status: 'planned',
  slideCount: 1,
  slides: [{ id: 'page', title: 'Page' }],
  history: [],
  production: {
    projectId: 'project',
    requestId: 'run',
    planRevision: 1,
    status: 'compiled',
    compiledCount: 1,
    total: 1,
    pages: [{ id: 'page', title: 'Page', state: 'compiled', attempt: 1 }],
  },
}
const attempt = () => ({
  version: 1 as const,
  id: '12345678-1234-4234-8234-123456789abc',
  source: 'production' as const,
  documentId: 'doc',
  projectId: 'project',
  requestId: 'run',
  artifactDigest: 'a'.repeat(64),
  pageId: 'page',
  hostSlideId: 'new-slide',
  startedAt,
  status: 'waiting' as const,
  finishedAt,
  errorCode: 'screenshot_unavailable' as const,
})
const events = (attempts: unknown[] = [attempt()], unavailable = false) =>
  presentationWorkflowSummary(project, undefined, undefined, undefined, undefined, {
    attempts: attempts as PresentationQaAttempt[],
    unavailable,
  })!.timeline.filter((event) => event.scope === 'page_qa_attempt')

it('folds actual start/end timestamps and keeps durable attempt identity stable on reread', () => {
  const first = events()[0]!
  expect(first).toMatchObject({
    type: 'qa.attempt.waiting',
    at: finishedAt,
    records: [
      { type: 'qa.attempt.started', at: startedAt },
      { type: 'qa.attempt.waiting', at: finishedAt },
    ],
  })
  expect(first.text).toContain('截图')
  expect(first.text).toContain('重试')
  expect(events()).toEqual(events())
  const different = { ...attempt(), id: '22345678-1234-4234-8234-123456789abc' }
  expect(events([different])[0]!.id).not.toBe(first.id)
})

it('shows unresolved starts and only actual recorded capture without implying active work or QA pass', () => {
  const { finishedAt: _end, errorCode: _error, ...base } = attempt()
  const start = events([{ ...base, status: 'started' }])[0]!
  expect(start).toMatchObject({ type: 'qa.attempt.started', at: startedAt })
  expect(start.records).toHaveLength(1)
  expect(start.text).toContain('不能证明仍在执行')
  const recorded = events([{ ...base, status: 'recorded', finishedAt }])[0]!
  expect(recorded.text).toContain('已保存截图')
  expect(recorded.text).toContain('不代表')
})

it('rejects malformed, foreign page/task, and unavailable attempt history without changing workflow decisions', () => {
  for (const change of [
    { projectId: 'other' },
    { requestId: 'other' },
    { pageId: 'other' },
    { id: 'not-uuid' },
    { startedAt: 'bad' },
    { finishedAt: '2026-09-28T00:00:00.000Z' },
  ])
    expect(events([{ ...attempt(), ...change }])).toEqual([])
  expect(events([attempt()], true)).toEqual([])
  const invalid = presentationWorkflowSummary(project, undefined, undefined, undefined, undefined, {
    attempts: [{ ...attempt(), id: 'invalid' }],
  })!
  expect(invalid.attention.some((item) => item.id === 'qa-attempt-history-unavailable')).toBe(true)
  const overCapacity = Array.from({ length: 65 }, (_, index) => ({
    ...attempt(),
    id: `${String(index).padStart(8, '0')}-1234-4234-8234-123456789abc`,
  }))
  expect(events(overCapacity)).toEqual([])
  const old = presentationWorkflowSummary(project, undefined, undefined)!
  const next = presentationWorkflowSummary(project, undefined, undefined, undefined, undefined, {
    attempts: [attempt()],
  })!
  expect(next.stages).toEqual(old.stages)
  expect(next.pages).toEqual(old.pages)
  expect(next.nextTool).toEqual(old.nextTool)
})

it('keeps failures and cancellations safe and does not expose raw SDK errors', () => {
  for (const status of ['failed', 'cancelled']) {
    const a = {
      ...attempt(),
      status,
      errorCode: status === 'cancelled' ? 'cancelled' : 'publication_failed',
    }
    const row = events([a])[0]!
    expect(row.text).not.toContain('publication_failed')
    expect(row.text).not.toContain('cancelled')
    expect(row.text).toContain('未')
  }
})

it('shows attempt-only pending screenshots, folded click history and explicit read failure without auto requests', async () => {
  const node = document.createElement('div'),
    root = createRoot(node)
  const snapshot = { phase: 'idle', project },
    noRequest = vi.fn()
  const controller = {
    snapshot: () => snapshot,
    subscribe: () => () => {},
    startProductionJob: noRequest,
  } as unknown as PresentationProjectController
  let fail = false
  const qa = {
    read: () => undefined,
    revision: () => 0,
    subscribe: () => () => {},
    attempts: () => {
      if (fail) throw Error('secret SDK failure')
      return [attempt()]
    },
  } as unknown as PresentationQaController
  try {
    await act(async () =>
      root.render(
        React.createElement(
          'div',
          {},
          React.createElement(PresentationQaCard, { controller: qa }),
          React.createElement(PresentationWorkflowCard, { project: controller, qa }),
        ),
      ),
    )
    expect(node.textContent).toContain('截图')
    expect(node.textContent).toContain('重试')
    const history = Array.from(node.querySelectorAll('details')).find((d) =>
      d.querySelector('summary')?.textContent?.includes('截图尝试记录'),
    )!
    expect(history).toBeDefined()
    expect(history.open).toBe(false)
    await act(async () => history.querySelector('summary')!.click())
    expect(history.open).toBe(true)
    expect(Array.from(history.querySelectorAll('time')).map((t) => t.dateTime)).toEqual([
      startedAt,
      finishedAt,
    ])
    fail = true
    await act(async () =>
      root.render(
        React.createElement(
          'div',
          {},
          React.createElement(PresentationQaCard, { controller: qa }),
          React.createElement(PresentationWorkflowCard, { project: controller, qa }),
        ),
      ),
    )
    expect(node.textContent).toContain('截图尝试历史暂不可读取')
    expect(node.textContent).not.toContain('secret SDK failure')
    expect(node.textContent).not.toContain('等待截图')
    expect(noRequest).not.toHaveBeenCalled()
  } finally {
    await act(async () => root.unmount())
  }
})
