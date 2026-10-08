// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it } from 'vitest'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
import { PresentationWorkflowCard } from '../src/agent/presentation-workflow-card.js'
import type {
  PresentationProjectStatus,
  PresentationProjectController,
} from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationImportProgress } from '../src/skills/powerpoint/presentation-page-delivery.js'
import type { PresentationImportProgressController } from '../src/agent/presentation-import-progress.js'

const startedAt = '2026-09-29T00:00:00.000Z'
const completedAt = '2026-09-29T00:01:00.000Z'
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
const progress = (): PresentationImportProgress => ({
  source: 'production',
  projectId: 'project',
  requestId: 'run',
  total: 1,
  completed: 1,
  status: 'complete',
  pages: [
    { id: 'page', title: 'Page', state: 'complete', slideId: 'new-slide', startedAt, completedAt },
  ],
})
const summary = (p = progress(), task = project) => presentationWorkflowSummary(task, p, undefined)!
const events = (p = progress(), task = project) =>
  summary(p, task).timeline.filter((event) => event.scope === 'host_page_import')

it('folds real start and recorded times with stable exact task page and host slide identity', () => {
  const event = events()[0]!
  expect(event).toMatchObject({
    type: 'host.import.recorded',
    scope: 'host_page_import',
    at: completedAt,
  })
  expect(event.records).toMatchObject([
    { type: 'host.import.started', at: startedAt },
    { type: 'host.import.recorded', at: completedAt },
  ])
  expect(event.text).toContain('不代表')
  expect(event.records![1]!.text).toContain('记录')
  expect(events()).toEqual(events())
  const changed = progress()
  changed.pages[0]!.slideId = '257#'
  expect(events(changed)[0]!.id).not.toBe(event.id)
  const other = progress()
  other.requestId = 'other'
  expect(events(other)).toEqual([])
  other.requestId = 'run'
  other.projectId = 'other'
  expect(events(other)).toEqual([])
  other.projectId = 'project'
  other.pages[0]!.id = 'other'
  expect(events(other)).toEqual([])
})

it('retains legacy completion without fabricating start and no-time progress without timeline events', () => {
  const p = progress()
  delete p.pages[0]!.startedAt
  expect(events(p)[0]!.records).toMatchObject([{ type: 'host.import.recorded', at: completedAt }])
  delete p.pages[0]!.completedAt
  expect(events(p)).toEqual([])
  expect(summary(p).pages[0]!.imported).toContain('已记录')
})

it('shows uncertain original start without inventing completion or live activity', () => {
  const p = progress()
  p.completed = 0
  p.status = 'uncertain'
  p.pages[0]!.state = 'uncertain'
  delete p.pages[0]!.slideId
  delete p.pages[0]!.completedAt
  expect(events(p)[0]).toMatchObject({
    type: 'host.import.uncertain',
    at: startedAt,
    records: [{ type: 'host.import.started', at: startedAt }],
  })
  expect(events(p)[0]!.text).toContain('待核查')
  const changed = structuredClone(p)
  changed.pages[0]!.startedAt = completedAt
  expect(events(changed)[0]!.id).not.toBe(events(p)[0]!.id)
})

it('rejects invalid timestamp or reversed chronology only from historical rows', () => {
  for (const changes of [
    { startedAt: 'invalid' },
    { completedAt: '2026-02-30T00:00:00.000Z' },
    { startedAt: completedAt, completedAt: startedAt },
    { completedAt: undefined },
    { slideId: '' },
    { slideId: 'x'.repeat(257) },
  ]) {
    const p = progress()
    Object.assign(p.pages[0]!, changes)
    expect(events(p)).toEqual([])
    expect(summary(p).pages).toEqual(summary().pages)
    expect(summary(p).nextTool).toBe(summary().nextTool)
    expect(summary(p).attention).toEqual(summary().attention)
  }
})

it('renders the true dual times in a closed history without implying QA completion', async () => {
  const node = document.createElement('div'),
    root = createRoot(node)
  const snapshot = { phase: 'idle', project },
    saved = progress()
  const controller = {
    snapshot: () => snapshot,
    subscribe: () => () => {},
  } as unknown as PresentationProjectController
  const imported = {
    read: () => saved,
    revision: () => 0,
    subscribe: () => () => {},
  } as unknown as PresentationImportProgressController
  try {
    await act(async () =>
      root.render(React.createElement(PresentationWorkflowCard, { project: controller, imported })),
    )
    const details = Array.from(node.querySelectorAll('details')).find((d) =>
      d.querySelector('summary')?.textContent?.includes('宿主页导入记录'),
    )!
    expect(details).toBeDefined()
    expect(details.open).toBe(false)
    await act(async () => details.querySelector('summary')!.click())
    expect(details.open).toBe(true)
    expect(Array.from(details.querySelectorAll('time')).map((t) => t.dateTime)).toEqual([
      startedAt,
      completedAt,
    ])
    expect(details.textContent).toContain('记录')
    expect(details.textContent).toContain('不代表')
  } finally {
    await act(async () => root.unmount())
  }
})
