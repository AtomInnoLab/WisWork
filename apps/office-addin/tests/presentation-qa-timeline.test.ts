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
import type { PresentationQaRecord } from '../src/skills/powerpoint/presentation-qa.js'
import type { PresentationQaController } from '../src/agent/presentation-qa-card.js'

const capturedAt = '2026-09-29T00:00:00.000Z'
const reviewedAt = '2026-09-29T00:01:00.000Z'
const invalidatedAt = '2026-09-29T00:02:00.000Z'
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
const qa = (): PresentationQaRecord => ({
  version: 1,
  source: 'production',
  documentId: 'doc',
  projectId: 'project',
  requestId: 'run',
  artifactDigest: 'a'.repeat(64),
  pages: [
    {
      pageId: 'page',
      title: 'Page',
      hostSlideId: 'new-slide',
      capturedAt,
      screenshotDigest: 'b'.repeat(64),
      screenshotBytes: 128,
      structure: {
        status: 'passed',
        shapeCount: 1,
        overflowCount: 0,
        overlapCount: 0,
        shapesTruncated: false,
        overlapsTruncated: false,
      },
      visual: {
        status: 'pass',
        reviewer: 'agent',
        notes: 'Historical agent opinion only',
        reviewedAt,
      },
    },
  ],
})
const summary = (record = qa(), task = project) =>
  presentationWorkflowSummary(task, undefined, record)!
const events = (record = qa(), task = project) =>
  summary(record, task).timeline.filter((event) => event.scope === 'saved_page_qa')

it('folds immutable capture and Agent visual history with stable full evidence identity', () => {
  const event = events()[0]!
  expect(event).toMatchObject({
    type: 'qa.visual.recorded',
    at: reviewedAt,
    records: [
      { type: 'qa.capture.recorded', at: capturedAt },
      { type: 'qa.visual.recorded', at: reviewedAt },
    ],
  })
  expect(event.text).toContain('不代表')
  expect(event.text).toContain('Agent 历史视觉复核：通过')
  expect(events()).toEqual(events())
  for (const change of ['document', 'host', 'capture', 'digest']) {
    const r = qa()
    if (change === 'document') r.documentId = 'other'
    if (change === 'host') r.pages[0]!.hostSlideId = 'other'
    if (change === 'capture') r.pages[0]!.capturedAt = '2026-09-28T00:00:00.000Z'
    if (change === 'digest') r.pages[0]!.screenshotDigest = 'c'.repeat(64)
    expect(events(r)[0]!.id).not.toBe(event.id)
  }
})

it('records actual first invalidation but does not invent a timestamp for legacy stale evidence', () => {
  const r = qa()
  Object.assign(r.pages[0]!, { recheckRequired: true, invalidatedAt })
  expect(events(r)[0]).toMatchObject({
    type: 'qa.evidence.invalidated',
    at: invalidatedAt,
    records: [
      { type: 'qa.capture.recorded', at: capturedAt },
      { type: 'qa.visual.recorded', at: reviewedAt },
      { type: 'qa.evidence.invalidated', at: invalidatedAt },
    ],
  })
  const changed = qa()
  Object.assign(changed.pages[0]!, { recheckRequired: true })
  expect(events(changed)[0]!.type).toBe('qa.visual.recorded')
  expect(events(changed)[0]!.text).toContain('失效')
  expect(events(changed)[0]!.records).toHaveLength(2)
  expect(events(changed)[0]!.records!.some((r) => r.type === 'qa.evidence.invalidated')).toBe(false)
})

it('keeps fallback appearance pending and capture-only or partial records truthful', () => {
  const r = qa()
  r.pages[0]!.screenshotRenderer = 'libreoffice'
  r.pages[0]!.visual = { status: 'needs_review' }
  expect(events(r)[0]).toMatchObject({ type: 'qa.capture.recorded', at: capturedAt })
  expect(events(r)[0]!.text).toContain('LibreOffice')
  expect(events(r)[0]!.text).toContain('宿主外观未验')
  const partial = structuredClone(project)
  partial.production!.total = 2
  partial.production!.pages.push({ id: 'second', title: 'Second', state: 'compiled', attempt: 1 })
  expect(events(r, partial)).toHaveLength(1)
})

it('hides foreign task pages and malformed history without changing QA progress or next actions', () => {
  for (const change of [
    'project',
    'request',
    'page',
    'capture',
    'review',
    'invalidation',
    'host',
  ]) {
    const r = qa()
    if (change === 'project') r.projectId = 'other'
    if (change === 'request') r.requestId = 'other'
    if (change === 'page') r.pages[0]!.pageId = 'other'
    if (change === 'capture') r.pages[0]!.capturedAt = '2026-02-30T00:00:00.000Z'
    if (change === 'review') r.pages[0]!.visual.reviewedAt = '2026-09-28T00:00:00.000Z'
    if (change === 'invalidation')
      Object.assign(r.pages[0]!, { recheckRequired: true, invalidatedAt: capturedAt })
    if (change === 'host') r.pages[0]!.hostSlideId = ''
    expect(events(r)).toEqual([])
    if (['capture', 'review', 'host'].includes(change)) {
      expect(summary(r).pages).toEqual(summary().pages)
      expect(summary(r).stages).toEqual(summary().stages)
      expect(summary(r).attention).toEqual(summary().attention)
      expect(summary(r).nextTool).toBe(summary().nextTool)
    }
  }
})

it('requires the production project and exact current import host mapping for history', () => {
  const foreign = structuredClone(project)
  foreign.production!.projectId = 'foreign'
  expect(events(qa(), foreign)).toEqual([])
  const imported = {
    source: 'production' as const,
    projectId: 'project',
    requestId: 'run',
    total: 1,
    completed: 1,
    status: 'complete' as const,
    pages: [{ id: 'page', title: 'Page', state: 'complete' as const, slideId: 'new-slide' }],
  }
  const matching = presentationWorkflowSummary(project, imported, qa())!
  expect(matching.timeline.filter((event) => event.scope === 'saved_page_qa')).toHaveLength(1)
  imported.pages[0]!.slideId = 'different-host'
  const mismatched = presentationWorkflowSummary(project, imported, qa())!
  expect(mismatched.timeline.filter((event) => event.scope === 'saved_page_qa')).toEqual([])
  expect(mismatched.pages).toEqual(matching.pages)
  expect(mismatched.stages).toEqual(matching.stages)
  expect(mismatched.nextTool).toEqual(matching.nextTool)
})

it('starts collapsed and expands on an actual click to show distinct real event times', async () => {
  const node = document.createElement('div'),
    root = createRoot(node),
    r = qa()
  Object.assign(r.pages[0]!, { recheckRequired: true, invalidatedAt })
  const snapshot = { phase: 'idle', project }
  const controller = {
    snapshot: () => snapshot,
    subscribe: () => () => {},
  } as unknown as PresentationProjectController
  const q = {
    read: () => r,
    revision: () => 0,
    subscribe: () => () => {},
  } as unknown as PresentationQaController
  try {
    await act(async () =>
      root.render(React.createElement(PresentationWorkflowCard, { project: controller, qa: q })),
    )
    const details = Array.from(node.querySelectorAll('details')).find((d) =>
      d.querySelector('summary')?.textContent?.includes('页面 QA 记录'),
    )!
    expect(details).toBeDefined()
    expect(details.open).toBe(false)
    await act(async () => details.querySelector('summary')!.click())
    expect(details.open).toBe(true)
    expect(Array.from(details.querySelectorAll('time')).map((t) => t.dateTime)).toEqual([
      capturedAt,
      reviewedAt,
      invalidatedAt,
    ])
    expect(details.textContent).toContain('Agent')
    expect(details.textContent).toContain('专业')
  } finally {
    await act(async () => root.unmount())
  }
})
