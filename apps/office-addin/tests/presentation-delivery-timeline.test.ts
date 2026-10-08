import { expect, it } from 'vitest'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
import type { PresentationProjectStatus } from '../src/skills/powerpoint/presentation-project.js'
import { hostBundleReceipt } from './presentation-host-bundle-fixture.js'
const project: PresentationProjectStatus = {
  projectId: 'project-1',
  title: 'delivery',
  status: 'compiled',
  slideCount: 0,
  slides: [],
  history: [],
  production: {
    projectId: 'project-1',
    requestId: 'pages',
    planRevision: 1,
    status: 'compiled',
    total: 0,
    compiledCount: 0,
    pages: [],
  },
}
it('replays original upload and ready receipt times as one delivery step without project completion', () => {
  const ready = hostBundleReceipt()
  ready.completedAt = '2026-09-29T00:00:03.000Z'
  const pending = structuredClone(ready)
  pending.state = 'uploading'
  pending.receivedBytes = 25
  delete pending.completedAt
  const before = presentationWorkflowSummary(project, undefined, undefined, undefined, {
    bundles: [pending],
  })!
  const start = before.timeline.find((event) => event.scope === 'current_document_delivery_bundle')!
  expect(start.type).toBe('delivery.bundle.started')
  expect(start.at).toBe(pending.createdAt)
  expect(start.records).toHaveLength(1)
  const after = presentationWorkflowSummary(project, undefined, undefined, undefined, {
    bundles: [ready],
  })!
  const end = after.timeline.find((event) => event.scope === 'current_document_delivery_bundle')!
  expect(end.id).toBe(start.id)
  expect(end.type).toBe('delivery.bundle.ready')
  expect(end.at).toBe(ready.completedAt)
  expect(end.records?.map((record) => record.at)).toEqual([ready.createdAt, ready.completedAt])
  expect(end.text).toContain('验收')
  expect(end.text).toContain('PDF')
  expect(end.text).toContain('重开')
  expect(after.timeline.some((event) => event.type?.startsWith('project.completed'))).toBe(false)
  expect(after.nextAction).toBe(before.nextAction)
  expect(after.pages).toEqual(before.pages)
  expect(before.attention.some((item) => item.id === 'delivery-bundle-unfinished')).toBe(true)
  expect(after.attention.some((item) => item.id === 'delivery-bundle-unfinished')).toBe(false)
  expect(
    presentationWorkflowSummary(project, undefined, undefined, undefined, { bundles: [ready] })!
      .timeline,
  ).toEqual(after.timeline)
})
it('hides mismatched or malformed receipts and reports an unavailable history without carrying stale entries', () => {
  const ready = hostBundleReceipt()
  for (const bad of [
    { ...ready, requestId: 'other', manifest: { ...ready.manifest, requestId: 'other' } },
    { ...ready, projectId: 'other', manifest: { ...ready.manifest, projectId: 'other' } },
    { ...ready, manifest: { ...ready.manifest, planRevision: 2 } },
    { ...ready, completedAt: undefined },
  ]) {
    const result = presentationWorkflowSummary(project, undefined, undefined, undefined, {
      bundles: [bad],
    })!
    expect(
      result.timeline.some((event) => event.scope === 'current_document_delivery_bundle'),
    ).toBe(false)
    expect(result.attention.some((item) => item.id === 'delivery-bundle-history-unavailable')).toBe(
      true,
    )
  }
  const unavailable = presentationWorkflowSummary(project, undefined, undefined, undefined, {
    bundles: [ready],
    unavailable: true,
  })!
  expect(
    unavailable.timeline.some((event) => event.scope === 'current_document_delivery_bundle'),
  ).toBe(false)
  expect(
    presentationWorkflowSummary(project, undefined, undefined, undefined, { bundles: [] })!
      .timeline,
  ).toEqual(presentationWorkflowSummary(project, undefined, undefined)!.timeline)
})
