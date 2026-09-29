import { expect, it } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { presentationHostAssociations } from '../src/skills/powerpoint/presentation-host-associations.js'
import type { PresentationImportRecord } from '../src/skills/powerpoint/presentation-delivery.js'

function receipt(documentId = 'doc'): PresentationImportRecord {
  return {
    state: 'complete',
    documentId,
    slideIds: ['host-a', 'host-b'],
    checkpoint: {
      version: 2,
      artifactDigest: 'a'.repeat(64),
      pageIds: benchmarkPlan()
        .slides.slice(0, 2)
        .map((p) => p.id),
      sourceSlideIds: ['256#', '256#'],
      baselineSlideIds: [],
      completed: [
        { sourceSlideId: '256#', slideId: 'host-a' },
        { sourceSlideId: '256#', slideId: 'host-b' },
      ],
    },
  }
}
it('associates exact page IDs across historical imports and distinguishes current presence without inferring content or QA', () => {
  const plan = benchmarkPlan()
  const result = presentationHostAssociations(
    plan,
    3,
    'doc',
    [
      { key: `production/${plan.projectId}/old`, record: receipt() },
      { key: `production/${plan.projectId}/foreign`, record: receipt('foreign') },
      {
        key: `${plan.projectId}/legacy`,
        record: { state: 'complete', documentId: 'doc', slideIds: ['legacy'] },
      },
    ],
    [{ requestId: 'old', planRevision: 1 }],
    ['host-a', 'manual', 'legacy'],
  )
  expect(result.pages[0]).toEqual({
    pageId: plan.slides[0]!.id,
    hostPages: [
      {
        requestId: 'old',
        slideId: 'host-a',
        planRevision: 1,
        revisionRelation: 'historical',
        presence: 'present',
        position: 1,
      },
    ],
  })
  expect(result.pages[1]!.hostPages[0]!.presence).toBe('missing')
  expect(result.pages[2]!.hostPages).toEqual([])
  expect(result.legacyImports).toBe(1)
  expect(result.uncertainImports).toBe(0)
})
it('keeps multiple copies, partial checkpoints and unknown revision identity explicit', () => {
  const plan = benchmarkPlan()
  const second = receipt()
  second.state = 'pending'
  delete second.slideIds
  second.checkpoint!.completed = [{ sourceSlideId: '256#', slideId: 'copy' }]
  second.checkpoint!.inFlight = { sourceSlideId: '256#' }
  const result = presentationHostAssociations(
    plan,
    3,
    'doc',
    [
      { key: `production/${plan.projectId}/first`, record: receipt() },
      { key: `production/${plan.projectId}/second`, record: second },
    ],
    [{ requestId: 'first', planRevision: 3 }],
    ['host-a', 'host-b', 'copy'],
  )
  expect(result.pages[0]!.hostPages.map((page) => page.revisionRelation)).toEqual([
    'current',
    'unknown',
  ])
  expect(result.pages[0]!.hostPages.map((page) => page.slideId)).toEqual(['host-a', 'copy'])
  expect(result.uncertainImports).toBe(1)
})
it('refuses malformed, duplicate or conflicting identities instead of guessing a mapping', () => {
  const plan = benchmarkPlan()
  const entries = [{ key: `production/${plan.projectId}/first`, record: receipt() }]
  expect(() =>
    presentationHostAssociations(plan, 3, 'doc', entries, [], ['host-a', 'host-a']),
  ).toThrow('presentation_host_association_invalid')
  expect(() =>
    presentationHostAssociations(
      plan,
      3,
      'doc',
      [...entries, { ...entries[0]!, key: `production/${plan.projectId}/second` }],
      [],
      ['host-a'],
    ),
  ).toThrow('presentation_host_association_invalid')
  const invalid = receipt()
  invalid.checkpoint!.pageIds = [plan.slides[0]!.id, plan.slides[0]!.id]
  expect(() =>
    presentationHostAssociations(plan, 3, 'doc', [{ ...entries[0]!, record: invalid }], [], []),
  ).toThrow('presentation_host_association_invalid')
})
