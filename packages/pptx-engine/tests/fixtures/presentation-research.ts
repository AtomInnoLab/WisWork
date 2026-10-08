import type { PresentationResearchRecord } from '@wiswork/project-store/presentation-research'
import { benchmarkPlan } from './presentation-plan'

export function researchFixture() {
  const plan = benchmarkPlan()
  plan.sources[0]!.uri = `attachment:${'b'.repeat(64)}`
  const record: PresentationResearchRecord = {
    version: 1,
    documentId: 'doc',
    projectId: plan.projectId,
    id: 'research-A',
    sequence: 1,
    draftDigest: 'a'.repeat(64),
    state: 'completed',
    startedAt: '2026-09-29T00:00:00.000Z',
    finishedAt: '2026-09-29T00:01:00.000Z',
    draft: {
      scope: 'research',
      sources: [{ ...plan.sources[0]!, id: 'original-source' }],
      facts: [
        {
          claimId: 'original-claim',
          statement: plan.claims[0]!.statement,
          type: 'assumption',
          sourceRefs: ['original-source'],
          sourceTier: 'unverified',
          slideRefs: [],
          confidence: 'low',
          reviewStatus: 'needs_review',
          conflictsWith: [],
        },
      ],
    },
    sources: [{ sourceId: 'original-source', status: 'missing', provenance: 'unavailable' }],
    checks: {
      scope: 'research_draft',
      support: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
    },
  }
  plan.research = {
    ledgerId: record.id,
    sequence: record.sequence,
    draftDigest: record.draftDigest,
    sources: [{ sourceId: 'source', researchSourceId: 'original-source' }],
    claims: [{ claimId: 'source-1', researchClaimId: 'original-claim' }],
  }
  return { plan, record }
}
