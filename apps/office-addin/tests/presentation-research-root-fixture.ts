import { createHash } from 'node:crypto'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import type {
  PresentationResearchDraft,
  PresentationResearchRecord,
} from '@wiswork/project-store/presentation-research'
export function researchDraftFixture(): PresentationResearchDraft {
  return {
    scope: '科研汇报的原文结论与争议',
    sources: [
      {
        id: 'source1',
        title: '原始研究资料',
        uri: `attachment:${'a'.repeat(64)}`,
        excerpt: '原始研究存在限制',
      },
    ],
    facts: [
      {
        claimId: 'claim1',
        statement: '试验观察到效果',
        type: 'fact',
        sourceRefs: ['source1'],
        sourceTier: 'primary',
        slideRefs: ['proposed1'],
        confidence: 'medium',
        reviewStatus: 'needs_review',
        conflictsWith: ['claim2'],
      },
      {
        claimId: 'claim2',
        statement: '观察结果不足以推广',
        type: 'judgment',
        sourceRefs: ['source1'],
        sourceTier: 'unverified',
        slideRefs: ['proposed2'],
        confidence: 'low',
        reviewStatus: 'needs_review',
        conflictsWith: ['claim1'],
      },
    ],
  }
}
export function researchRecordFixture(
  documentId = 'doc',
  projectId = 'project',
): PresentationResearchRecord {
  const draft = researchDraftFixture()
  return {
    version: 1,
    documentId,
    projectId,
    id: 'research1',
    sequence: 1,
    draftDigest: createHash('sha256').update(canonicalPresentationValue(draft)).digest('hex'),
    draft,
    state: 'completed',
    startedAt: '2026-09-29T00:00:00.000Z',
    finishedAt: '2026-09-29T00:01:00.000Z',
    sources: [
      {
        sourceId: 'source1',
        attachmentId: 'a'.repeat(64),
        status: 'found',
        offset: 0,
        provenance: 'user_supplied',
        sha256: 'a'.repeat(64),
      },
    ],
    checks: {
      scope: 'research_draft',
      support: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
    },
  }
}
