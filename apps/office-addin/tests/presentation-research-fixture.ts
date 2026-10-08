import { createHash } from 'node:crypto'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import type {
  PresentationResearchDraft,
  PresentationResearchRecord,
  PresentationResearchSummary,
} from '@wiswork/project-store/presentation-research'
export const researchChecks = {
  scope: 'research_draft',
  support: 'not_verified',
  sourceAuthority: 'not_verified',
  timeliness: 'not_verified',
} as const
export function researchRecord(): PresentationResearchRecord {
  const draft: PresentationResearchDraft = {
    scope: '销售趋势研究',
    sources: [
      {
        id: 'source1',
        title: '报告原文',
        uri: 'https://example.com/report',
        excerpt: '原文',
        snapshotAttachmentId: 'b'.repeat(64),
      },
    ],
    facts: ['claim1', 'claim2'].map((claimId) => ({
      claimId,
      statement: claimId === 'claim1' ? '销售增长' : '销售下降',
      type: claimId === 'claim1' ? 'fact' : 'judgment',
      sourceRefs: ['source1'],
      sourceTier: 'primary',
      slideRefs: [],
      confidence: 'high',
      reviewStatus: 'needs_review',
      conflictsWith: [claimId === 'claim1' ? 'claim2' : 'claim1'],
    })),
  }
  return {
    version: 1,
    documentId: 'doc',
    projectId: 'research',
    id: 'ledger1',
    sequence: 1,
    draftDigest: createHash('sha256').update(canonicalPresentationValue(draft)).digest('hex'),
    state: 'completed',
    startedAt: '2026-09-29T00:00:00.000Z',
    finishedAt: '2026-09-29T00:00:01.000Z',
    checks: researchChecks,
    draft,
    sources: [{ sourceId: 'source1', status: 'missing', provenance: 'unavailable' }],
  }
}
export function researchSummary(): PresentationResearchSummary {
  const record = researchRecord()
  return {
    version: 1,
    documentId: 'doc',
    projectId: 'research',
    revision: 2,
    totalRecords: 1,
    records: [
      {
        id: record.id,
        sequence: 1,
        draftDigest: record.draftDigest,
        state: record.state,
        startedAt: record.startedAt,
        finishedAt: record.finishedAt,
        sourceCount: 1,
        factCount: 2,
        conflictCount: 1,
      },
    ],
  }
}
