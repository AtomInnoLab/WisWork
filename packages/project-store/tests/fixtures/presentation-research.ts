import type { PresentationResearchDraft } from '../../src/presentation-research'
export function researchDraft(attachmentId?: string): PresentationResearchDraft {
  return {
    scope: '研究范围：原文出现不等于事实成立。',
    sources: [
      {
        id: 'original',
        title: '原始材料',
        uri: attachmentId ? `attachment:${attachmentId}` : 'https://example.com/research',
        excerpt: '收入增长仅是管理层预测。',
        asOf: '2026-09-29',
      },
    ],
    facts: [
      {
        claimId: 'growth',
        statement: '收入增长仅是管理层预测。',
        type: 'fact',
        sourceRefs: ['original'],
        sourceTier: 'primary',
        slideRefs: ['suggested-slide'],
        confidence: 'high',
        reviewStatus: 'needs_review',
        conflictsWith: ['uncertain'],
      },
      {
        claimId: 'uncertain',
        statement: '增长假设仍存在不确定性。',
        type: 'assumption',
        sourceRefs: [],
        sourceTier: 'unverified',
        slideRefs: [],
        confidence: 'low',
        reviewStatus: 'needs_review',
        conflictsWith: ['growth'],
      },
    ],
  }
}
