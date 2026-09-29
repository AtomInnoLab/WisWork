import { benchmarkPlan } from './presentation-plan'
import { benchmarkDeck } from './presentation-benchmark'
import { presentationPlanClaims } from '../../src/presentation-plan'
import type { PresentationDeliveryReportInput } from '../../src/presentation-delivery-report'
import type { PresentationProfessionalContext } from '@wiswork/project-store/presentation-professional-context'
import type { PresentationClaimReview } from '../../src/presentation-claim-review'
export function professionalAssessmentInput(
  context: PresentationProfessionalContext = {
    domain: 'science',
    sample: '样本仅100人',
    method: '随机对照',
    limitations: '仅适用于声明样本',
  },
): PresentationDeliveryReportInput {
  const plan = benchmarkPlan(),
    deck = benchmarkDeck()
  plan.sources[0]!.uri = `attachment:${'a'.repeat(64)}`
  plan.claims[0]!.professionalContext = context
  deck.claims = presentationPlanClaims(plan)
  const metadata = {
    projectId: plan.projectId,
    documentId: 'doc',
    requestId: 'run',
    planRevision: 1,
    inputDigest: 'b'.repeat(64),
    planDigest: 'c'.repeat(64),
  }
  return {
    plan,
    deck,
    metadata,
    reviews: [],
    pageStates: plan.slides.map((page) => ({ pageId: page.id, state: 'pending' })),
    issueLedger: {
      version: 1,
      projectId: metadata.projectId,
      documentId: metadata.documentId,
      requestId: metadata.requestId,
      inputDigest: metadata.inputDigest,
      planDigest: metadata.planDigest,
      revision: 0,
      actions: [],
    },
  }
}
export function professionalAssessmentReview(
  input: PresentationDeliveryReportInput,
  reviewId = 'professional-review',
  outcomes = ['conflict', 'uncertain'],
): PresentationClaimReview {
  const context = input.plan.claims[0]!.professionalContext!
  return {
    version: 1,
    projectId: input.plan.projectId,
    requestId: 'run',
    planRevision: 1,
    inputDigest: input.metadata.inputDigest,
    planDigest: input.metadata.planDigest,
    reviewId,
    pageId: input.plan.slides[0]!.id,
    claimId: input.plan.claims[0]!.id,
    sourceId: input.plan.sources[0]!.id,
    attachmentId: 'a'.repeat(64),
    offset: 0,
    maxChars: 50,
    evidenceDigest: 'd'.repeat(64),
    outcome: 'supported',
    notes: '历史Agent意见',
    reviewer: 'agent',
    createdAt: '2026-09-29T00:00:00.000Z',
    sourceAssessment: {
      scope: '该页该来源原窗口',
      authority: { outcome: 'uncertain', sourceTier: 'unverified', reason: '尚待认证' },
      timeliness: { outcome: 'uncertain', referenceDate: '2026-09-29', reason: '时效尚待核验' },
      basis: [{ offset: 0, text: '原文逐字依据' }],
      professional: {
        context: structuredClone(context),
        checks: (context.domain === 'finance'
          ? ['comparability', 'forecast']
          : ['conclusion_scope', 'qualifications']
        ).map((aspect, index) => ({
          aspect,
          outcome: outcomes[index],
          reason: index === 0 ? '原材料不足以覆盖该范围' : '原限定需保留',
        })),
      },
    },
    checks: {
      support: 'agent_reviewed',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      host: 'not_checked',
    },
  } as PresentationClaimReview
}
