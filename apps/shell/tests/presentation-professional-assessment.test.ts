import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import type { PresentationProfessionalContext } from '@wiswork/project-store/presentation-professional-context'
import { presentationClaimEvidenceContent } from '@wiswork/pptx-engine/presentation-claim-review'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const hash = (v: string | Buffer) => createHash('sha256').update(v).digest('hex')
async function fixture(context?: PresentationProfessionalContext) {
  const root = mkdtempSync(join(tmpdir(), 'professional-assessment-'))
  roots.push(root)
  let service = createPresentationService({ userDataPath: root })
  const call = async (operation: string, fields: Record<string, unknown> = {}) =>
    JSON.parse(
      Buffer.from(
        await service({ operation, documentId: 'doc', ...fields }, new AbortController().signal),
      ).toString(),
    )
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck(),
    raw = Buffer.from('😀before 原文 after\f'),
    attachmentId = hash(raw)
  plan.sources[0]!.uri = `attachment:${attachmentId}`
  plan.sources[0]!.excerpt = '原文'
  if (context) plan.claims[0]!.professionalContext = context
  deck.claims = presentationPlanClaims(plan)
  expect(
    await call('attachment_begin', {
      attachmentId,
      sha256: attachmentId,
      name: 'synthetic.txt',
      sizeBytes: raw.length,
    }),
  ).not.toHaveProperty('error')
  await call('attachment_chunk', { attachmentId, offset: 0, base64: raw.toString('base64') })
  expect(await call('attachment_finish', { attachmentId })).not.toHaveProperty('error')
  expect(
    await call('save_plan', { projectId: plan.projectId, expectedRevision: 0, plan }),
  ).not.toHaveProperty('error')
  expect(
    await call('production_begin', {
      projectId: plan.projectId,
      requestId: 'run',
      planRevision: 1,
      deck,
    }),
  ).not.toHaveProperty('error')
  const request = {
    projectId: plan.projectId,
    requestId: 'run',
    pageId: deck.slides[0]!.id,
    claimId: 'source-1',
    sourceId: 'source',
    offset: 2,
    maxChars: 8000,
  }
  const evidence = await call('production_claim_evidence', request)
  expect(evidence).not.toHaveProperty('error')
  const review = {
    ...request,
    reviewId: 'opinion',
    evidenceDigest: hash(presentationClaimEvidenceContent(evidence)),
    outcome: 'supported',
    notes: 'Historical Agent opinion only',
  }
  const assessment = {
    scope: 'this frozen claim only',
    authority: { outcome: 'uncertain', sourceTier: 'unverified', reason: 'unverified' },
    timeliness: {
      outcome: 'uncertain',
      referenceDate: '2026-09-29',
      ...(plan.claims[0]!.asOf ? { claimAsOf: plan.claims[0]!.asOf } : {}),
      ...(plan.sources[0]!.asOf ? { sourceAsOf: plan.sources[0]!.asOf } : {}),
      reason: 'historical',
    },
    basis: [{ offset: 9, text: '原文' }],
    ...(context
      ? {
          professional: {
            context,
            checks: (context.domain === 'finance'
              ? ['comparability', 'forecast']
              : ['conclusion_scope', 'qualifications']
            ).map((aspect) => ({
              aspect,
              outcome: 'consistent',
              reason: 'declared-context comparison',
            })),
          },
        }
      : {}),
  }
  return {
    call,
    review,
    assessment,
    restart: () => {
      service = createPresentationService({ userDataPath: root })
    },
  }
}
it.each(['science', 'law', 'finance'] as const)(
  'persists %s historical professional opinions immutably across retry and restart',
  async (domain) => {
    const context: PresentationProfessionalContext =
      domain === 'science'
        ? { domain, sample: 'specific sample', limitations: 'limited' }
        : domain === 'law'
          ? { domain, jurisdiction: 'declared jurisdiction', applicabilityDate: '2026-09-29' }
          : { domain, currency: 'CNY', unit: '万元', asOf: '2026-09-29' }
    const f = await fixture(context),
      fields = { ...f.review, sourceAssessment: f.assessment }
    const result = await f.call('production_record_claim_review', fields)
    expect(result.sourceAssessment).toEqual(f.assessment)
    expect(result.checks).toMatchObject({
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
    })
    expect(await f.call('production_record_claim_review', fields)).toEqual(result)
    f.restart()
    expect(
      await f.call('production_read_claim_review', {
        projectId: f.review.projectId,
        requestId: 'run',
        reviewId: 'opinion',
      }),
    ).toEqual(result)
  },
)
it('rejects full context deletion, substitution and domain changes before writing; preserves actual UTF16 window rules', async () => {
  const f = await fixture({
    domain: 'finance',
    currency: 'CNY',
    unit: '万元',
    asOf: '2026-09-29',
    limitations: 'original limitation',
  })
  const p = f.assessment.professional!
  for (const sourceAssessment of [
    { ...f.assessment, professional: { ...p, context: { domain: 'finance', currency: 'CNY' } } },
    { ...f.assessment, professional: { ...p, context: { ...p.context, unit: '元' } } },
    {
      ...f.assessment,
      professional: {
        context: { domain: 'science' },
        checks: [
          { aspect: 'conclusion_scope', outcome: 'consistent', reason: 'opinion' },
          { aspect: 'qualifications', outcome: 'consistent', reason: 'opinion' },
        ],
      },
    },
    { ...f.assessment, basis: [] },
    { ...f.assessment, basis: [{ offset: 8, text: '原文' }] },
    { ...f.assessment, basis: [{ offset: 1, text: '😀' }] },
  ])
    expect(
      await f.call('production_record_claim_review', { ...f.review, sourceAssessment }),
    ).toEqual({ error: 'invalid_request' })
  expect(
    await f.call('production_read_claim_review', {
      projectId: f.review.projectId,
      requestId: 'run',
      reviewId: 'opinion',
    }),
  ).toEqual({ error: 'not_found' })
  expect(
    await f.call('production_record_claim_review', { ...f.review, sourceAssessment: f.assessment }),
  ).not.toHaveProperty('error')
})
it('rejects professional opinions without a frozen context and retains old assessment/read shape', async () => {
  const f = await fixture()
  const professional = {
    context: { domain: 'science' },
    checks: [
      { aspect: 'conclusion_scope', outcome: 'uncertain', reason: 'unknown' },
      { aspect: 'qualifications', outcome: 'uncertain', reason: 'unknown' },
    ],
  }
  expect(
    await f.call('production_record_claim_review', {
      ...f.review,
      sourceAssessment: { ...f.assessment, professional },
    }),
  ).toEqual({ error: 'invalid_request' })
  const result = await f.call('production_record_claim_review', {
    ...f.review,
    sourceAssessment: f.assessment,
  })
  expect(result.sourceAssessment).toEqual(f.assessment)
  expect(result.sourceAssessment).not.toHaveProperty('professional')
  f.restart()
  expect(
    await f.call('production_read_claim_review', {
      projectId: f.review.projectId,
      requestId: 'run',
      reviewId: 'opinion',
    }),
  ).toEqual(result)
})
