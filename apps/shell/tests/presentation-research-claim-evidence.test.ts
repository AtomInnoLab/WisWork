import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync, readdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { PresentationStore } from '@wiswork/project-store'
import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import type { PresentationResearchDraft } from '@wiswork/project-store/presentation-research'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import { presentationClaimEvidenceContent } from '@wiswork/pptx-engine/presentation-claim-review'
import { parsePresentationClaimEvidence } from '@wiswork/pptx-engine/presentation-claim-evidence'
import { createPresentationService } from '../src/main/presentation-service'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const hash = (v: string | Uint8Array) => createHash('sha256').update(v).digest('hex')
async function setup(bound = true) {
  const root = mkdtempSync(join(tmpdir(), 'research-claim-'))
  roots.push(root)
  let pc = createPresentationService({ userDataPath: root })
  const plan = benchmarkPlan()
  const text = '示例数据仅用于测试。并非事实认证。',
    id = hash(text)
  plan.sources[0]!.uri = `attachment:${id}`
  delete plan.sources[0]!.locator
  plan.sources[0]!.asOf = '2026-09-29'
  plan.claims[0]!.asOf = '2026-09-29'
  plan.claims[0]!.jurisdiction = '中国大陆，仅合成数据，不提供专业结论'
  const base = { documentId: 'doc', projectId: plan.projectId }
  const raw = async (body: Record<string, unknown>, signal = new AbortController().signal) =>
    JSON.parse(
      Buffer.from(
        await pc(
          {
            ...(String(body.operation).startsWith('attachment_') ? { documentId: 'doc' } : base),
            ...body,
          },
          signal,
        ),
      ).toString(),
    )
  await raw({
    operation: 'attachment_begin',
    attachmentId: id,
    sha256: id,
    name: 'source.txt',
    sizeBytes: Buffer.byteLength(text),
  })
  await raw({
    operation: 'attachment_chunk',
    attachmentId: id,
    offset: 0,
    base64: Buffer.from(text).toString('base64'),
  })
  expect(await raw({ operation: 'attachment_finish', attachmentId: id })).toMatchObject({
    status: 'ready',
  })
  const draft: PresentationResearchDraft = {
    scope: '原文只证明字面出现。研究整理完成不是事实支持。',
    sources: [{ ...plan.sources[0]!, id: 'research-source' }],
    facts: [
      {
        claimId: 'research-claim',
        statement: plan.claims[0]!.statement,
        type: plan.claims[0]!.type,
        asOf: plan.claims[0]!.asOf,
        jurisdiction: plan.claims[0]!.jurisdiction,
        sourceRefs: ['research-source'],
        sourceTier: 'primary',
        slideRefs: [],
        confidence: 'high',
        reviewStatus: 'needs_review',
        conflictsWith: ['opposite'],
      },
      {
        claimId: 'opposite',
        statement: '相反观点仍需保留，不允许将限定改成确定事实。',
        type: 'judgment',
        sourceRefs: [],
        sourceTier: 'unverified',
        slideRefs: [],
        confidence: 'low',
        reviewStatus: 'needs_review',
        conflictsWith: ['research-claim'],
      },
    ],
  }
  const result = await raw({
    operation: 'research_build',
    ledgerId: 'research-A',
    expectedRevision: 0,
    draft,
  })
  const record = result.record
  if (bound)
    plan.research = {
      ledgerId: record.id,
      sequence: record.sequence,
      draftDigest: record.draftDigest,
      sources: [{ sourceId: plan.sources[0]!.id, researchSourceId: 'research-source' }],
      claims: [{ claimId: plan.claims[0]!.id, researchClaimId: 'research-claim' }],
    }
  expect(await raw({ operation: 'save_plan', expectedRevision: 0, plan })).not.toHaveProperty(
    'error',
  )
  const deck = benchmarkPlannedDeck()
  deck.claims = presentationPlanClaims(plan)
  expect(
    await raw({ operation: 'production_begin', requestId: 'frozen', planRevision: 1, deck }),
  ).not.toHaveProperty('error')
  const evidenceRequest = {
    operation: 'production_claim_evidence',
    requestId: 'frozen',
    pageId: plan.slides.find((page) => page.claimIds.includes(plan.claims[0]!.id))!.id,
    claimId: plan.claims[0]!.id,
    sourceId: plan.sources[0]!.id,
    offset: 0,
    maxChars: 8000,
  }
  const store = new PresentationStore(root),
    research = new PresentationResearchStore(root)
  return {
    root,
    plan,
    record,
    draft,
    raw,
    evidenceRequest,
    store,
    research,
    restart: () => {
      pc = createPresentationService({ userDataPath: root })
    },
  }
}
it('returns exact frozen A full claim/research findings after newer B and preserves professional qualifiers', async () => {
  const f = await setup()
  const response = await f.raw(f.evidenceRequest)
  expect(response).not.toHaveProperty('error')
  const evidence = parsePresentationClaimEvidence(response)
  expect(evidence).toMatchObject({
    documentId: 'doc',
    claim: f.plan.claims[0],
    source: { asOf: '2026-09-29' },
    research: { binding: f.plan.research, record: f.record },
  })
  expect(evidence.research!.findings).toContainEqual({
    code: 'omitted_conflict_partner',
    claimId: f.plan.claims[0]!.id,
    researchClaimId: 'research-claim',
    relatedResearchClaimId: 'opposite',
  })
  const newer = await f.raw({
    operation: 'research_build',
    ledgerId: 'research-B',
    expectedRevision: 2,
    draft: { ...f.draft, scope: 'newer research B' },
  })
  expect(newer.record.id).toBe('research-B')
  f.restart()
  expect(await f.raw(f.evidenceRequest)).toEqual(evidence)
})
it('rejects legacy/context-tampered digests before supported review and retains the new whole-context digest after restart', async () => {
  const f = await setup()
  const response = await f.raw(f.evidenceRequest)
  expect(response).not.toHaveProperty('error')
  const evidence = parsePresentationClaimEvidence(response)
  const legacy = { ...evidence, source: { ...evidence.source } }
  delete legacy.documentId
  delete legacy.claim
  delete legacy.research
  delete legacy.source.asOf
  const review = {
    ...f.evidenceRequest,
    operation: 'production_record_claim_review',
    reviewId: 'review-A',
    outcome: 'supported',
    notes: '原文窗口已读；限定、冲突与来源权威/时效仍需进一步复核。',
  }
  const legacyDigest = hash(presentationClaimEvidenceContent(legacy))
  expect(await f.raw({ ...review, evidenceDigest: legacyDigest })).toEqual({
    error: 'evidence_changed',
  })
  expect(f.store.claimReview(f.plan.projectId, 'doc', 'frozen', 'review-A')).toBeUndefined()
  const digest = hash(presentationClaimEvidenceContent(evidence))
  const result = await f.raw({ ...review, evidenceDigest: digest })
  expect(result).toMatchObject({
    reviewId: 'review-A',
    evidenceDigest: digest,
    outcome: 'supported',
  })
  f.restart()
  expect(await f.raw({ ...review, evidenceDigest: digest })).toEqual(result)
  const changed = structuredClone(evidence)
  changed.claim!.jurisdiction = '错误新限定'
  expect(() => presentationClaimEvidenceContent(changed)).toThrow()
})
it('does not write a review when the exact archive is damaged or request is cancelled', async () => {
  const f = await setup()
  const evidence = await f.raw(f.evidenceRequest),
    digest = hash(presentationClaimEvidenceContent(evidence))
  const review = {
    ...f.evidenceRequest,
    operation: 'production_record_claim_review',
    reviewId: 'bad',
    outcome: 'supported',
    notes: '读过窗口',
    evidenceDigest: digest,
  }
  const controller = new AbortController()
  controller.abort()
  expect(await f.raw(review, controller.signal)).toEqual({ error: 'aborted' })
  expect(f.store.claimReview(f.plan.projectId, 'doc', 'frozen', 'bad')).toBeUndefined()
  const root = join(f.root, 'presentation-research'),
    doc = join(root, readdirSync(root)[0]!),
    project = join(doc, readdirSync(doc)[0]!)
  writeFileSync(join(project, 'state.json'), 'sensitive failed archive')
  expect(await f.raw(review)).toEqual({ error: 'research_unavailable' })
  expect(f.store.claimReview(f.plan.projectId, 'doc', 'frozen', 'bad')).toBeUndefined()
})
it('preserves exact legacy unbound evidence fields despite stored source and claim qualifiers', async () => {
  const f = await setup(false),
    evidence = await f.raw(f.evidenceRequest)
  expect(evidence).not.toHaveProperty('documentId')
  expect(evidence).not.toHaveProperty('claim')
  expect(evidence).not.toHaveProperty('research')
  expect(evidence.source).not.toHaveProperty('asOf')
  expect(parsePresentationClaimEvidence(evidence)).toEqual(evidence)
})

it('saves actual source assessment with literal basis and frozen labels; restart/lost ACK is immutable', async () => {
  const f = await setup()
  const evidence = parsePresentationClaimEvidence(await f.raw(f.evidenceRequest))
  const sourceAssessment = {
    scope: '仅此冻结主张',
    authority: {
      outcome: 'appropriate_for_claim',
      sourceTier: 'primary',
      reason: 'Agent语义判断，不是认证',
    },
    timeliness: {
      outcome: 'current_for_claim',
      referenceDate: '2026-09-29',
      claimAsOf: f.plan.claims[0]!.asOf,
      sourceAsOf: f.plan.sources[0]!.asOf,
      reason: '按冻结标签比较',
    },
    jurisdiction: {
      claimJurisdiction: f.plan.claims[0]!.jurisdiction,
      outcome: 'applicable',
      reason: '范围判断',
    },
    basis: [{ offset: 0, text: evidence.attachment.text.slice(0, 5) }],
  }
  const request = {
    ...f.evidenceRequest,
    operation: 'production_record_claim_review',
    reviewId: 'assessed',
    outcome: 'supported',
    notes: '意见',
    evidenceDigest: hash(presentationClaimEvidenceContent(evidence)),
    sourceAssessment,
  }
  const receipt = await f.raw(request)
  expect(receipt).not.toHaveProperty('error')
  expect(receipt.sourceAssessment).toEqual(sourceAssessment)
  expect(receipt.checks).toMatchObject({
    sourceAuthority: 'not_verified',
    timeliness: 'not_verified',
  })
  f.restart()
  expect(await f.raw(request)).toEqual(receipt)
  expect(
    await f.raw({ ...request, sourceAssessment: { ...sourceAssessment, scope: '不同' } }),
  ).toEqual({ error: 'request_conflict' })
  for (const change of [
    { basis: [{ offset: 1, text: '伪造' }] },
    { timeliness: { ...sourceAssessment.timeliness, claimAsOf: '伪造' } },
    { jurisdiction: { ...sourceAssessment.jurisdiction, claimJurisdiction: '其他' } },
  ]) {
    expect(
      await f.raw({
        ...request,
        reviewId: 'invalid',
        sourceAssessment: { ...sourceAssessment, ...change },
      }),
    ).toEqual({ error: 'invalid_request' })
    expect(f.store.claimReview(f.plan.projectId, 'doc', 'frozen', 'invalid')).toBeUndefined()
  }
})
it('checks assessment labels against unbound frozen plan rather than evidence optional context', async () => {
  const f = await setup(false)
  const evidence = parsePresentationClaimEvidence(await f.raw(f.evidenceRequest))
  const sourceAssessment = {
    scope: '无标签也不可猜测',
    authority: { outcome: 'uncertain', sourceTier: 'unverified', reason: '未验证' },
    timeliness: {
      outcome: 'uncertain',
      referenceDate: '2026-09-29',
      claimAsOf: '猜测',
      reason: '未验证',
    },
    basis: [],
  }
  expect(
    await f.raw({
      ...f.evidenceRequest,
      operation: 'production_record_claim_review',
      reviewId: 'bad-label',
      outcome: 'insufficient_evidence',
      notes: '意见',
      evidenceDigest: hash(presentationClaimEvidenceContent(evidence)),
      sourceAssessment,
    }),
  ).toEqual({ error: 'invalid_request' })
})
