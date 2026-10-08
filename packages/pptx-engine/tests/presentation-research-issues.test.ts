import { createHash } from 'node:crypto'
import { describe, expect, it } from 'vitest'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import { researchFixture } from './fixtures/presentation-research'
import { benchmarkDeck } from './fixtures/presentation-benchmark'
import { presentationPlanClaims } from '../src/presentation-plan'
import {
  buildPresentationDeliveryReport,
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
  type PresentationDeliveryReportInput,
} from '../src/presentation-delivery-report'

function fixture(): PresentationDeliveryReportInput {
  const { plan, record } = researchFixture()
  record.draft.facts.push({
    ...record.draft.facts[0]!,
    claimId: 'conflict',
    statement: 'Opposing position',
    conflictsWith: ['original-claim'],
  })
  plan.claims.push({ ...plan.claims[0]!, id: 'opposition', statement: 'Opposing position' })
  plan.research!.claims.push({ claimId: 'opposition', researchClaimId: 'conflict' })
  plan.slides[0]!.claimIds.push('opposition')
  record.draftDigest = createHash('sha256')
    .update(canonicalPresentationValue(record.draft))
    .digest('hex')
  plan.research!.draftDigest = record.draftDigest
  const deck = benchmarkDeck()
  deck.claims = presentationPlanClaims(plan)
  deck.slides[0]!.claimIds = plan.slides[0]!.claimIds
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
    researchRecord: record,
    deck,
    metadata,
    reviews: [],
    pageStates: plan.slides.map((p) => ({ pageId: p.id, state: 'pending' })),
    issueLedger: {
      version: 1,
      projectId: plan.projectId,
      documentId: 'doc',
      requestId: 'run',
      inputDigest: metadata.inputDigest,
      planDigest: metadata.planDigest,
      revision: 0,
      actions: [],
    },
  }
}

describe('page-scoped research issues', () => {
  it('retains conflict issues even when both original claims are mapped', async () => {
    const report = await buildPresentationDeliveryReport(fixture())
    const issues = report.pages[0]!.issues.filter(
      (issue) => issue.code === 'research_claim_conflict',
    )
    expect(issues).toHaveLength(2)
    expect(issues[0]?.research?.relatedClaimIds).toEqual(['conflict'])
    expect(issues[1]?.research?.relatedClaimIds).toEqual(['original-claim'])
  })
  it('rejects forged issue research context', async () => {
    const report = await buildPresentationDeliveryReport(fixture())
    const issue = report.pages[0]!.issues.find((issue) => issue.code === 'research_claim_conflict')!
    issue.research!.relatedClaimIds = []
    expect(() => parsePresentationDeliveryReport(report)).toThrow()
  })
  it('groups multiple omitted sources and conflicts without losing original IDs', async () => {
    const input = fixture()
    input.plan.research!.claims = input.plan.research!.claims.filter(
      (mapping) => mapping.claimId !== 'opposition',
    )
    for (const id of ['unused-a', 'unused-b']) {
      input.researchRecord!.draft.sources.push({ ...input.researchRecord!.draft.sources[0]!, id })
      input.researchRecord!.sources!.push({
        sourceId: id,
        status: 'missing',
        provenance: 'unavailable',
      })
      input.researchRecord!.draft.facts[0]!.sourceRefs.push(id)
    }
    input.researchRecord!.draftDigest = createHash('sha256')
      .update(canonicalPresentationValue(input.researchRecord!.draft))
      .digest('hex')
    input.plan.research!.draftDigest = input.researchRecord!.draftDigest
    const report = await buildPresentationDeliveryReport(input)
    const issues = report.pages[0]!.issues
    const unselected = issues.filter(
      (issue) => issue.code === 'research_source_reference_unselected',
    )
    expect(unselected).toHaveLength(1)
    expect(unselected[0]?.research?.sourceIds).toEqual(['unused-a', 'unused-b'])
    expect(issues).toContainEqual(
      expect.objectContaining({
        code: 'research_conflict_partner_omitted',
        research: expect.objectContaining({ relatedClaimIds: ['conflict'] }),
      }),
    )
    expect(issues).toContainEqual(
      expect.objectContaining({
        code: 'research_unmapped_claim',
        claimId: 'opposition',
        category: 'unverifiable',
      }),
    )
    const markdown = presentationDeliveryMarkdown(report)
    expect(markdown).toContain('Frozen research issue context')
    expect(markdown).toContain('Opposing position')
  })
  it('keeps explained issues scoped to one page and opens stale dispositions after evidence changes', async () => {
    const input = fixture()
    const before = await buildPresentationDeliveryReport(input)
    const issue = before.pages[0]!.issues.find(
      (item) => item.code === 'research_claim_conflict' && item.claimId === 'source-1',
    )!
    input.issueLedger = {
      ...input.issueLedger,
      revision: 1,
      actions: [
        {
          actionId: 'explain',
          issueId: issue.id,
          issueDigest: issue.digest,
          state: 'explained',
          note: 'Human explanation',
          sequence: 1,
          createdAt: '2026-09-29T00:02:00.000Z',
        },
      ],
    }
    const explained = await buildPresentationDeliveryReport(input)
    expect(explained.pages[0]!.issues.find((item) => item.id === issue.id)?.disposition).toEqual({
      state: 'explained',
      stale: false,
      actionId: 'explain',
    })
    const otherPage = explained.pages
      .slice(1)
      .flatMap((page) => page.issues)
      .find((item) => item.claimId === 'source-1' && item.code === 'research_claim_conflict')!
    expect(otherPage.id).not.toBe(issue.id)
    expect(otherPage.disposition).toEqual({ state: 'open', stale: false })
    input.researchRecord!.sources![0]!.status = 'not_found'
    const changed = await buildPresentationDeliveryReport(input)
    expect(changed.pages[0]!.issues.find((item) => item.id === issue.id)?.disposition).toEqual({
      state: 'open',
      stale: true,
      actionId: 'explain',
    })
    expect(changed.checks.sourceAuthority).toBe('not_verified')
  })
  it('does not close research conflicts when a historical source review is supported', async () => {
    const input = fixture()
    const before = await buildPresentationDeliveryReport(input)
    const conflict = before.pages[0]!.issues.find(
      (issue) => issue.code === 'research_claim_conflict',
    )!
    input.reviews = [
      {
        version: 1,
        projectId: input.metadata.projectId,
        requestId: 'run',
        planRevision: 1,
        inputDigest: input.metadata.inputDigest,
        planDigest: input.metadata.planDigest,
        reviewId: 'supported',
        pageId: input.plan.slides[0]!.id,
        claimId: 'source-1',
        sourceId: 'source',
        attachmentId: 'b'.repeat(64),
        offset: 0,
        maxChars: 50,
        evidenceDigest: 'd'.repeat(64),
        outcome: 'supported',
        notes: 'Historical judgment',
        reviewer: 'agent',
        createdAt: '2026-09-29T00:02:00.000Z',
        checks: {
          support: 'agent_reviewed',
          sourceAuthority: 'not_verified',
          timeliness: 'not_verified',
          host: 'not_checked',
        },
      },
    ]
    const after = await buildPresentationDeliveryReport(input)
    const retained = after.pages[0]!.issues.find((issue) => issue.id === conflict.id)!
    expect(retained.digest).toBe(conflict.digest)
    expect(retained.disposition).toEqual({ state: 'open', stale: false })
  })
})
