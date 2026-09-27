import { describe, expect, it } from 'vitest'
import {
  buildPresentationDeliveryReport,
  parsePresentationDeliveryReport,
  presentationDeliveryMarkdown,
  type PresentationDeliveryReportInput,
} from '../src/presentation-delivery-report'
import { parsePresentationPlan, presentationPlanClaims } from '../src/presentation-plan'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { benchmarkDeck } from './fixtures/presentation-benchmark'
import type { PresentationClaimReview } from '../src/presentation-claim-review'
function fixture(): PresentationDeliveryReportInput {
  const plan = benchmarkPlan(),
    deck = benchmarkDeck()
  plan.sources[0]!.uri = `attachment:${'a'.repeat(64)}`
  plan.claims[0]!.type = 'calculation'
  plan.claims[0]!.calculation = {
    formula: 'a + 2',
    inputs: ['a input'],
    reproduction: {
      bindings: [{ name: 'a', inputIndex: 0, value: 1, sourceId: 'source' }],
      expected: 3,
    },
  }
  deck.claims = presentationPlanClaims(plan)
  const metadata = {
    projectId: plan.projectId,
    documentId: 'doc',
    requestId: 'req',
    planRevision: 1,
    inputDigest: 'b'.repeat(64),
    planDigest: 'c'.repeat(64),
  }
  return {
    plan,
    deck,
    metadata,
    pageStates: plan.slides.map((page) => ({ pageId: page.id, state: 'pending' })),
    reviews: [],
    issueLedger: {
      version: 1,
      ...metadata,
      revision: 0,
      actions: [],
    } as PresentationDeliveryReportInput['issueLedger'],
  }
}
function input(): PresentationDeliveryReportInput {
  const value = fixture()
  delete (value.issueLedger as unknown as Record<string, unknown>).planRevision
  return value
}
function review(
  value: PresentationDeliveryReportInput,
  id: string,
  outcome: PresentationClaimReview['outcome'],
): PresentationClaimReview {
  return {
    version: 1,
    ...value.metadata,
    reviewId: id,
    pageId: value.plan.slides[0]!.id,
    claimId: 'source-1',
    sourceId: 'source',
    attachmentId: 'a'.repeat(64),
    offset: 0,
    maxChars: 50,
    evidenceDigest: 'd'.repeat(64),
    outcome,
    notes: 'Historical agent judgment',
    reviewer: 'agent',
    createdAt: '2026-09-24T00:00:00.000Z',
    checks: {
      support: 'agent_reviewed',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      host: 'not_checked',
    },
  } as PresentationClaimReview
}
describe('delivery evidence report', () => {
  it('makes source dispositions stale when current attachment evidence disappears', async () => {
    const value = input()
    value.sourceAudit = [
      { sourceId: 'source', attachmentId: 'a'.repeat(64), status: 'found', offset: 0 },
    ]
    const before = await buildPresentationDeliveryReport(value)
    const issue = before.pages[0]!.issues.find((item) => item.code === 'source_review_missing')!
    value.issueLedger.actions = [
      {
        actionId: 'source-action',
        issueId: issue.id,
        issueDigest: issue.digest,
        state: 'explained',
        note: 'Original text was present',
        sequence: 1,
        createdAt: '2026-09-24T00:00:00.000Z',
      },
    ]
    value.issueLedger.revision = 1
    expect(
      (await buildPresentationDeliveryReport(value)).pages[0]!.issues.find(
        (item) => item.id === issue.id,
      )?.disposition.state,
    ).toBe('explained')
    value.sourceAudit = [{ sourceId: 'source', attachmentId: 'a'.repeat(64), status: 'missing' }]
    const after = await buildPresentationDeliveryReport(value)
    expect(after.pages[0]!.issues.find((item) => item.id === issue.id)?.disposition).toMatchObject({
      state: 'open',
      stale: true,
    })
    expect(after.pages[0]!.issues.some((item) => item.code === 'source_attachment_missing')).toBe(
      true,
    )
    const forged = structuredClone(after)
    forged.sourceAudit![0]!.attachmentId = 'b'.repeat(64)
    expect(() => parsePresentationDeliveryReport(forged)).toThrow()
  })
  it('covers all frozen pages and preserves all input and calculation evidence', async () => {
    const value = input(),
      report = await buildPresentationDeliveryReport(value)
    expect(report.pages).toHaveLength(value.plan.slides.length)
    expect(report.plan).toEqual(value.plan)
    expect(report.pages[0]!.calculations[0]!.status).toBe('reproduced')
    expect(
      report.pages
        .flatMap((page) => page.issues)
        .some((issue) => issue.code === 'calculation_not_reproduced'),
    ).toBe(false)
    expect(report.checks.host).toBe('not_checked')
    expect(parsePresentationDeliveryReport(report)).toEqual(report)
  })
  it('rejects forged arithmetic, machine passed states, missing issues and bad source attribution', async () => {
    const original = await buildPresentationDeliveryReport(input())
    const alterations = [
      (r: typeof original) => {
        r.pages[0]!.calculations[0]!.actual = 9
      },
      (r: typeof original) => {
        ;(r.checks as unknown as Record<string, string>).content = 'passed'
      },
      (r: typeof original) => {
        r.pages[0]!.issues = []
      },
      (r: typeof original) => {
        r.pages[0]!.issues[0]!.sourceId = 'missing'
      },
      (r: typeof original) => {
        r.pages[0]!.issues[0]!.disposition.state = 'explained'
      },
      (r: typeof original) => {
        r.pages.reverse()
      },
    ]
    for (const alter of alterations) {
      const report = structuredClone(original)
      alter(report)
      expect(() => parsePresentationDeliveryReport(report)).toThrow()
    }
  })
  it('keeps stale disposition history and limits review digest invalidation to relevant pages', async () => {
    const value = input()
    value.reviews = [review(value, 'r1', 'contradicted')]
    // review metadata deliberately carries no documentId in its strict schema
    delete (value.reviews[0] as unknown as Record<string, unknown>).documentId
    const before = await buildPresentationDeliveryReport(value)
    const issue = before.pages[0]!.issues.find(
      (issue) => issue.code === 'source_review_contradicted',
    )!
    value.issueLedger.actions = [
      {
        actionId: 'action',
        issueId: issue.id,
        issueDigest: issue.digest,
        state: 'explained',
        note: 'Review later',
        sequence: 1,
        createdAt: '2026-09-24T00:00:00.000Z',
      },
    ]
    value.issueLedger.revision = 1
    expect(
      (await buildPresentationDeliveryReport(value)).pages[0]!.issues.find(
        (item) => item.id === issue.id,
      )!.disposition.state,
    ).toBe('explained')
    const second = review(value, 'r2', 'contradicted')
    delete (second as unknown as Record<string, unknown>).documentId
    value.reviews.push(second)
    const after = await buildPresentationDeliveryReport(value)
    expect(after.pages[0]!.issues.find((item) => item.id === issue.id)!.disposition).toEqual({
      state: 'open',
      stale: true,
      actionId: 'action',
    })
    expect(after.pages[1]!.issues).toEqual(before.pages[1]!.issues)
    expect(after.issueLedger.actions).toHaveLength(1)
  })
  it('shows source and disposition text as inert text with all warnings', async () => {
    const value = input()
    value.plan.sources[0]!.excerpt = '<img src=x onerror=alert(1)> [click](javascript:alert(1))'
    value.plan.sources[0]!.title = '`danger`'
    const report = await buildPresentationDeliveryReport(value),
      markdown = presentationDeliveryMarkdown(report)
    expect(markdown).not.toContain('<img')
    expect(markdown).not.toContain('[click]')
    expect(markdown).toContain('&#60;img')
    expect(markdown).toContain('NOT VERIFIED')
    expect(markdown).toContain('unverifiable')
    expect(markdown).toContain('All disposition history')
  })
  it('keeps arithmetic disposition current when a source review changes and accepts host document identities', async () => {
    const value = input()
    value.metadata.documentId = '文件 / deck.pptx'
    value.issueLedger.documentId = value.metadata.documentId
    value.plan.claims[0]!.calculation!.reproduction!.expected = 4
    const before = await buildPresentationDeliveryReport(value)
    const mismatch = before.pages[0]!.issues.find((issue) => issue.code === 'calculation_mismatch')!
    value.issueLedger.actions = [
      {
        actionId: 'math',
        issueId: mismatch.id,
        issueDigest: mismatch.digest,
        state: 'deferred',
        note: 'Check declared inputs',
        sequence: 1,
        createdAt: '2026-09-24T00:00:00.000Z',
      },
    ]
    value.issueLedger.revision = 1
    const sourceReview = review(value, 'new', 'contradicted')
    delete (sourceReview as unknown as Record<string, unknown>).documentId
    value.reviews.push(sourceReview)
    const after = await buildPresentationDeliveryReport(value)
    expect(after.pages[0]!.issues.find((issue) => issue.id === mismatch.id)).toEqual({
      ...mismatch,
      disposition: { state: 'deferred', stale: false, actionId: 'math' },
    })
  })
  it('reports every issue in the bounded worst-case page without truncation', async () => {
    const value = input()
    value.plan.sources = ['s1', 's2', 's3'].map((id) => ({
      id,
      title: id,
      uri: 'urn:test',
      excerpt: '',
    }))
    value.plan.claims = Array.from({ length: 32 }, (_, index) => ({
      id: `c${index}`,
      statement: `Missing statement ${index}`,
      type: 'calculation' as const,
      sourceIds: ['s1', 's2', 's3'],
      confidence: 'low' as const,
      reviewStatus: 'needs_review' as const,
      asOf: '2026',
      calculation: { formula: '1+2', inputs: ['one', 'two'] },
    }))
    value.plan.slides = [
      { ...value.plan.slides[0]!, claimIds: value.plan.claims.map((claim) => claim.id) },
    ]
    value.deck.slides = [
      {
        ...value.deck.slides[0]!,
        claimIds: value.plan.slides[0]!.claimIds,
        elements: [{ id: 'shape', kind: 'shape', shape: 'rect', x: 1, y: 1, w: 1, h: 1 }],
      },
    ]
    value.deck.claims = presentationPlanClaims(value.plan)
    value.pageStates = [value.pageStates[0]!]
    const report = await buildPresentationDeliveryReport(value)
    expect(report.pages[0]!.issues).toHaveLength(448)
    expect(new Set(report.pages[0]!.issues.map((issue) => issue.id)).size).toBe(448)
    expect(presentationDeliveryMarkdown(report).match(/digest /g)).toHaveLength(448)
  })
  it('preserves contradictory review history and rejects invalid review ownership', async () => {
    const value = input()
    const first = review(value, 'first', 'supported'),
      second = review(value, 'second', 'contradicted')
    delete (first as unknown as Record<string, unknown>).documentId
    delete (second as unknown as Record<string, unknown>).documentId
    value.reviews = [first, second]
    const report = await buildPresentationDeliveryReport(value)
    expect(report.reviews).toHaveLength(2)
    expect(report.pages[0]!.issues.some((issue) => issue.code === 'source_review_mixed')).toBe(true)
    for (const changes of [
      { pageId: 'unknown' },
      { sourceId: 'unknown' },
      { planRevision: 2 },
      { attachmentId: 'f'.repeat(64) },
    ]) {
      const forged = structuredClone(report)
      Object.assign(forged.reviews[0]!, changes)
      expect(() => parsePresentationDeliveryReport(forged)).toThrow()
    }
    report.reviews.push(report.reviews[0]!)
    expect(() => parsePresentationDeliveryReport(report)).toThrow()
  })
  it('deduplicates shared source and claim bodies within the attachment byte budget', async () => {
    const value = input()
    value.plan.sources = ['s1', 's2', 's3'].map((id) => ({
      id,
      title: id,
      uri: 'urn:test',
      excerpt: 'EvidenceBody' + 'x'.repeat(11988),
      locator: 'page 1',
    }))
    value.plan.claims = Array.from({ length: 32 }, (_, index) => ({
      id: `c${index}`,
      statement: `ClaimBody${index}` + 'y'.repeat(900),
      type: 'assumption' as const,
      sourceIds: ['s1', 's2', 's3'],
      confidence: 'low' as const,
      reviewStatus: 'needs_review' as const,
    }))
    value.plan.slides = Array.from({ length: 32 }, (_, index) => ({
      ...value.plan.slides[0]!,
      id: `page${index}`,
      claimIds: value.plan.claims.map((claim) => claim.id),
    }))
    value.deck.slides = value.plan.slides.map((slide) => ({
      id: slide.id,
      title: slide.title,
      claimIds: slide.claimIds,
      elements: [{ id: 'shape', kind: 'shape', shape: 'rect', x: 1, y: 1, w: 1, h: 1 }],
    }))
    value.deck.claims = presentationPlanClaims(value.plan)
    value.pageStates = value.plan.slides.map((page) => ({ pageId: page.id, state: 'pending' }))
    const report = await buildPresentationDeliveryReport(value)
    const markdown = presentationDeliveryMarkdown(report)
    expect(new TextEncoder().encode(markdown).byteLength).toBeLessThan(20 * 1024 * 1024)
    // The human-readable catalogs and the complete frozen-plan appendix each retain a copy.
    expect(markdown.match(/EvidenceBody/g)).toHaveLength(6)
    expect(markdown.match(/ClaimBody0y/g)).toHaveLength(2)
    expect(markdown.match(/## page[0-9]+ /g)).toHaveLength(32)
    expect(markdown).toContain('All source review history')
    expect(markdown).toContain('All disposition history')
    expect(markdown).toContain('NOT VERIFIED')
    for (const source of value.plan.sources) source.excerpt = '<'.repeat(12000)
    for (const claim of value.plan.claims) claim.statement = '<'.repeat(900)
    value.deck.claims = presentationPlanClaims(value.plan)
    const escapedMarkdown = presentationDeliveryMarkdown(
      await buildPresentationDeliveryReport(value),
    )
    expect(new TextEncoder().encode(escapedMarkdown).byteLength).toBeLessThan(20 * 1024 * 1024)
    expect(escapedMarkdown).not.toContain('<')
    expect(escapedMarkdown).toContain('&#60;'.repeat(12000))
  }, 20000)
  it('bounds report bytes, page issue counts and exact states', async () => {
    const report = await buildPresentationDeliveryReport(input())
    const oversized = structuredClone(report)
    oversized.plan.title = 'x'.repeat(8 * 1024 * 1024)
    expect(() => parsePresentationDeliveryReport(oversized)).toThrow()
    const crowded = structuredClone(report)
    crowded.pages[0]!.issues = Array.from({ length: 513 }, () => crowded.pages[0]!.issues[0]!)
    expect(() => parsePresentationDeliveryReport(crowded)).toThrow()
    const state = structuredClone(report)
    ;(state.pages[0] as unknown as Record<string, unknown>).productionState = 'passed'
    expect(() => parsePresentationDeliveryReport(state)).toThrow()
  })
  it('validates reproduction binding source/index/name and finite bounds while retaining old plans', () => {
    const value = input().plan
    for (const change of [
      { inputIndex: 0.5 },
      { inputIndex: 1 },
      { sourceId: 'other' },
      { name: 'constructor' },
      { value: Infinity },
      { value: 1e13 },
    ]) {
      const plan = structuredClone(value)
      Object.assign(plan.claims[0]!.calculation!.reproduction!.bindings[0]!, change)
      expect(() => parsePresentationPlan(plan)).toThrow()
    }
    delete value.claims[0]!.calculation!.reproduction
    expect(parsePresentationPlan(value)).toEqual(value)
  })
})
