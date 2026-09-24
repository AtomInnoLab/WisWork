import { expect, it } from 'vitest'
import {
  parsePresentationPageReviews,
  summarizePresentationPageReviews,
} from '../src/presentation-page-reviews'
import { benchmarkPlan, benchmarkPlannedDeck } from './fixtures/presentation-plan'
const metadata = {
  projectId: benchmarkPlan().projectId,
  requestId: 'run',
  pageId: benchmarkPlan().slides[0]!.id,
  planRevision: 1,
  inputDigest: 'a'.repeat(64),
  planDigest: 'b'.repeat(64),
}
it('retains every unreviewed source and rejects forged status and extra fields', () => {
  const report = summarizePresentationPageReviews(
    benchmarkPlan(),
    benchmarkPlannedDeck(),
    metadata,
    [],
  )
  expect(report.claims[0]!.sources).toHaveLength(1)
  expect(report.claims[0]!.status).toBe('unreviewed')
  expect(() => parsePresentationPageReviews({ ...report, extra: true })).toThrow()
  report.claims[0]!.status = 'supported'
  expect(() => parsePresentationPageReviews(report)).toThrow()
})

import { presentationPlanClaims } from '../src/presentation-plan'
import type { PresentationClaimReview } from '../src/presentation-claim-review'
function fixture() {
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  plan.sources[0]!.uri = `attachment:${'c'.repeat(64)}`
  plan.sources.push({ ...plan.sources[0]!, id: 'second' })
  plan.claims[0]!.sourceIds.push('second')
  deck.claims = presentationPlanClaims(plan)
  const review: PresentationClaimReview = {
    ...metadata,
    version: 1,
    reviewId: 'r1',
    claimId: 'source-1',
    sourceId: 'source',
    attachmentId: 'c'.repeat(64),
    offset: 0,
    maxChars: 100,
    evidenceDigest: 'd'.repeat(64),
    outcome: 'supported',
    notes: 'history',
    reviewer: 'agent',
    createdAt: '2026-01-01T00:00:00.000Z',
    checks: {
      support: 'agent_reviewed',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      host: 'not_checked',
    },
  }
  return {
    plan,
    deck,
    review,
    summarize: (reviews: PresentationClaimReview[]) =>
      summarizePresentationPageReviews(plan, deck, metadata, reviews),
  }
}
it('preserves older reviews and reports partial and mixed without inferring truth', () => {
  const f = fixture()
  const partial = f.summarize([f.review])
  expect(partial.claims[0]!.status).toBe('partial')
  expect(partial.claims[0]!.sources[1]!.status).toBe('unreviewed')
  const mixed = f.summarize([
    f.review,
    { ...f.review, reviewId: 'r2', outcome: 'contradicted', offset: 100 },
  ])
  expect(mixed.claims[0]!.status).toBe('mixed')
  expect(mixed.claims[0]!.sources[0]!.reviews.map((r) => r.reviewId)).toEqual(['r1', 'r2'])
  expect(mixed.claims[0]!.sources[0]!.reviews[0]).not.toHaveProperty('notes')
  expect(
    f.summarize([
      f.review,
      { ...f.review, reviewId: 'r2', sourceId: 'second', outcome: 'insufficient_evidence' },
    ]).claims[0]!.status,
  ).toBe('mixed')
  expect(
    f.summarize([f.review, { ...f.review, reviewId: 'r2', sourceId: 'second' }]).claims[0]!.status,
  ).toBe('supported')
})
it('isolates other pages and rejects selected-page invalid membership and frozen identity', () => {
  const f = fixture()
  expect(f.summarize([{ ...f.review, pageId: f.plan.slides[1]!.id }]).claims[0]!.status).toBe(
    'unreviewed',
  )
  for (const patch of [
    { claimId: 'missing' },
    { sourceId: 'missing' },
    { requestId: 'other' },
    { attachmentId: 'e'.repeat(64) },
  ])
    expect(() => f.summarize([{ ...f.review, ...patch }])).toThrow('invalid_state')
  f.plan.claims[0]!.sourceIds = []
  f.deck.claims = presentationPlanClaims(f.plan)
  expect(f.summarize([]).claims[0]!.status).toBe('no_sources')
})
it('rejects duplicate IDs, forged distributions, invalid times and fractional windows', () => {
  const f = fixture(),
    original = f.summarize([f.review])
  const mutations = [
    (r: typeof original) => {
      r.claims.push(r.claims[0]!)
    },
    (r: typeof original) => {
      r.claims[0]!.sources.push(r.claims[0]!.sources[0]!)
    },
    (r: typeof original) => {
      r.claims[0]!.sources[0]!.reviews.push(r.claims[0]!.sources[0]!.reviews[0]!)
    },
    (r: typeof original) => {
      r.claims[0]!.sources[0]!.status = 'mixed'
    },
    (r: typeof original) => {
      r.claims[0]!.sources[0]!.reviews[0]!.createdAt = '2026-02-30T00:00:00.000Z'
    },
    (r: typeof original) => {
      r.claims[0]!.sources[0]!.reviews[0]!.offset = 0.5
    },
  ]
  for (const mutate of mutations) {
    const report = structuredClone(original)
    mutate(report)
    expect(() => parsePresentationPageReviews(report)).toThrow()
  }
})

it('rejects metadata from a different project even without historical reviews', () => {
  const f = fixture()
  expect(() =>
    summarizePresentationPageReviews(f.plan, f.deck, { ...metadata, projectId: 'other' }, []),
  ).toThrow('invalid_state')
})

it('enforces global review quota across sources and strict reference fields', () => {
  const f = fixture()
  expect(() =>
    f.summarize(
      Array.from({ length: 33 }, (_, i) => ({
        ...f.review,
        reviewId: `r${i}`,
        sourceId: i % 2 ? 'second' : 'source',
      })),
    ),
  ).toThrow()
  const report = f.summarize([f.review])
  for (const patch of [
    { notes: 'must not leak' },
    { outcome: 'verified' },
    { evidenceDigest: 'bad' },
    { maxChars: 8001 },
    { offset: -1 },
    { reviewId: '../bad' },
  ]) {
    const copy = structuredClone(report)
    Object.assign(copy.claims[0]!.sources[0]!.reviews[0]!, patch)
    expect(() => parsePresentationPageReviews(copy)).toThrow()
  }
  expect(() =>
    parsePresentationPageReviews({
      ...report,
      claims: Array.from({ length: 33 }, (_, i) => ({
        claimId: `c${i}`,
        status: 'no_sources',
        sources: [],
      })),
    }),
  ).toThrow()
})
