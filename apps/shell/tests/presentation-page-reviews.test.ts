import { PresentationStore } from '@wiswork/project-store'
import { createPresentationService } from '../src/main/presentation-service'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import { mkdtempSync, rmSync, readdirSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function setup() {
  const root = mkdtempSync(join(tmpdir(), 'page-reviews-'))
  roots.push(root)
  const store = new PresentationStore(root),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  plan.sources[0]!.uri = `attachment:${'c'.repeat(64)}`
  plan.sources.push({ ...plan.sources[0]!, id: 'second' })
  plan.claims[0]!.sourceIds.push('second')
  deck.claims = presentationPlanClaims(plan)
  store.savePlan(deck.id, 'doc', 0, plan)
  store.beginProduction(deck.id, 'doc', 'run', deck, { revision: 1, plan })
  const compile = vi.fn(),
    service = createPresentationService({ userDataPath: root, compile })
  const request = {
    operation: 'production_page_reviews',
    documentId: 'doc',
    projectId: deck.id,
    requestId: 'run',
    pageId: plan.slides[0]!.id,
  }
  const call = async (patch = {}, signal = new AbortController().signal) =>
    JSON.parse(Buffer.from(await service({ ...request, ...patch }, signal)).toString())
  const review = {
    pageId: request.pageId,
    claimId: plan.claims[0]!.id,
    sourceId: 'source',
    attachmentId: 'c'.repeat(64),
    offset: 0,
    maxChars: 10,
    evidenceDigest: 'd'.repeat(64),
    outcome: 'supported',
    notes: 'history',
    reviewer: 'agent',
  }
  return { root, store, plan, deck, call, review, compile, request }
}
const files = (root: string) =>
  readdirSync(root, { recursive: true, withFileTypes: true })
    .filter((e) => e.isFile())
    .map((e) => [
      join(e.parentPath, e.name),
      readFileSync(join(e.parentPath, e.name)).toString('base64'),
    ])
it('reads frozen complete historical source coverage without writes or attachment access', async () => {
  const f = setup()
  f.store.saveClaimReview(f.deck.id, 'doc', 'run', 'r1', f.review)
  f.store.saveClaimReview(f.deck.id, 'doc', 'run', 'r2', { ...f.review, outcome: 'contradicted' })
  f.store.saveClaimReview(f.deck.id, 'doc', 'run', 'other', {
    ...f.review,
    pageId: f.plan.slides[1]!.id,
  })
  f.plan.claims[0]!.sourceIds = ['source']
  f.store.savePlan(f.deck.id, 'doc', 1, f.plan)
  const before = files(f.root),
    result = await f.call()
  expect(result.planRevision).toBe(1)
  expect(result.claims[0].status).toBe('mixed')
  expect(result.claims[0].sources.map((s: { sourceId: string }) => s.sourceId)).toEqual([
    'source',
    'second',
  ])
  expect(result.claims[0].sources[0].reviews.map((r: { reviewId: string }) => r.reviewId)).toEqual([
    'r1',
    'r2',
  ])
  expect(files(f.root)).toEqual(before)
  expect(f.compile).not.toHaveBeenCalled()
})
it('requires exact request fields, rejects cross-document and missing identities, and supports cancellation', async () => {
  const f = setup()
  for (const key of Object.keys(f.request))
    expect(await f.call({ [key]: undefined })).toHaveProperty('error')
  expect(await f.call({ notes: 'extra' })).toEqual({ error: 'invalid_request' })
  expect(await f.call({ documentId: 'other' })).toHaveProperty('error')
  expect(await f.call({ requestId: 'missing' })).toEqual({ error: 'not_found' })
  expect(await f.call({ pageId: 'missing' })).toEqual({ error: 'not_found' })
  const controller = new AbortController()
  controller.abort()
  expect(await f.call({}, controller.signal)).toEqual({ error: 'aborted' })
})
it('rejects stored reviews whose selected-page source is outside the frozen claim', async () => {
  const f = setup()
  f.store.saveClaimReview(f.deck.id, 'doc', 'run', 'r1', { ...f.review, sourceId: 'missing' })
  expect(await f.call()).toEqual({ error: 'invalid_state' })
})
