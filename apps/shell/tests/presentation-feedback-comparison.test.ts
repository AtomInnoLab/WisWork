import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { PresentationStore } from '@wiswork/project-store'
import { createPresentationService } from '../src/main/presentation-service'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { PRESENTATION_DOMAIN_PROFILES } from '@wiswork/pptx-engine/presentation-plan'
import { parsePresentationFeedbackComparison } from '@wiswork/pptx-engine/presentation-feedback-comparison'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
it('requires exact private scope and both actual persisted tasks compiled; comparison is read-only', async () => {
  const root = mkdtempSync(join(tmpdir(), 'feedback-compare-'))
  roots.push(root)
  const store = new PresentationStore(root),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const industry = structuredClone(plan)
  industry.domain = 'pitch'
  industry.slides.forEach(
    (page, index) => (page.domainSection = PRESENTATION_DOMAIN_PROFILES.pitch.sections[index % 5]),
  )
  let base = store.beginProduction(plan.projectId, 'doc', 'base', deck, { revision: 1, plan }),
    candidate = store.beginProduction(plan.projectId, 'doc', 'candidate', deck, {
      revision: 2,
      plan: industry,
    })
  const service = createPresentationService({ userDataPath: root })
  const body = {
    operation: 'production_feedback_compare',
    projectId: plan.projectId,
    documentId: 'doc',
    baselineRequestId: 'base',
    requestId: 'candidate',
  }
  const call = async (extra: Record<string, unknown> = {}, signal = new AbortController().signal) =>
    JSON.parse(Buffer.from(await service({ ...body, ...extra }, signal)).toString('utf8'))
  expect(await call()).toEqual({ error: 'page_not_ready' })
  for (const extra of [
    { baselineRequestId: 'candidate' },
    { baselineRequestId: '../base' },
    { requestId: '../candidate' },
    { source: 'user_reported' },
    { expectedRevision: 0 },
    { pages: [] },
  ])
    expect(await call(extra)).toEqual({ error: 'invalid_request' })
  expect(await call({ documentId: 'foreign' })).toEqual({ error: 'document_mismatch' })
  expect(await call({ baselineRequestId: 'missing' })).toEqual({ error: 'not_found' })
  for (const page of base.pages) {
    base = store.updateProductionPage(base, page.pageId, { state: 'building', attempt: 1 })
    base = store.updateProductionPage(base, page.pageId, {
      state: 'compiled',
      attempt: 1,
      result: { pptxBase64: 'UEsDBAAAAAA=', sourceSlideId: '256#', report: {} },
    })
  }
  expect(await call()).toEqual({ error: 'page_not_ready' })
  for (const page of candidate.pages) {
    candidate = store.updateProductionPage(candidate, page.pageId, {
      state: 'building',
      attempt: 1,
    })
    candidate = store.updateProductionPage(candidate, page.pageId, {
      state: 'compiled',
      attempt: 1,
      result: { pptxBase64: 'UEsDBAAAAAA=', sourceSlideId: '256#', report: {} },
    })
  }
  const path = join(
    root,
    'projects',
    'presentations',
    createHash('sha256').update(plan.projectId).digest('hex'),
  )
  const before = readdirSync(path).sort()
  const response = await call()
  expect(Object.keys(response)).toEqual(['comparison'])
  const report = parsePresentationFeedbackComparison(response.comparison)
  expect(report.delta).toBeNull()
  expect(report.gaps).toContain('baseline_feedback_missing')
  expect(report.gaps).toContain('candidate_feedback_missing')
  expect(report.baseline.plan).toEqual(plan)
  expect(report.candidate.plan).toEqual(industry)
  const abort = new AbortController()
  abort.abort()
  expect(await call({}, abort.signal)).toEqual({ error: 'aborted' })
  expect(readdirSync(path).sort()).toEqual(before)
  expect(store.production(plan.projectId, 'doc', 'base')).toEqual(base)
  expect(store.production(plan.projectId, 'doc', 'candidate')).toEqual(candidate)
  expect(store.productionFeedback(plan.projectId, 'doc', 'base')).toBeUndefined()
})
