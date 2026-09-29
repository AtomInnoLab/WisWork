import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import { readBoundPresentationResearch } from '../src/main/presentation-research-plan-binding'
import { handlePresentationDeliveryReport } from '../src/main/presentation-delivery-report'
import { mkdtempSync, rmSync, readdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationStore } from '@wiswork/project-store'
import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import type { PresentationResearchDraft } from '@wiswork/project-store/presentation-research'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
async function setup() {
  const root = mkdtempSync(join(tmpdir(), 'research-plan-'))
  roots.push(root)
  const plan = benchmarkPlan()
  plan.sources[0]!.uri = 'https://example.com/original'
  const store = new PresentationStore(root),
    research = new PresentationResearchStore(root)
  const source = plan.sources[0]!,
    claim = plan.claims[0]!
  const draft: PresentationResearchDraft = {
    scope: 'research before plan',
    sources: [{ ...source, id: 'research-source' }],
    facts: [
      {
        claimId: 'research-claim',
        statement: claim.statement,
        type: claim.type,
        sourceRefs: ['research-source'],
        sourceTier: 'unverified',
        slideRefs: [],
        confidence: claim.confidence,
        reviewStatus: 'needs_review',
        conflictsWith: [],
      },
    ],
  }
  const { record: running } = await research.begin('doc', plan.projectId, 0, 'ledger-a', draft)
  const record = await research.finish('doc', plan.projectId, running.id, {
    state: 'completed',
    sources: [{ sourceId: 'research-source', status: 'missing', provenance: 'unavailable' }],
  })
  const bound = {
    ...plan,
    research: {
      ledgerId: record.id,
      sequence: record.sequence,
      draftDigest: record.draftDigest,
      sources: [{ sourceId: source.id, researchSourceId: 'research-source' }],
      claims: [{ claimId: claim.id, researchClaimId: 'research-claim' }],
    },
  }
  const compile = vi.fn()
  const pc = createPresentationService({ userDataPath: root, compile })
  const call = async (operation: string, fields: Record<string, unknown> = {}) =>
    JSON.parse(
      Buffer.from(
        await pc(
          { operation, documentId: 'doc', projectId: plan.projectId, ...fields },
          new AbortController().signal,
        ),
      ).toString(),
    )
  return { root, plan, bound, record, store, research, call, compile }
}
it('saves the specified original research version and keeps lost-response CAS retry idempotent', async () => {
  const f = await setup()
  const first = await f.call('save_plan', { expectedRevision: 0, plan: f.bound })
  expect(first).toMatchObject({ revision: 1, plan: f.bound })
  expect(await f.call('save_plan', { expectedRevision: 0, plan: f.bound })).toEqual(first)
})

it('rejects changed original qualifiers or unknown versions before any plan write, and preserves the old plan', async () => {
  const f = await setup()
  for (const altered of [
    { ...f.bound, research: { ...f.bound.research, ledgerId: 'missing' } },
    { ...f.bound, research: { ...f.bound.research, sequence: 2 } },
    { ...f.bound, sources: [{ ...f.bound.sources[0], excerpt: 'changed qualifier' }] },
    {
      ...f.bound,
      claims: [{ ...f.bound.claims[0], statement: 'confirmed fact instead of assumption' }],
    },
  ]) {
    const result = await f.call('save_plan', { expectedRevision: 0, plan: altered })
    expect(['research_binding_invalid', 'research_unavailable']).toContain(result.error)
    expect(f.store.plan(f.plan.projectId, 'doc')).toBeUndefined()
  }
  await f.call('save_plan', { expectedRevision: 0, plan: f.bound })
  expect(
    await f.call('save_plan', { expectedRevision: 0, plan: { ...f.bound, title: 'different' } }),
  ).toEqual({ error: 'revision_conflict' })
  expect(f.store.plan(f.plan.projectId, 'doc')!.revision).toBe(1)
})
it('freezes original research for production and reports after newer research and plan edits, and restores that exact history', async () => {
  const f = await setup()
  expect(await f.call('save_plan', { expectedRevision: 0, plan: f.bound })).toMatchObject({
    revision: 1,
  })
  const deck = benchmarkPlannedDeck()
  deck.claims = presentationPlanClaims(f.bound)
  expect(
    await f.call('production_begin', { requestId: 'old-run', planRevision: 1, deck }),
  ).not.toHaveProperty('error')
  await f.research.begin('doc', f.plan.projectId, 2, 'ledger-b', {
    ...f.record.draft,
    scope: 'newer research',
  })
  await f.research.finish('doc', f.plan.projectId, 'ledger-b', {
    state: 'completed',
    sources: f.record.sources!,
  })
  expect(
    await f.call('save_plan', {
      expectedRevision: 1,
      plan: { ...f.bound, title: 'new plan still explicit A' },
    }),
  ).toMatchObject({ revision: 2 })
  const report = await f.call('production_delivery_report', { requestId: 'old-run' })
  expect(report.research.record).toEqual(f.record)
  expect(report.plan.title).toBe(f.bound.title)
  expect(report.research.findings).toContainEqual(
    expect.objectContaining({ code: 'source_unavailable' }),
  )
  const historical = await f.call('get_plan', { revision: 1 })
  expect(await f.call('save_plan', { expectedRevision: 2, plan: historical.plan })).toMatchObject({
    revision: 3,
    plan: f.bound,
  })
  expect(
    await f.call('production_begin', { requestId: 'old-run', planRevision: 1, deck }),
  ).not.toHaveProperty('error')
})
it('fails safely on damaged archives before new production and preserves existing plan and task history', async () => {
  const f = await setup()
  await f.call('save_plan', { expectedRevision: 0, plan: f.bound })
  const deck = benchmarkPlannedDeck()
  deck.claims = presentationPlanClaims(f.bound)
  await f.call('production_begin', { requestId: 'old-run', planRevision: 1, deck })
  const root = join(f.root, 'presentation-research')
  const doc = join(root, readdirSync(root)[0]!)
  const project = join(doc, readdirSync(doc)[0]!)
  writeFileSync(join(project, 'state.json'), 'private database error with credentials')
  expect(await f.call('production_begin', { requestId: 'new-run', planRevision: 1, deck })).toEqual(
    { error: 'research_unavailable' },
  )
  expect(f.store.production(f.plan.projectId, 'doc', 'new-run')).toBeUndefined()
  expect(await f.call('production_job_start', { requestId: 'old-run' })).toEqual({
    error: 'research_unavailable',
  })
  expect(f.store.productionJob(f.plan.projectId, 'doc', 'old-run')).toBeUndefined()
  expect(await f.call('production_delivery_report', { requestId: 'old-run' })).toEqual({
    error: 'research_unavailable',
  })
  expect(
    await f.call('save_plan', {
      expectedRevision: 1,
      plan: { ...f.bound, title: 'must not persist' },
    }),
  ).toEqual({ error: 'research_unavailable' })
  expect(f.store.plan(f.plan.projectId, 'doc')!.revision).toBe(1)
  expect(f.store.production(f.plan.projectId, 'doc', 'old-run')).toBeDefined()
})
it('keeps legacy report signatures and rejects absent callbacks or cross-document research for bound plans', async () => {
  const f = await setup()
  const deck = benchmarkPlannedDeck()
  deck.claims = presentationPlanClaims(f.plan)
  f.store.beginProduction(f.plan.projectId, 'doc', 'legacy', deck, { revision: 1, plan: f.plan })
  const legacy = await handlePresentationDeliveryReport(
    {
      projectId: f.plan.projectId,
      documentId: 'doc',
      requestId: 'legacy',
      operation: 'production_delivery_report',
    },
    f.store,
    async () => {
      throw new Error('not_found')
    },
    new AbortController().signal,
  )
  expect(legacy).not.toHaveProperty('research')
  await expect(readBoundPresentationResearch(f.bound, 'doc', f.plan.projectId)).rejects.toThrow(
    'research_unavailable',
  )
  await expect(
    readBoundPresentationResearch(f.bound, 'foreign', f.plan.projectId, async () => f.record),
  ).rejects.toThrow('research_binding_invalid')
  const controller = new AbortController()
  controller.abort()
  await expect(
    readBoundPresentationResearch(
      f.bound,
      'doc',
      f.plan.projectId,
      async () => f.record,
      controller.signal,
    ),
  ).rejects.toThrow('aborted')
})

it('refuses legacy resume of a frozen bound request when its original research archive is damaged', async () => {
  const f = await setup()
  await f.call('save_plan', { expectedRevision: 0, plan: f.bound })
  const deck = benchmarkPlannedDeck()
  deck.claims = presentationPlanClaims(f.bound)
  f.store.begin(f.plan.projectId, 'doc', 'legacy-resume', deck, { revision: 1, plan: f.bound })
  const archiveRoot = join(f.root, 'presentation-research')
  const doc = join(archiveRoot, readdirSync(archiveRoot)[0]!)
  const project = join(doc, readdirSync(doc)[0]!)
  writeFileSync(join(project, 'state.json'), 'damaged private archive')
  expect(await f.call('resume', { requestId: 'legacy-resume' })).toEqual({
    error: 'research_unavailable',
  })
  expect(f.compile).not.toHaveBeenCalled()
  expect(f.store.request(f.plan.projectId, 'doc', 'legacy-resume')!.status).toBe('pending')
})
