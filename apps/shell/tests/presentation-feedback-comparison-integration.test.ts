import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { PRESENTATION_DOMAIN_PROFILES } from '@wiswork/pptx-engine/presentation-plan'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationProjectController } from '../../office-addin/src/skills/powerpoint/presentation-project'
import { createPresentationService } from '../src/main/presentation-service'

async function fixture(
  domain: 'pitch' | 'report' | 'training' | 'research' | 'sales',
  differentBrief = false,
) {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-feedback-comparison-'))
  let service = createPresentationService({ userDataPath: root })
  const baseline = benchmarkPlan(),
    candidate = structuredClone(baseline)
  candidate.domain = domain
  candidate.title = `行业计划 ${domain}`
  if (differentBrief) candidate.brief.audience = '另一组受众'
  const sections = PRESENTATION_DOMAIN_PROFILES[domain].sections
  candidate.slides.forEach((page, index) => {
    page.domainSection = sections[index % sections.length]
  })
  const call = async (operation: string, extra: Record<string, unknown> = {}) =>
    JSON.parse(
      Buffer.from(
        await service(
          { operation, documentId: 'doc', projectId: baseline.projectId, ...extra },
          new AbortController().signal,
        ),
      ).toString('utf8'),
    )
  try {
    expect(await call('save_plan', { expectedRevision: 0, plan: baseline })).toMatchObject({
      revision: 1,
    })
    expect(
      await call('production_begin', {
        requestId: 'generic',
        planRevision: 1,
        deck: benchmarkPlannedDeck(),
      }),
    ).not.toHaveProperty('error')
    expect(await call('production_run', { requestId: 'generic' })).toMatchObject({
      status: 'compiled',
    })
    expect(await call('save_plan', { expectedRevision: 1, plan: candidate })).toMatchObject({
      revision: 2,
    })
    const deck = benchmarkPlannedDeck()
    deck.title = candidate.title
    expect(
      await call('production_begin', { requestId: 'domain', planRevision: 2, deck }),
    ).not.toHaveProperty('error')
    expect(await call('production_run', { requestId: 'domain' })).toMatchObject({
      status: 'compiled',
    })
    return {
      call,
      root,
      baseline,
      candidate,
      request: async (body: unknown, signal?: AbortSignal) =>
        new Response(await service(body, signal ?? new AbortController().signal)),
      restart: () => {
        service = createPresentationService({ userDataPath: root })
      },
      dispose: () => rmSync(root, { recursive: true, force: true }),
    }
  } catch (error) {
    rmSync(root, { recursive: true, force: true })
    throw error
  }
}

it.each(['pitch', 'report', 'training', 'research', 'sales'] as const)(
  'reads a %s pair from two actual compiled frozen versions after PC restart',
  async (domain) => {
    const f = await fixture(domain)
    try {
      for (const requestId of ['generic', 'domain'])
        expect(
          await f.call('production_feedback_record', {
            requestId,
            expectedRevision: 0,
            pages: f.baseline.slides.map((page, index) => ({
              pageId: page.id,
              status: requestId === 'generic' || index === 0 ? 'needs_correction' : 'no_correction',
            })),
          }),
        ).not.toHaveProperty('error')
      const before = await f.call('production_delivery_report', { requestId: 'domain' })
      const nextPlan = structuredClone(f.candidate)
      nextPlan.brief.audience = '后续新任务的受众'
      expect(await f.call('save_plan', { expectedRevision: 2, plan: nextPlan })).toMatchObject({
        revision: 3,
      })
      f.restart()
      const compared = await f.call('production_feedback_compare', {
        requestId: 'domain',
        baselineRequestId: 'generic',
      })
      expect(compared).not.toHaveProperty('error')
      expect(compared.comparison).toMatchObject({
        source: 'user_reported',
        effect: 'not_verified',
        comparable: true,
        gaps: [],
        delta: { needsCorrectionPages: 1 - f.baseline.slides.length },
      })
      expect(compared.comparison.baseline.plan).toEqual(f.baseline)
      expect(compared.comparison.candidate.plan).toEqual(f.candidate)
      expect(compared.comparison.baseline).toMatchObject({
        requestId: 'generic',
        planRevision: 1,
        feedbackRevision: 1,
      })
      expect(compared.comparison.candidate).toMatchObject({
        requestId: 'domain',
        planRevision: 2,
        feedbackRevision: 1,
      })
      expect(
        compared.comparison.conditions.every((condition: { match: boolean }) => condition.match),
      ).toBe(true)
      expect(compared.comparison.candidate.inputDigest).toBe(before.inputDigest)
      expect(compared.comparison.candidate.planDigest).toBe(before.planDigest)
      const after = await f.call('production_delivery_report', { requestId: 'domain' })
      expect(after.checks).toEqual(before.checks)
      expect(after.issueLedger).toEqual(before.issueLedger)
    } finally {
      f.dispose()
    }
  },
)

it('keeps absent and partial feedback unknown until both actual tasks are fully evaluated', async () => {
  const f = await fixture('report')
  try {
    const compare = () =>
      f.call('production_feedback_compare', { requestId: 'domain', baselineRequestId: 'generic' })
    const absent = await compare()
    expect(absent).not.toHaveProperty('error')
    expect(absent.comparison).toMatchObject({ comparable: false, delta: null })
    expect(absent.comparison.baseline).toMatchObject({
      feedbackRevision: null,
      feedbackRecordedAt: null,
      counts: {
        evaluatedPages: 0,
        notEvaluatedPages: f.baseline.slides.length,
        needsCorrectionRate: null,
      },
    })
    expect(
      await f.call('production_feedback_record', {
        requestId: 'generic',
        expectedRevision: 0,
        pages: f.baseline.slides.map((page) => ({ pageId: page.id, status: 'needs_correction' })),
      }),
    ).not.toHaveProperty('error')
    expect(
      await f.call('production_feedback_record', {
        requestId: 'domain',
        expectedRevision: 0,
        pages: [{ pageId: f.candidate.slides[0]!.id, status: 'no_correction' }],
      }),
    ).not.toHaveProperty('error')
    const partial = await compare()
    expect(partial.comparison).toMatchObject({
      comparable: false,
      delta: null,
      candidate: {
        feedbackRevision: 1,
        counts: { evaluatedPages: 1, notEvaluatedPages: f.candidate.slides.length - 1 },
      },
    })
    expect(
      await f.call('production_feedback_record', {
        requestId: 'domain',
        expectedRevision: 1,
        pages: f.candidate.slides
          .slice(1)
          .map((page) => ({ pageId: page.id, status: 'no_correction' })),
      }),
    ).not.toHaveProperty('error')
    const complete = await compare()
    expect(complete.comparison).toMatchObject({
      comparable: true,
      candidate: { feedbackRevision: 2 },
      delta: { needsCorrectionPages: -f.baseline.slides.length },
    })
    expect(partial.comparison.candidate.feedbackRevision).toBe(1)
    expect(partial.comparison.delta).toBeNull()
  } finally {
    f.dispose()
  }
})

it('refuses an improvement delta when real frozen task requirements differ', async () => {
  const f = await fixture('report', true)
  try {
    for (const requestId of ['generic', 'domain'])
      expect(
        await f.call('production_feedback_record', {
          requestId,
          expectedRevision: 0,
          pages: f.baseline.slides.map((page) => ({
            pageId: page.id,
            status: requestId === 'generic' ? 'needs_correction' : 'no_correction',
          })),
        }),
      ).not.toHaveProperty('error')
    const compared = await f.call('production_feedback_compare', {
      requestId: 'domain',
      baselineRequestId: 'generic',
    })
    expect(compared).not.toHaveProperty('error')
    expect(compared.comparison).toMatchObject({
      comparable: false,
      effect: 'not_verified',
      delta: null,
    })
    expect(
      compared.comparison.conditions.find((condition: { key: string }) => condition.key === 'brief')
        .match,
    ).toBe(false)
    expect(
      await f.call('production_feedback_compare', {
        requestId: 'generic',
        baselineRequestId: 'domain',
      }),
    ).toMatchObject({ comparison: { comparable: false, delta: null } })
    expect(
      await f.call('production_feedback_compare', {
        requestId: 'domain',
        baselineRequestId: 'domain',
      }),
    ).toHaveProperty('error')
    expect(
      await f.call('production_feedback_compare', {
        requestId: 'domain',
        baselineRequestId: 'generic',
        documentId: 'other',
      }),
    ).toHaveProperty('error')
  } finally {
    f.dispose()
  }
})

it('reads an immutable comparison through the actual Office controller and PC service without Agent tools', async () => {
  const f = await fixture('sales')
  try {
    for (const requestId of ['generic', 'domain'])
      expect(
        await f.call('production_feedback_record', {
          requestId,
          expectedRevision: 0,
          pages: f.baseline.slides.map((page) => ({ pageId: page.id, status: 'no_correction' })),
        }),
      ).not.toHaveProperty('error')
    let toolCalls = 0
    const controller = createPresentationProjectController({
      request: f.request,
      available: () => true,
      documentId: async () => 'doc',
      lastProject: () => f.baseline.projectId,
      executeTool: async () => {
        toolCalls++
        throw new Error('comparison must be read directly')
      },
    })
    await controller.refresh()
    controller.selectFeedbackComparisonBaseline!('generic')
    await controller.readFeedbackComparison!()
    expect(controller.snapshot().feedbackComparison).toMatchObject({
      baseline: { requestId: 'generic' },
      candidate: { requestId: 'domain' },
      comparable: true,
      effect: 'not_verified',
    })
    const saved = controller.snapshot().feedbackComparison
    f.restart()
    await controller.readFeedbackComparison!()
    expect(controller.snapshot().feedbackComparison).toEqual(saved)
    expect(toolCalls).toBe(0)
    controller.clear()
  } finally {
    f.dispose()
  }
})
