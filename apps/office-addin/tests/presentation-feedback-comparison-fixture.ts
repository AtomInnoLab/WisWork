import { vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { PRESENTATION_DOMAIN_PROFILES } from '@wiswork/pptx-engine/presentation-plan'
import { buildPresentationFeedbackComparison } from '@wiswork/pptx-engine/presentation-feedback-comparison'
import type { PresentationProductionFeedbackLedger } from '@wiswork/project-store/presentation-feedback'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
export const comparisonProduction = {
  projectId: 'p',
  requestId: 'candidate',
  planRevision: 2,
  status: 'compiled',
  compiledCount: 5,
  total: 5,
  pages: ['a', 'b', 'c', 'd', 'e'].map((id) => ({ id, title: id, state: 'compiled', attempt: 1 })),
}
export const comparisonProject = {
  projectId: 'p',
  title: '行业对照',
  status: 'compiled',
  latestRequestId: 'candidate',
  latestCompiledRequestId: 'candidate',
  slideCount: 5,
  slides: comparisonProduction.pages.map(({ id, title }) => ({ id, title })),
  history: [{ requestId: 'candidate', sequence: 3, status: 'compiled', slideCount: 5 }],
  production: comparisonProduction,
  productionTasks: [
    {
      requestId: 'candidate',
      sequence: 3,
      planRevision: 2,
      status: 'compiled',
      compiledCount: 5,
      total: 5,
    },
    {
      requestId: 'unfinished',
      sequence: 2,
      planRevision: 1,
      status: 'partial',
      compiledCount: 0,
      total: 2,
    },
    {
      requestId: 'baseline',
      sequence: 1,
      planRevision: 1,
      status: 'compiled',
      compiledCount: 2,
      total: 2,
    },
  ],
  checks: {
    structure: 'passed',
    geometry: 'passed',
    render: 'not_run',
    sources: 'not_verified',
    roundTrip: 'not_run',
  },
}
export function comparisonReport(
  options: {
    missing?: boolean
    differentStyle?: boolean
    baselineRevision?: number
    candidateRevision?: number
    candidateNeeds?: number
  } = {},
) {
  const baseline = benchmarkPlan()
  baseline.projectId = 'p'
  baseline.slides = baseline.slides
    .slice(0, 2)
    .map((slide, index) => ({ ...slide, id: ['a', 'b'][index]! }))
  const candidate = structuredClone(baseline)
  candidate.domain = 'report'
  candidate.slides = PRESENTATION_DOMAIN_PROFILES.report.sections.map((section, index) => ({
    ...baseline.slides[0]!,
    id: comparisonProduction.pages[index]!.id,
    domainSection: section,
  }))
  if (options.differentStyle) candidate.style.textColor = '123456'
  const task = (
    requestId: string,
    plan: typeof baseline,
    planRevision: number,
    revision: number,
    needs: number,
    missing = false,
  ) => {
    const inputDigest = (requestId === 'baseline' ? 'a' : 'c').repeat(64),
      planDigest = (requestId === 'baseline' ? 'b' : 'd').repeat(64)
    const feedback: PresentationProductionFeedbackLedger | null = missing
      ? null
      : {
          version: 1,
          source: 'user_reported',
          projectId: 'p',
          documentId: 'doc',
          requestId,
          inputDigest,
          planDigest,
          planRevision,
          pageIds: plan.slides.map((slide) => slide.id),
          revision,
          snapshots: Array.from({ length: revision }, (_, index) => ({
            revision: index + 1,
            recordedAt: `2026-09-29T00:00:${String(index).padStart(2, '0')}.000Z`,
            pages: plan.slides.map((slide, i) => ({
              pageId: slide.id,
              status: i < needs ? 'needs_correction' : 'no_correction',
            })),
          })),
        }
    return { requestId, inputDigest, planDigest, planRevision, plan, feedback }
  }
  return buildPresentationFeedbackComparison({
    projectId: 'p',
    documentId: 'doc',
    baseline: task('baseline', baseline, 1, options.baselineRevision ?? 1, 2),
    candidate: task(
      'candidate',
      candidate,
      2,
      options.candidateRevision ?? 1,
      options.candidateNeeds ?? 1,
      options.missing,
    ),
  })
}
export function comparisonFixture() {
  let comparison: unknown
  const request = vi.fn(async (body: unknown) => {
    const operation = (body as { operation: string }).operation
    return new Response(
      JSON.stringify(
        operation === 'status'
          ? comparisonProject
          : operation === 'production_feedback_compare' && comparison !== undefined
            ? { comparison }
            : { error: 'invalid_request' },
      ),
    )
  })
  const available = vi.fn(() => true),
    documentId = vi.fn(async () => 'doc'),
    executeTool = vi.fn(async () => ({ output: '{}', summary: '', mutated: false }))
  const controller = createPresentationProjectController({
    request,
    available,
    documentId,
    executeTool,
    lastProject: () => 'p',
  })
  return {
    controller,
    request,
    available,
    documentId,
    executeTool,
    setComparison: (value: unknown) => {
      comparison = value
    },
  }
}
