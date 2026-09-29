import { parsePresentationPlan, type PresentationPlan } from './presentation-plan'
import { canonicalPresentationValue as canonical } from '@wiswork/project-store/presentation-canonical'
import { parsePresentationProductionFeedbackLedger } from '@wiswork/project-store/presentation-feedback'
import type {
  PresentationProductionFeedbackLedger,
  PresentationProductionFeedbackStatus,
} from '@wiswork/project-store/presentation-feedback'
export const MAX_PRESENTATION_FEEDBACK_COMPARISON_BYTES = 1024 * 1024
export const MAX_PRESENTATION_FEEDBACK_COMPARISON_RESPONSE_BYTES =
  MAX_PRESENTATION_FEEDBACK_COMPARISON_BYTES +
  new TextEncoder().encode('{"comparison":}').byteLength
export const PRESENTATION_FEEDBACK_COMPARISON_GAPS = Object.freeze([
  'baseline_not_generic',
  'candidate_not_industry',
  'input_conditions_differ',
  'baseline_feedback_missing',
  'candidate_feedback_missing',
  'baseline_not_fully_evaluated',
  'candidate_not_fully_evaluated',
] as const)
export type PresentationFeedbackComparisonGap =
  (typeof PRESENTATION_FEEDBACK_COMPARISON_GAPS)[number]
export type PresentationFeedbackComparisonCondition =
  'brief' | 'sources' | 'claims' | 'research' | 'style' | 'brandKit' | 'parallelism'
export interface PresentationFeedbackComparisonObservation {
  requestId: string
  inputDigest: string
  planDigest: string
  planRevision: number
  plan: PresentationPlan
  feedbackRevision: number | null
  feedbackRecordedAt: string | null
  pages: { pageId: string; status: PresentationProductionFeedbackStatus }[]
  counts: {
    totalPages: number
    evaluatedPages: number
    needsCorrectionPages: number
    noCorrectionPages: number
    notEvaluatedPages: number
    needsCorrectionRate: number | null
  }
}
export interface PresentationFeedbackComparison {
  version: 1
  source: 'user_reported'
  effect: 'not_verified'
  projectId: string
  documentId: string
  baseline: PresentationFeedbackComparisonObservation
  candidate: PresentationFeedbackComparisonObservation
  conditions: { key: PresentationFeedbackComparisonCondition; match: boolean }[]
  comparable: boolean
  gaps: PresentationFeedbackComparisonGap[]
  delta: { needsCorrectionPages: number; needsCorrectionRate: number } | null
}
export interface PresentationFeedbackComparisonTaskInput {
  requestId: string
  inputDigest: string
  planDigest: string
  planRevision: number
  plan: PresentationPlan
  feedback: PresentationProductionFeedbackLedger | null
}
export interface PresentationFeedbackComparisonInput {
  projectId: string
  documentId: string
  baseline: PresentationFeedbackComparisonTaskInput
  candidate: PresentationFeedbackComparisonTaskInput
}

type SelectedObservation = Omit<PresentationFeedbackComparisonObservation, 'counts'>
const conditionKeys: PresentationFeedbackComparisonCondition[] = [
  'brief',
  'sources',
  'claims',
  'research',
  'style',
  'brandKit',
  'parallelism',
]
function invalid(): never {
  throw new Error('presentation_feedback_comparison_invalid')
}
function object(value: unknown, keys: string[]): Record<string, unknown> {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Object.prototype
  )
    invalid()
  const record = value as Record<string, unknown>
  if (
    Object.keys(record).some((key) => !keys.includes(key)) ||
    Object.values(Object.getOwnPropertyDescriptors(record)).some((d) => !('value' in d))
  )
    invalid()
  return record
}
const id = (v: unknown): v is string => typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
const digest = (v: unknown): v is string => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
function plan(value: unknown): PresentationPlan {
  try {
    return parsePresentationPlan(value)
  } catch {
    invalid()
  }
}
function time(value: unknown): value is string {
  return (
    typeof value === 'string' &&
    /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(value) &&
    Number.isFinite(Date.parse(value)) &&
    new Date(value).toISOString() === value
  )
}
function selected(value: unknown): SelectedObservation {
  const side = object(value, [
    'requestId',
    'inputDigest',
    'planDigest',
    'planRevision',
    'plan',
    'feedbackRevision',
    'feedbackRecordedAt',
    'pages',
    'counts',
  ])
  const frozen = plan(side.plan)
  if (
    !id(side.requestId) ||
    !digest(side.inputDigest) ||
    !digest(side.planDigest) ||
    !Number.isSafeInteger(side.planRevision) ||
    Number(side.planRevision) < 1 ||
    !(
      (side.feedbackRevision === null && side.feedbackRecordedAt === null) ||
      (Number.isSafeInteger(side.feedbackRevision) &&
        Number(side.feedbackRevision) >= 1 &&
        Number(side.feedbackRevision) <= 64 &&
        time(side.feedbackRecordedAt))
    ) ||
    !Array.isArray(side.pages) ||
    side.pages.length !== frozen.slides.length
  )
    invalid()
  const pages = side.pages.map((item, index) => {
    const page = object(item, ['pageId', 'status'])
    if (
      page.pageId !== frozen.slides[index]!.id ||
      !['needs_correction', 'no_correction', 'not_evaluated'].includes(page.status as string) ||
      (side.feedbackRevision === null && page.status !== 'not_evaluated')
    )
      invalid()
    return {
      pageId: page.pageId as string,
      status: page.status as PresentationProductionFeedbackStatus,
    }
  })
  if (Object.hasOwn(side, 'counts'))
    object(side.counts, [
      'totalPages',
      'evaluatedPages',
      'needsCorrectionPages',
      'noCorrectionPages',
      'notEvaluatedPages',
      'needsCorrectionRate',
    ])
  return {
    requestId: side.requestId,
    inputDigest: side.inputDigest,
    planDigest: side.planDigest,
    planRevision: side.planRevision as number,
    plan: frozen,
    feedbackRevision: side.feedbackRevision as number | null,
    feedbackRecordedAt: side.feedbackRecordedAt as string | null,
    pages,
  }
}
function counted(side: SelectedObservation): PresentationFeedbackComparisonObservation {
  const totalPages = side.pages.length,
    needsCorrectionPages = side.pages.filter((p) => p.status === 'needs_correction').length,
    noCorrectionPages = side.pages.filter((p) => p.status === 'no_correction').length,
    evaluatedPages = needsCorrectionPages + noCorrectionPages
  return {
    ...side,
    counts: {
      totalPages,
      evaluatedPages,
      needsCorrectionPages,
      noCorrectionPages,
      notEvaluatedPages: totalPages - evaluatedPages,
      needsCorrectionRate: evaluatedPages ? needsCorrectionPages / evaluatedPages : null,
    },
  }
}
function derive(
  projectId: unknown,
  documentId: unknown,
  b: SelectedObservation,
  c: SelectedObservation,
): PresentationFeedbackComparison {
  if (
    !id(projectId) ||
    typeof documentId !== 'string' ||
    !documentId.trim() ||
    documentId.length > 2048 ||
    b.plan.projectId !== projectId ||
    c.plan.projectId !== projectId ||
    b.requestId === c.requestId
  )
    invalid()
  const baseline = counted(b),
    candidate = counted(c)
  const conditions = conditionKeys.map((key) => ({
    key,
    match:
      canonical(key === 'parallelism' ? (b.plan.parallelism ?? 1) : b.plan[key]) ===
      canonical(key === 'parallelism' ? (c.plan.parallelism ?? 1) : c.plan[key]),
  }))
  const gaps: PresentationFeedbackComparisonGap[] = []
  if (Object.hasOwn(b.plan, 'domain')) gaps.push('baseline_not_generic')
  if (!['pitch', 'report', 'training', 'research', 'sales'].includes(c.plan.domain ?? ''))
    gaps.push('candidate_not_industry')
  if (conditions.some((condition) => !condition.match)) gaps.push('input_conditions_differ')
  if (b.feedbackRevision === null) gaps.push('baseline_feedback_missing')
  if (c.feedbackRevision === null) gaps.push('candidate_feedback_missing')
  if (baseline.counts.notEvaluatedPages) gaps.push('baseline_not_fully_evaluated')
  if (candidate.counts.notEvaluatedPages) gaps.push('candidate_not_fully_evaluated')
  const comparable = !gaps.length
  const result: PresentationFeedbackComparison = {
    version: 1,
    source: 'user_reported',
    effect: 'not_verified',
    projectId,
    documentId,
    baseline,
    candidate,
    conditions,
    comparable,
    gaps,
    delta: comparable
      ? {
          needsCorrectionPages:
            candidate.counts.needsCorrectionPages - baseline.counts.needsCorrectionPages,
          needsCorrectionRate:
            candidate.counts.needsCorrectionRate! - baseline.counts.needsCorrectionRate!,
        }
      : null,
  }
  if (
    new TextEncoder().encode(JSON.stringify(result)).byteLength >
    MAX_PRESENTATION_FEEDBACK_COMPARISON_BYTES
  )
    invalid()
  return result
}
function fromTask(value: unknown, projectId: string, documentId: string): SelectedObservation {
  const task = object(value, [
    'requestId',
    'inputDigest',
    'planDigest',
    'planRevision',
    'plan',
    'feedback',
  ])
  const frozen = plan(task.plan)
  let feedback: PresentationProductionFeedbackLedger | null = null
  if (task.feedback !== null) {
    try {
      feedback = parsePresentationProductionFeedbackLedger(task.feedback)
    } catch {
      invalid()
    }
  }
  if (
    feedback &&
    (feedback.projectId !== projectId ||
      feedback.documentId !== documentId ||
      feedback.requestId !== task.requestId ||
      feedback.inputDigest !== task.inputDigest ||
      feedback.planDigest !== task.planDigest ||
      feedback.planRevision !== task.planRevision ||
      canonical(feedback.pageIds) !== canonical(frozen.slides.map((s) => s.id)))
  )
    invalid()
  return selected({
    requestId: task.requestId,
    inputDigest: task.inputDigest,
    planDigest: task.planDigest,
    planRevision: task.planRevision,
    plan: frozen,
    feedbackRevision: feedback?.revision ?? null,
    feedbackRecordedAt: feedback?.snapshots.at(-1)?.recordedAt ?? null,
    pages: feedback
      ? feedback.snapshots.at(-1)!.pages.map(({ pageId, status }) => ({ pageId, status }))
      : frozen.slides.map((s) => ({ pageId: s.id, status: 'not_evaluated' })),
  })
}
/** Actual ledger history is validated before selecting its latest metadata-only snapshot. */
export function buildPresentationFeedbackComparison(
  input: PresentationFeedbackComparisonInput,
): PresentationFeedbackComparison {
  object(input, ['projectId', 'documentId', 'baseline', 'candidate'])
  return derive(
    input.projectId,
    input.documentId,
    fromTask(input.baseline, input.projectId, input.documentId),
    fromTask(input.candidate, input.projectId, input.documentId),
  )
}
/** Recompute only the selected observation included in this report; no absent history is invented. */
export function parsePresentationFeedbackComparison(
  value: unknown,
): PresentationFeedbackComparison {
  const report = object(value, [
    'version',
    'source',
    'effect',
    'projectId',
    'documentId',
    'baseline',
    'candidate',
    'conditions',
    'comparable',
    'gaps',
    'delta',
  ])
  try {
    if (
      new TextEncoder().encode(JSON.stringify(report)).byteLength >
      MAX_PRESENTATION_FEEDBACK_COMPARISON_BYTES
    )
      invalid()
  } catch {
    invalid()
  }
  const derived = derive(
    report.projectId,
    report.documentId,
    selected(report.baseline),
    selected(report.candidate),
  )
  if (canonical(report) !== canonical(derived)) invalid()
  return derived
}
