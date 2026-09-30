import {
  parsePresentationFeedbackComparison,
  MAX_PRESENTATION_FEEDBACK_COMPARISON_RESPONSE_BYTES,
  type PresentationFeedbackComparison,
} from '@wiswork/pptx-engine/presentation-feedback-comparison'
import {
  parsePresentationProductionFeedbackLedger,
  parsePresentationProductionFeedbackPages,
  MAX_PRESENTATION_PRODUCTION_FEEDBACK_BYTES,
  type PresentationProductionFeedbackLedger,
  type PresentationProductionFeedbackPage,
} from '@wiswork/project-store/presentation-feedback'
import {
  parsePresentationDeliveryBundleReceipt,
  type PresentationDeliveryBundleReceipt,
} from '@wiswork/project-store/presentation-delivery-bundle'
import {
  parsePresentationResearchSummary,
  type PresentationResearchSummary,
} from '@wiswork/project-store/presentation-research'
import {
  parsePresentationAssetLedger,
  type PresentationAssetLedger,
} from '@wiswork/project-store/presentation-asset-events'
import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from '@wiswork/pptx-engine/presentation-source-limits'
import {
  parsePresentationDeliveryReport,
  type PresentationDeliveryReport,
} from '@wiswork/pptx-engine/presentation-delivery-report'
import type { PresentationIssueActionInput } from '@wiswork/project-store/presentation-issue'
import type { PresentationProductionJob } from '@wiswork/project-store/presentation-job'
import type { PresentationPlanRevisionSnapshot } from '@wiswork/project-store'
import {
  parsePresentationPlanAcceptances,
  type PresentationPlanAcceptanceLedger,
  type PresentationPlanAcceptance,
} from '@wiswork/project-store/presentation-plan-acceptance'
import {
  canonicalPresentationValue,
  presentationPlanSnapshotInputs,
} from '@wiswork/project-store/presentation-canonical'
import { parsePresentationJobResponse } from './presentation-jobs.js'
import {
  parsePresentationProductionStatus,
  type PresentationProductionStatus,
} from './presentation-production.js'
import {
  parsePresentationPlan,
  presentationSourceAttachmentId,
  type PresentationPlan,
} from '@wiswork/pptx-engine/presentation-plan'
import type { AgentSkill } from '@wiswork/agent-core'
import type { PresentationGenerationOptions } from './presentation-generation.js'
import {
  presentationHostAssociations,
  type PresentationHostAssociations,
} from './presentation-host-associations.js'
import type { PresentationImportRecord } from './presentation-delivery.js'
import {
  parsePresentationImportSource,
  type PresentationImportSource,
} from '@wiswork/project-store/presentation-import-source'
import {
  parsePresentationSourceAuditHistory,
  parsePresentationSourceAuditRun,
  presentationSourceAuditHistory,
  type PresentationSourceAuditHistory,
  type PresentationSourceAuditResult,
} from '@wiswork/project-store/presentation-source-audit'

export interface PresentationProductionTask {
  requestId: string
  sequence: number
  planRevision: number
  status: PresentationProductionStatus['status']
  compiledCount: number
  total: number
  jobState?: PresentationProductionJob['state']
  lastEvent?: {
    type: PresentationProductionJob['events'][number]['type']
    createdAt: string
    pageId?: string
  }
}
export interface PresentationProjectStatus {
  researchSummary?: PresentationResearchSummary
  researchHistoryUnavailable?: boolean
  deliveryBundlesAvailable?: true
  assetHistory?: PresentationAssetLedger
  assetHistoryUnavailable?: boolean
  planAcceptance?: PresentationPlanAcceptanceLedger
  planAcceptanceUnavailable?: boolean
  planAcceptanceCurrent?: PresentationPlanAcceptance
  sourceAuditHistory?: PresentationSourceAuditHistory
  sourceAuditHistoryUnavailable?: boolean
  hostAssociations?: PresentationHostAssociations
  hostAssociationsUnavailable?: boolean
  sourcePreparation?: {
    sourceId: string
    attachmentId: string
    status:
      | 'ready'
      | 'excerpt_matched'
      | 'uploading'
      | 'failed'
      | 'missing'
      | 'unsupported'
      | 'excerpt_mismatch'
      | 'excerpt_missing'
      | 'source_mismatch'
      | 'locator_mismatch'
  }[]
  sourcePreparationUnavailable?: boolean
  reviewComments?: {
    revision: number
    openCount: number
    resolvedCount: number
    recent: {
      id: string
      targetKind: 'slide' | 'claim' | 'source'
      targetId: string
      authorLabel: string
      text: string
      state: 'open' | 'resolved'
      planRevision: number
      createdAt: string
    }[]
  }
  commentsUnavailable?: boolean
  productionTasks?: PresentationProductionTask[]
  productionJob?: PresentationProductionJob | null
  jobsUnavailable?: boolean
  production?: PresentationProductionStatus
  projectId: string
  title: string
  status: 'planned' | 'pending' | 'compiled'
  plan?: {
    revision: number
    value: PresentationPlan
    revisions?: {
      revision: number
      inputDigest: string
      createdAt: string
      snapshot?: PresentationPlanRevisionSnapshot
    }[]
  }
  requestPlanRevision?: number
  latestRequestId?: string
  latestCompiledRequestId?: string
  slideCount: number
  slides: { id: string; title: string }[]
  history: {
    requestId: string
    sequence: number
    status: 'pending' | 'compiled'
    slideCount: number
  }[]
  checks?: {
    structure: 'passed'
    geometry: 'passed' | 'warning'
    render: 'not_run'
    sources: 'not_verified'
    roundTrip: 'not_run'
  }
}
export interface PresentationProjectSnapshot {
  feedbackComparisonBaselineRequestId?: string
  feedbackComparison?: PresentationFeedbackComparison
  feedbackComparisonUnavailable?: true
  productionFeedback?: PresentationProductionFeedbackLedger | null
  productionFeedbackUnavailable?: true
  deliveryBundlesUnavailable?: true
  deliveryBundles?: PresentationDeliveryBundleReceipt[]
  bundleNotice?: string
  planNotice?: string
  sourceAudit?: {
    auditId?: string
    finishedAt?: string
    planRevision: number
    sources: PresentationSourceAuditResult[]
  }
  deliveryReport?: PresentationDeliveryReport
  deliveryNotice?: string
  phase:
    | 'idle'
    | 'loading'
    | 'restoring'
    | 'resuming'
    | 'producing'
    | 'auditing'
    | 'planning'
    | 'accepting'
    | 'readingResearch'
    | 'bundling'
  project?: PresentationProjectStatus
  error?: string
}
export type PresentationPlanEdit =
  | { kind: 'move'; pageId: string; direction: 'up' | 'down' }
  | { kind: 'delete'; pageId: string }
  | { kind: 'restore'; revision: number }
  | { kind: 'lock'; pageId: string; locked: boolean }
export interface PresentationProjectController {
  selectFeedbackComparisonBaseline?(requestId?: string): void
  readFeedbackComparison?(): Promise<void>
  readProductionFeedback?(): Promise<void>
  recordProductionFeedback?(pages: PresentationProductionFeedbackPage[]): Promise<void>
  readBoundResearch?(): Promise<void>
  currentBundleAvailable?(): boolean
  exportCurrentBundle?(includePdf?: boolean, includePageScreenshots?: boolean): Promise<void>
  restoreDeliveryBundle?(bundleId: string): Promise<void>
  readDeliveryBundles?(): Promise<void>
  deleteDeliveryBundle?(bundleId: string): Promise<void>
  acceptPlan?(expectedRevision: number): Promise<void>
  editPlan?(expectedRevision: number, action: PresentationPlanEdit): Promise<void>
  pdfAvailable?(): boolean
  exportProductionPdf?(): Promise<void>
  auditSources(): Promise<void>
  readDeliveryReport(): Promise<void>
  exportDeliveryReport(): Promise<void>
  recordIssueAction(action: PresentationIssueActionInput): Promise<void>
  snapshot(): PresentationProjectSnapshot
  subscribe(listener: () => void): () => void
  refresh(): Promise<void>
  restore(): Promise<void>
  resume(requestId: string): Promise<void>
  runProduction(requestId: string): Promise<void>
  selectProduction(requestId: string): Promise<void>
  startProductionJob(requestId: string): Promise<void>
  pauseProductionJob(requestId: string): Promise<void>
  resumeProductionJob(requestId: string): Promise<void>
  cancelProductionJob(requestId: string): Promise<void>
  downloadProductionPage(pageId: string): Promise<void>
  prepareProduction(): Promise<void>
  cancel(): void
  clear(): void
  prepareReconnect(): void
}
const validId = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const pageCount = (value: unknown): value is number =>
  Number.isSafeInteger(value) && Number(value) >= 1 && Number(value) <= 100
function validRevisionSnapshot(value: unknown): value is PresentationPlanRevisionSnapshot {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const snapshot = value as PresentationPlanRevisionSnapshot
  return (
    Object.keys(snapshot).sort().join(',') ===
      'claimCount,claimsDigest,slideCount,slidesDigest,sourceCount,sourcesDigest,styleDigest' &&
    [snapshot.sourceCount, snapshot.claimCount, snapshot.slideCount].every(
      (count) => Number.isSafeInteger(count) && count >= 0,
    ) &&
    snapshot.sourceCount <= 256 &&
    snapshot.claimCount <= 256 &&
    snapshot.slideCount <= 32 &&
    [
      snapshot.sourcesDigest,
      snapshot.claimsDigest,
      snapshot.slidesDigest,
      snapshot.styleDigest,
    ].every((hash) => typeof hash === 'string' && /^[a-f0-9]{64}$/.test(hash))
  )
}
async function presentationDigest(input: string): Promise<string> {
  const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(input))
  return Array.from(new Uint8Array(bytes), (byte) => byte.toString(16).padStart(2, '0')).join('')
}
async function parseStatus(value: unknown, projectId: string): Promise<PresentationProjectStatus> {
  const p = value as PresentationProjectStatus | undefined
  let researchSummary: PresentationResearchSummary | undefined
  let researchHistoryUnavailable =
    Object.hasOwn(p ?? {}, 'researchHistoryUnavailable') && p?.researchHistoryUnavailable !== false
  if (!researchHistoryUnavailable && Object.hasOwn(p ?? {}, 'researchSummary')) {
    try {
      researchSummary = parsePresentationResearchSummary(p?.researchSummary)
      if (researchSummary.projectId !== projectId) throw new Error('presentation_response_invalid')
    } catch {
      researchSummary = undefined
      researchHistoryUnavailable = true
    }
  }
  if (p?.deliveryBundlesAvailable !== undefined && p.deliveryBundlesAvailable !== true)
    throw new Error('presentation_response_invalid')
  let planAcceptance: PresentationPlanAcceptanceLedger | undefined
  let planAcceptanceUnavailable = p?.planAcceptanceUnavailable === true
  if (p?.planAcceptance !== undefined) {
    try {
      planAcceptance = parsePresentationPlanAcceptances(p.planAcceptance)
      if (
        planAcceptance.projectId !== projectId ||
        (p.plan && planAcceptance.records.some((record) => record.planRevision > p.plan!.revision))
      )
        throw new Error('presentation_response_invalid')
    } catch {
      planAcceptance = undefined
      planAcceptanceUnavailable = true
    }
  }
  let sourceAuditHistory: PresentationSourceAuditHistory | undefined
  let sourceAuditHistoryUnavailable = p?.sourceAuditHistoryUnavailable === true
  if (p?.sourceAuditHistory !== undefined) {
    try {
      sourceAuditHistory = parsePresentationSourceAuditHistory(p.sourceAuditHistory)
      if (
        sourceAuditHistory.projectId !== projectId ||
        (p.plan && sourceAuditHistory.runs.some((run) => run.planRevision > p.plan!.revision))
      )
        throw new Error('presentation_response_invalid')
    } catch {
      sourceAuditHistory = undefined
      sourceAuditHistoryUnavailable = true
    }
  }
  let plan: PresentationProjectStatus['plan']
  if (p?.plan !== undefined) {
    if (!p.plan || !Number.isSafeInteger(p.plan.revision) || p.plan.revision < 1)
      throw new Error('presentation_response_invalid')
    let revisions: NonNullable<PresentationProjectStatus['plan']>['revisions']
    if (p.plan.revisions !== undefined) {
      if (
        !Array.isArray(p.plan.revisions) ||
        p.plan.revisions.length < 1 ||
        p.plan.revisions.length > 32
      )
        throw new Error('presentation_response_invalid')
      revisions = []
      for (const [index, event] of p.plan.revisions.entries()) {
        if (
          !event ||
          typeof event !== 'object' ||
          Array.isArray(event) ||
          Object.keys(event).sort().join(',') !==
            (event.snapshot === undefined
              ? 'createdAt,inputDigest,revision'
              : 'createdAt,inputDigest,revision,snapshot') ||
          !Number.isSafeInteger(event.revision) ||
          event.revision < 1 ||
          (index > 0 && event.revision !== revisions[index - 1]!.revision + 1) ||
          typeof event.inputDigest !== 'string' ||
          !/^[a-f0-9]{64}$/.test(event.inputDigest) ||
          (event.snapshot !== undefined && !validRevisionSnapshot(event.snapshot)) ||
          typeof event.createdAt !== 'string' ||
          !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(event.createdAt) ||
          !Number.isFinite(Date.parse(event.createdAt)) ||
          new Date(event.createdAt).toISOString() !== event.createdAt ||
          (index > 0 && event.createdAt < revisions[index - 1]!.createdAt)
        )
          throw new Error('presentation_response_invalid')
        revisions.push({
          revision: event.revision,
          inputDigest: event.inputDigest,
          createdAt: event.createdAt,
          ...(event.snapshot ? { snapshot: event.snapshot } : {}),
        })
      }
      if (revisions.at(-1)!.revision !== p.plan.revision)
        throw new Error('presentation_response_invalid')
    }
    plan = {
      revision: p.plan.revision,
      value: parsePresentationPlan(p.plan.value),
      ...(revisions ? { revisions } : {}),
    }
    if (plan.value.projectId !== projectId) throw new Error('presentation_response_invalid')
    const latestSnapshot = revisions?.at(-1)?.snapshot
    if (latestSnapshot) {
      const inputs = presentationPlanSnapshotInputs(p.plan.value)
      for (const key of Object.keys(inputs) as (keyof typeof inputs)[]) {
        const hash = await presentationDigest(inputs[key])
        if (hash !== latestSnapshot[key]) throw new Error('presentation_response_invalid')
      }
      if (
        (await presentationDigest(canonicalPresentationValue(p.plan.value))) !==
        revisions!.at(-1)!.inputDigest
      )
        throw new Error('presentation_response_invalid')
    }
    if (
      latestSnapshot &&
      (latestSnapshot.sourceCount !== plan.value.sources.length ||
        latestSnapshot.claimCount !== plan.value.claims.length ||
        latestSnapshot.slideCount !== plan.value.slides.length)
    )
      throw new Error('presentation_response_invalid')
  }
  const production =
    p?.production === undefined ? undefined : parsePresentationProductionStatus(p.production)
  if (production && production.projectId !== projectId)
    throw new Error('presentation_response_invalid')
  let productionTasks: PresentationProductionTask[] | undefined
  if (p?.productionTasks !== undefined) {
    const tasks = p.productionTasks
    if (
      !Array.isArray(tasks) ||
      tasks.length > 32 ||
      !tasks.every(
        (task, index) =>
          task &&
          typeof task === 'object' &&
          !Array.isArray(task) &&
          Object.keys(task).every((key) =>
            [
              'requestId',
              'sequence',
              'planRevision',
              'status',
              'compiledCount',
              'total',
              'jobState',
              'lastEvent',
            ].includes(key),
          ) &&
          validId(task.requestId) &&
          Number.isSafeInteger(task.sequence) &&
          task.sequence > 0 &&
          Number.isSafeInteger(task.planRevision) &&
          task.planRevision > 0 &&
          Number.isSafeInteger(task.total) &&
          task.total > 0 &&
          task.total <= 32 &&
          Number.isSafeInteger(task.compiledCount) &&
          task.compiledCount >= 0 &&
          task.compiledCount <= task.total &&
          ['pending', 'building', 'partial', 'compiled'].includes(task.status) &&
          (task.status === 'compiled') === (task.compiledCount === task.total) &&
          (task.status !== 'pending' || task.compiledCount === 0) &&
          (task.jobState === undefined ||
            [
              'running',
              'pausing',
              'paused',
              'cancelling',
              'cancelled',
              'interrupted',
              'completed',
              'failed',
            ].includes(task.jobState)) &&
          (task.jobState !== 'completed' || task.status === 'compiled') &&
          (task.lastEvent === undefined ||
            (!task.jobState
              ? false
              : task.lastEvent !== null &&
                typeof task.lastEvent === 'object' &&
                !Array.isArray(task.lastEvent) &&
                Object.keys(task.lastEvent).sort().join(',') ===
                  ('pageId' in task.lastEvent ? 'createdAt,pageId,type' : 'createdAt,type') &&
                [
                  'run.started',
                  'run.pause_requested',
                  'run.paused',
                  'run.cancel_requested',
                  'run.cancelled',
                  'run.interrupted',
                  'run.completed',
                  'run.failed',
                  'page.started',
                  'page.compiled',
                  'page.failed',
                ].includes(task.lastEvent.type) &&
                typeof task.lastEvent.createdAt === 'string' &&
                /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(task.lastEvent.createdAt) &&
                Number.isFinite(Date.parse(task.lastEvent.createdAt)) &&
                new Date(task.lastEvent.createdAt).toISOString() === task.lastEvent.createdAt &&
                (task.lastEvent.type.startsWith('page.')
                  ? validId(task.lastEvent.pageId)
                  : task.lastEvent.pageId === undefined))) &&
          (index === 0 || task.sequence < tasks[index - 1]!.sequence),
      ) ||
      new Set(tasks.map((task) => task.requestId)).size !== tasks.length ||
      (production &&
        !tasks.some(
          (task) =>
            task.requestId === production.requestId &&
            task.planRevision === production.planRevision &&
            task.total === production.total,
        ))
    )
      throw new Error('presentation_response_invalid')
    productionTasks = structuredClone(tasks)
  }
  const planned = p?.status === 'planned'
  const comments = p?.reviewComments
  if (p?.commentsUnavailable !== undefined && p.commentsUnavailable !== true)
    throw new Error('presentation_response_invalid')
  if (
    comments !== undefined &&
    (!comments || typeof comments !== 'object' || Array.isArray(comments))
  )
    throw new Error('presentation_response_invalid')
  if (comments !== undefined) {
    if (
      p?.commentsUnavailable ||
      !Number.isSafeInteger(comments.revision) ||
      comments.revision < 0 ||
      comments.revision > 256 ||
      !Number.isSafeInteger(comments.openCount) ||
      comments.openCount < 0 ||
      comments.openCount > 128 ||
      !Number.isSafeInteger(comments.resolvedCount) ||
      comments.resolvedCount < 0 ||
      comments.resolvedCount > 128 ||
      comments.openCount + comments.resolvedCount > 128 ||
      !Array.isArray(comments.recent) ||
      comments.recent.length > 8 ||
      comments.recent.length > comments.openCount + comments.resolvedCount ||
      comments.recent.some(
        (comment) =>
          !comment ||
          typeof comment !== 'object' ||
          Object.keys(comment).sort().join(',') !==
            'authorLabel,createdAt,id,planRevision,state,targetId,targetKind,text' ||
          !validId(comment.id) ||
          !validId(comment.targetId) ||
          !['slide', 'claim', 'source'].includes(comment.targetKind) ||
          typeof comment.authorLabel !== 'string' ||
          !comment.authorLabel ||
          comment.authorLabel.length > 80 ||
          typeof comment.text !== 'string' ||
          !comment.text ||
          comment.text.length > 400 ||
          !['open', 'resolved'].includes(comment.state) ||
          !Number.isSafeInteger(comment.planRevision) ||
          comment.planRevision < 1 ||
          (plan && comment.planRevision > plan.revision) ||
          typeof comment.createdAt !== 'string' ||
          !Number.isFinite(Date.parse(comment.createdAt)) ||
          new Date(comment.createdAt).toISOString() !== comment.createdAt,
      ) ||
      new Set(comments.recent.map((comment) => comment.id)).size !== comments.recent.length
    )
      throw new Error('presentation_response_invalid')
  }
  if (p?.sourcePreparationUnavailable !== undefined && p.sourcePreparationUnavailable !== true)
    throw new Error('presentation_response_invalid')
  if (p?.sourcePreparation !== undefined) {
    const expected = plan?.value.sources.flatMap((source) => {
      const attachmentId = presentationSourceAttachmentId(source)
      return attachmentId ? [{ sourceId: source.id, attachmentId }] : []
    })
    if (
      !expected ||
      p.sourcePreparationUnavailable ||
      !Array.isArray(p.sourcePreparation) ||
      p.sourcePreparation.length !== expected.length ||
      p.sourcePreparation.some(
        (item, index) =>
          !item ||
          typeof item !== 'object' ||
          Object.keys(item).sort().join(',') !== 'attachmentId,sourceId,status' ||
          item.sourceId !== expected[index]!.sourceId ||
          item.attachmentId !== expected[index]!.attachmentId ||
          ![
            'ready',
            'excerpt_matched',
            'uploading',
            'failed',
            'missing',
            'unsupported',
            'excerpt_mismatch',
            'excerpt_missing',
            'source_mismatch',
            'locator_mismatch',
          ].includes(item.status),
      )
    )
      throw new Error('presentation_response_invalid')
  }
  if (p?.sourcePreparationUnavailable && !plan) throw new Error('presentation_response_invalid')
  if (
    !p ||
    p.projectId !== projectId ||
    typeof p.title !== 'string' ||
    p.title.length > 500 ||
    !['planned', 'pending', 'compiled'].includes(p.status) ||
    (!planned && !validId(p.latestRequestId)) ||
    (planned &&
      (!plan ||
        p.latestRequestId !== undefined ||
        p.latestCompiledRequestId !== undefined ||
        p.checks !== undefined)) ||
    (p.requestPlanRevision !== undefined &&
      (!Number.isSafeInteger(p.requestPlanRevision) || p.requestPlanRevision < 1)) ||
    (p.latestCompiledRequestId !== undefined && !validId(p.latestCompiledRequestId)) ||
    !pageCount(p.slideCount) ||
    !Array.isArray(p.slides) ||
    p.slides.length !== p.slideCount ||
    !p.slides.every(
      (s) => s && validId(s.id) && typeof s.title === 'string' && s.title.length <= 500,
    ) ||
    new Set(p.slides.map((s) => s.id)).size !== p.slideCount ||
    !Array.isArray(p.history) ||
    (!planned && p.history.length < 1) ||
    (planned && p.history.length !== 0) ||
    p.history.length > 20 ||
    !p.history.every(
      (r, i) =>
        r &&
        validId(r.requestId) &&
        Number.isSafeInteger(r.sequence) &&
        r.sequence > 0 &&
        ['pending', 'compiled'].includes(r.status) &&
        pageCount(r.slideCount) &&
        (i === 0 || r.sequence < p.history[i - 1]!.sequence),
    ) ||
    new Set(p.history.map((r) => r.requestId)).size !== p.history.length ||
    (!planned &&
      (p.history[0]!.requestId !== p.latestRequestId ||
        p.history[0]!.status !== p.status ||
        p.history[0]!.slideCount !== p.slideCount)) ||
    (p.status === 'pending' && p.checks !== undefined) ||
    (p.status === 'compiled' && (!p.checks || p.latestCompiledRequestId !== p.latestRequestId)) ||
    (p.checks &&
      (p.checks.structure !== 'passed' ||
        !['passed', 'warning'].includes(p.checks.geometry) ||
        p.checks.render !== 'not_run' ||
        p.checks.sources !== 'not_verified' ||
        p.checks.roundTrip !== 'not_run'))
  )
    throw new Error('presentation_response_invalid')
  let assetHistory: PresentationAssetLedger | undefined
  let assetHistoryUnavailable = p?.assetHistoryUnavailable === true
  if (p?.assetHistory !== undefined) {
    try {
      assetHistory = parsePresentationAssetLedger(p.assetHistory)
      if (
        assetHistory.projectId !== projectId ||
        !p.production ||
        assetHistory.requestId !== p.production.requestId ||
        assetHistory.events.some(
          (event) =>
            !p.production!.pages.some(
              (page) => page.id === event.pageId && page.attempt >= event.attempt,
            ),
        )
      )
        throw new Error('invalid_state')
    } catch {
      assetHistory = undefined
      assetHistoryUnavailable = true
    }
  }
  // Copy only the bounded public projection; never retain arbitrary server fields or binary data.
  return {
    ...(researchSummary ? { researchSummary } : {}),
    ...(researchHistoryUnavailable ? { researchHistoryUnavailable: true } : {}),
    ...(p.deliveryBundlesAvailable ? { deliveryBundlesAvailable: true } : {}),
    ...(assetHistory ? { assetHistory } : {}),
    ...(assetHistoryUnavailable ? { assetHistoryUnavailable: true } : {}),
    ...(planAcceptance ? { planAcceptance } : {}),
    ...(planAcceptanceUnavailable ? { planAcceptanceUnavailable: true } : {}),
    ...(sourceAuditHistory ? { sourceAuditHistory } : {}),
    ...(sourceAuditHistoryUnavailable ? { sourceAuditHistoryUnavailable: true } : {}),
    ...(p.sourcePreparation ? { sourcePreparation: structuredClone(p.sourcePreparation) } : {}),
    ...(p.sourcePreparationUnavailable ? { sourcePreparationUnavailable: true } : {}),
    ...(comments ? { reviewComments: structuredClone(comments) } : {}),
    ...(p.commentsUnavailable ? { commentsUnavailable: true } : {}),
    ...(production ? { production } : {}),
    ...(productionTasks ? { productionTasks } : {}),
    projectId: p.projectId,
    title: p.title,
    status: p.status,
    ...(p.latestRequestId ? { latestRequestId: p.latestRequestId } : {}),
    ...(plan ? { plan } : {}),
    ...(p.requestPlanRevision ? { requestPlanRevision: p.requestPlanRevision } : {}),
    ...(p.latestCompiledRequestId ? { latestCompiledRequestId: p.latestCompiledRequestId } : {}),
    slideCount: p.slideCount,
    slides: p.slides.map(({ id, title }) => ({ id, title })),
    history: p.history.map(({ requestId, sequence, status, slideCount }) => ({
      requestId,
      sequence,
      status,
      slideCount,
    })),
    ...(p.checks
      ? {
          checks: {
            structure: p.checks.structure,
            geometry: p.checks.geometry,
            render: p.checks.render,
            sources: p.checks.sources,
            roundTrip: p.checks.roundTrip,
          },
        }
      : {}),
  }
}
function message(error: unknown): string {
  const code = error instanceof Error ? error.message : ''
  if (code === 'presentation_research_binding_invalid')
    return '计划绑定的研究身份或内容不一致；请核对指定研究，不自动替换为最近记录。'
  if (code === 'presentation_research_unavailable')
    return '计划绑定的原研究记录暂不可读取；原计划与素材保留，请刷新后读取指定记录。'
  if (code === 'presentation_quota_exceeded')
    return '本机交付包容量已满，请明确删除不再需要的本机包后重试。'
  if (code === 'presentation_delivery_bundle_limit' || code === 'vfs_limit')
    return '交付包超过本机或会话容量限制，请减小文稿体积后重试。'
  if (code === 'office_document_export_unavailable')
    return '当前 Office 宿主不支持原生文稿导出，请使用支持的 PowerPoint 环境。'
  if (code === 'office_document_export_failed')
    return 'Office 未能完成文稿导出，原始文稿保留，请检查宿主状态后重试。'
  if (code === 'office_document_export_cancelled') return '宿主文稿导出已取消，原始文稿保留。'
  if (code === 'office_document_export_timeout')
    return '宿主文稿导出等待超时，请检查 Office 状态后重试。'
  if (
    code === 'office_document_export_invalid' ||
    code === 'presentation_response_invalid' ||
    code === 'presentation_invalid_state'
  )
    return '交付包内容或回执校验未通过，请刷新已有包后核对。'
  if (code === 'presentation_delivery_bundle_history_invalid')
    return '历史 QA 或保存点记录不一致，请刷新相关记录后重试。'
  if (code === 'presentation_delivery_bundle_incomplete')
    return '本机包尚未完整保存，可明确删除未完成包释放容量。'

  if (
    ['presentation_upgrade_required', 'presentation_invalid_request', 'invalid_request'].includes(
      code,
    )
  )
    return '当前 PC 尚不支持此项目操作，请升级 WisWork PC 后重试。'
  if (code === 'presentation_revision_conflict') return '记录已有更新，请重新读取后继续。'
  if (code === 'presentation_acceptance_capacity')
    return '接受决定记录已达到容量上限；现有历史与页面保留，普通制作仍可继续。'
  if (code === 'presentation_page_locked')
    return '此操作影响锁定页，请先在工作台明确解除该页锁定后继续。'
  if (code === 'presentation_plan_revision_unavailable')
    return '该历史版本的完整计划不可用，请选择其它版本或保留当前计划。'
  if (code === 'presentation_invalid_plan')
    return 'PC 未接受此计划，请核对品牌版本和计划约束后继续。'
  if (code === 'presentation_plan_invalid:page_dependency')
    return '页面依赖不允许这次排序或删除，请先调整依赖再继续。'
  if (code === 'presentation_plan_invalid:domain_section')
    return '此页包含行业方案要求的章节，请保留该章节或先调整故事结构。'
  if (code === 'presentation_plan_last_page') return '计划至少需要一页，不能删除最后一页。'
  if (code === 'presentation_plan_page_missing') return '计划页已有变化，请刷新计划后继续。'
  if (code.startsWith('presentation_plan_invalid:'))
    return '这次调整不符合计划约束，请核对品牌、章节和页面依赖后继续。'
  if (code === 'presentation_issue_changed') return '问题证据已有变化，请重新读取报告后处理。'
  if (code === 'presentation_plan_mismatch')
    return '编译内容与保存的计划不一致，请先更新计划或修正内容。'
  if (code === 'presentation_not_found')
    return 'PC 上尚未找到保存的编译请求。请重新发起制作；已有文稿不会改动。'
  if (['presentation_document_changed', 'presentation_document_mismatch'].includes(code))
    return '当前文档与项目不匹配，请返回原文档后刷新。'
  if (code === 'presentation_unavailable') return 'PC 连接不可用，请重新连接后刷新项目。'
  if (code === 'presentation_selection_save_failed')
    return '无法保存所选页任务，原有项目状态已保留。请检查文档设置后重试。'
  if (code === 'cancelled' || code === 'presentation_aborted')
    return '已停止等待；PC 可能已保存结果，可刷新查看。'
  return '暂时无法恢复项目，已有成果已保留。请检查 PC 连接后重试。'
}

export function createPresentationProjectController(
  options: Pick<
    PresentationGenerationOptions,
    | 'request'
    | 'available'
    | 'pdfAvailable'
    | 'productionPdfAvailable'
    | 'documentId'
    | 'lastProject'
    | 'selectedProduction'
    | 'rememberSelectedProduction'
  > &
    Pick<AgentSkill, 'executeTool'> & {
      readResearchRecord?(projectId: string, ledgerId: string, signal?: AbortSignal): Promise<void>
      nativeDocumentExportAvailable?(): boolean
      listReceipts?(): { key: string; record: PresentationImportRecord }[]
      hostSlideIds?(signal?: AbortSignal): Promise<string[]>
    },
): PresentationProjectController {
  let state: PresentationProjectSnapshot = { phase: 'idle' }
  let comparisonPending = false
  let verifiedComparison: PresentationFeedbackComparison | undefined
  let verifiedFeedback: PresentationProductionFeedbackLedger | undefined
  const listeners = new Set<() => void>()
  let epoch = 0
  let projectDocument: string | undefined
  let selection: { documentId: string; projectId: string; requestId: string } | undefined
  let ignoreStoredSelection = false
  let recoverInterruptedJob = false
  let active: AbortController | undefined
  let poll: ReturnType<typeof setTimeout> | undefined
  const stopPolling = () => {
    clearTimeout(poll)
    poll = undefined
  }
  const publish = (next: PresentationProjectSnapshot) => {
    state = next
    for (const listener of listeners) listener()
  }
  const readBundleList = async (
    project: PresentationProjectStatus,
    documentId: string,
    signal: AbortSignal,
    check: () => void,
  ) => {
    const requestId = project.production?.requestId
    if (!project.deliveryBundlesAvailable || !requestId) return undefined
    const response = await options.request(
      { operation: 'delivery_bundle_list', documentId, projectId: project.projectId, requestId },
      signal,
    )
    check()
    if ((await options.documentId()) !== documentId)
      throw new Error('presentation_document_changed')
    check()
    if (!response.ok) throw new Error('presentation_service_unavailable')
    const text = await response.text()
    check()
    if (new TextEncoder().encode(text).byteLength > 256 * 1024)
      throw new Error('presentation_response_invalid')
    const value = JSON.parse(text)
    if (
      !value ||
      Object.keys(value).join(',') !== 'bundles' ||
      !Array.isArray(value.bundles) ||
      value.bundles.length > 32
    )
      throw new Error('presentation_response_invalid')
    const bundles: PresentationDeliveryBundleReceipt[] = value.bundles.map(
      parsePresentationDeliveryBundleReceipt,
    )
    if (
      bundles.some(
        (receipt) =>
          receipt.documentId !== documentId ||
          receipt.projectId !== project.projectId ||
          receipt.requestId !== requestId,
      ) ||
      new Set(bundles.map((item) => item.bundleId)).size !== bundles.length
    )
      throw new Error('presentation_response_invalid')
    if ((await options.documentId()) !== documentId)
      throw new Error('presentation_document_changed')
    check()
    return bundles
  }
  const restoreSourceAudit = async (
    project: PresentationProjectStatus,
    documentId: string,
    signal: AbortSignal,
    check: () => void,
  ): Promise<PresentationProjectSnapshot['sourceAudit']> => {
    const history = project.sourceAuditHistory
    if (!history) return
    if (history.documentId !== documentId) throw new Error('presentation_response_invalid')
    if (project.sourceAuditHistoryUnavailable || !project.plan) return
    const planDigest = await presentationDigest(canonicalPresentationValue(project.plan.value))
    check()
    const expected = [...history.runs]
      .reverse()
      .find(
        (run) =>
          run.state === 'completed' &&
          run.planRevision === project.plan!.revision &&
          run.planDigest === planDigest,
      )
    if (!expected) return
    const response = await options.request(
      {
        operation: 'read_source_audit',
        documentId,
        projectId: project.projectId,
        auditId: expected.id,
      },
      signal,
    )
    check()
    if ((await options.documentId()) !== documentId)
      throw new Error('presentation_document_changed')
    check()
    if (!response.ok) throw new Error('presentation_service_unavailable')
    const text = await response.text()
    check()
    if (new TextEncoder().encode(text).byteLength > 256 * 1024)
      throw new Error('presentation_response_invalid')
    const value = JSON.parse(text)
    if (
      !value ||
      Object.keys(value).sort().join(',') !== 'audit,documentId,projectId' ||
      value.documentId !== documentId ||
      value.projectId !== project.projectId
    )
      throw new Error('presentation_response_invalid')
    const run = parsePresentationSourceAuditRun(value.audit)
    const observed = presentationSourceAuditHistory({
      version: 1,
      projectId: project.projectId,
      documentId,
      revision: history.revision,
      runs: [run],
    }).runs[0]
    const refs = project.plan.value.sources.flatMap((source) => {
      const attachmentId = presentationSourceAttachmentId(source)
      return attachmentId ? [{ sourceId: source.id, attachmentId }] : []
    })
    if (
      canonicalPresentationValue(observed) !== canonicalPresentationValue(expected) ||
      canonicalPresentationValue(refs) !== canonicalPresentationValue(run.sourceRefs)
    )
      throw new Error('presentation_response_invalid')
    const current = await options.request(
      { operation: 'get_plan', documentId, projectId: project.projectId },
      signal,
    )
    check()
    if (!current.ok) throw new Error('presentation_service_unavailable')
    const currentText = await current.text()
    check()
    if (new TextEncoder().encode(currentText).byteLength > 512 * 1024)
      throw new Error('presentation_response_invalid')
    const saved = JSON.parse(currentText)
    if (
      saved.projectId !== project.projectId ||
      saved.revision !== project.plan.revision ||
      (await presentationDigest(canonicalPresentationValue(parsePresentationPlan(saved.plan)))) !==
        planDigest
    )
      throw new Error('presentation_response_invalid')
    if ((await options.documentId()) !== documentId)
      throw new Error('presentation_document_changed')
    check()
    return {
      auditId: run.id,
      finishedAt: run.finishedAt,
      planRevision: run.planRevision,
      sources: run.sources!,
    }
  }
  const stop = (error?: string, resetSelection = false) => {
    stopPolling()
    comparisonPending = false
    verifiedComparison = undefined
    verifiedFeedback = undefined
    projectDocument = undefined
    selection = undefined
    ignoreStoredSelection = resetSelection
    recoverInterruptedJob = false
    epoch += 1
    active?.abort()
    active = undefined
    publish({ phase: 'idle', ...(error ? { error } : {}) })
  }
  const run = async (
    phase: 'loading' | 'restoring' | 'resuming' | 'producing',
    requestId?: string,
    tool?: string,
    pageId?: string,
  ) => {
    if (active) return
    const projectId = phase === 'loading' ? options.lastProject() : state.project?.projectId
    if (!projectId) {
      publish({ phase: 'idle' })
      return
    }
    if (
      !validId(projectId) ||
      (phase === 'producing' &&
        !tool &&
        (!validId(requestId) ||
          requestId !== state.project?.production?.requestId ||
          state.project.production.status === 'compiled')) ||
      (phase === 'resuming' &&
        (state.project?.status !== 'pending' || requestId !== state.project.latestRequestId)) ||
      (phase === 'restoring' && !state.project?.latestCompiledRequestId)
    )
      return
    stopPolling()
    const controller = new AbortController()
    active = controller
    const captured = ++epoch
    const previous = state.project?.projectId === projectId ? state.project : undefined
    let boundDocument: string | undefined
    publish({ phase, ...(previous ? { project: previous } : {}) })
    const check = () => {
      if (controller.signal.aborted || captured !== epoch) throw new Error('cancelled')
      if (!options.available()) throw new Error('presentation_unavailable')
    }
    try {
      check()
      const documentId = await options.documentId()
      boundDocument = documentId
      check()
      if (selection && (selection.documentId !== documentId || selection.projectId !== projectId))
        selection = undefined
      if (phase === 'loading' && requestId && projectDocument !== documentId)
        throw new Error('presentation_document_changed')
      if (phase !== 'loading') {
        const result = await options.executeTool(
          {
            id: `presentation-control-${captured}`,
            name:
              tool ??
              (phase === 'restoring'
                ? 'restore_presentation_project'
                : phase === 'producing'
                  ? 'run_presentation_production'
                  : 'resume_presentation_project'),
            input: {
              project_id: projectId,
              ...(requestId ? { request_id: requestId } : {}),
              ...(pageId ? { page_id: pageId } : {}),
            },
          },
          controller.signal,
        )
        check()
        if (result.isError) throw new Error(result.output)
      }
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
      const response = await options.request(
        { operation: 'status', documentId, projectId },
        controller.signal,
      )
      check()
      if (!response.ok) throw new Error('presentation_service_unavailable')
      const text = await response.text()
      check()
      if (text.length > 384 * 1024) throw new Error('presentation_response_invalid')
      const value = JSON.parse(text)
      if (value?.error) throw new Error(`presentation_${value.error}`)
      const project = await parseStatus(value, projectId)
      check()
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
      if (project.researchSummary && project.researchSummary.documentId !== documentId) {
        delete project.researchSummary
        project.researchHistoryUnavailable = true
      }
      let selectedRequest =
        phase === 'loading' && requestId
          ? requestId
          : (selection?.requestId ??
            (ignoreStoredSelection
              ? undefined
              : options.selectedProduction?.(projectId, documentId)))
      if (
        selectedRequest &&
        !project.productionTasks?.some((task) => task.requestId === selectedRequest)
      )
        selectedRequest = undefined
      const productionRequest = selectedRequest ?? project.production?.requestId
      if (productionRequest) {
        const response = await options.request(
          {
            operation: 'production_job_status',
            documentId,
            projectId,
            requestId: productionRequest,
          },
          controller.signal,
        )
        check()
        if (!response.ok) throw new Error('presentation_service_unavailable')
        const text = await response.text()
        check()
        if (new TextEncoder().encode(text).byteLength > 256 * 1024)
          throw new Error('presentation_response_invalid')
        const value = JSON.parse(text)
        if (['invalid_request', 'upgrade_required'].includes(value?.error))
          project.jobsUnavailable = true
        else {
          if (value?.error) throw new Error(`presentation_${value.error}`)
          const result = parsePresentationJobResponse(
            value,
            documentId,
            projectId,
            productionRequest,
          )
          project.production = result.production
          project.productionJob = result.job
          if (recoverInterruptedJob && result.job?.state === 'interrupted') {
            if ((await options.documentId()) !== documentId)
              throw new Error('presentation_document_changed')
            check()
            const resumed = await options.executeTool(
              {
                id: `presentation-reconnect-${captured}`,
                name: 'resume_presentation_production_job',
                input: { project_id: projectId, request_id: productionRequest },
              },
              controller.signal,
            )
            check()
            if (resumed.isError) throw new Error(resumed.output)
            const refreshed = await options.request(
              {
                operation: 'production_job_status',
                documentId,
                projectId,
                requestId: productionRequest,
              },
              controller.signal,
            )
            check()
            if (!refreshed.ok) throw new Error('presentation_service_unavailable')
            const refreshedText = await refreshed.text()
            check()
            if (new TextEncoder().encode(refreshedText).byteLength > 256 * 1024)
              throw new Error('presentation_response_invalid')
            const refreshedValue = JSON.parse(refreshedText)
            if (refreshedValue?.error) throw new Error(`presentation_${refreshedValue.error}`)
            const resumedStatus = parsePresentationJobResponse(
              refreshedValue,
              documentId,
              projectId,
              productionRequest,
            )
            project.production = resumedStatus.production
            project.productionJob = resumedStatus.job
          }
        }
      }
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
      recoverInterruptedJob = false
      if (phase === 'loading' && requestId && selectedRequest && !project.jobsUnavailable) {
        try {
          await options.rememberSelectedProduction?.(projectId, documentId, selectedRequest)
        } catch (error) {
          if (error instanceof Error && error.message === 'presentation_document_changed')
            throw error
          throw new Error('presentation_selection_save_failed', { cause: error })
        }
        check()
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        ignoreStoredSelection = false
      }
      selection =
        selectedRequest && !project.jobsUnavailable
          ? { documentId, projectId, requestId: selectedRequest }
          : undefined
      if (project.plan && options.listReceipts && options.hostSlideIds) {
        try {
          const receipts = options.listReceipts()
          const capturedReceipts = JSON.stringify(receipts)
          check()
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          check()
          if (JSON.stringify(options.listReceipts()) !== capturedReceipts)
            throw new Error('presentation_host_association_invalid')
          const sources: PresentationImportSource[] = []
          const references = receipts.flatMap(({ key, record }) => {
            const match = /^(production\/)?([A-Za-z0-9_-]{1,128})\/([A-Za-z0-9_-]{1,128})$/.exec(
              key,
            )
            return match && record.documentId === documentId && match[2] === projectId
              ? [
                  {
                    source: match[1] ? ('production' as const) : ('compiled' as const),
                    requestId: match[3]!,
                  },
                ]
              : []
          })
          for (let start = 0; start < references.length; start += 4) {
            if ((await options.documentId()) !== documentId)
              throw new Error('presentation_document_changed')
            check()
            const batch = await Promise.all(
              references.slice(start, start + 4).map(async (reference) => {
                const response = await options.request(
                  { operation: 'read_import_source', documentId, projectId, ...reference },
                  controller.signal,
                )
                check()
                if (!response.ok) return undefined
                const text = await response.text()
                check()
                if (new TextEncoder().encode(text).byteLength > 64 * 1024)
                  throw new Error('presentation_host_association_invalid')
                const value = JSON.parse(text)
                if (
                  ['not_found', 'page_not_ready', 'invalid_request', 'upgrade_required'].includes(
                    value?.error,
                  )
                )
                  return undefined
                const source = parsePresentationImportSource(value)
                if (
                  source.documentId !== documentId ||
                  source.projectId !== projectId ||
                  source.requestId !== reference.requestId ||
                  source.source !== reference.source
                )
                  throw new Error('presentation_host_association_invalid')
                return source
              }),
            )
            check()
            for (const source of batch) if (source) sources.push(source)
          }
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          check()
          const hostIds = await options.hostSlideIds(controller.signal)
          check()
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          check()
          if (JSON.stringify(options.listReceipts()) !== capturedReceipts)
            throw new Error('presentation_host_association_invalid')
          project.hostAssociations = presentationHostAssociations(
            project.plan.value,
            project.plan.revision,
            documentId,
            receipts,
            project.productionTasks ?? [],
            hostIds,
            sources,
          )
        } catch (error) {
          if (
            controller.signal.aborted ||
            captured !== epoch ||
            (error instanceof Error && error.message === 'presentation_document_changed')
          )
            throw error
          project.hostAssociationsUnavailable = true
        }
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
      }
      if (
        project.assetHistory &&
        (project.assetHistory.documentId !== documentId ||
          project.assetHistory.requestId !== project.production?.requestId)
      ) {
        delete project.assetHistory
        project.assetHistoryUnavailable = true
      }
      projectDocument = documentId
      if (project.planAcceptance && !project.planAcceptanceUnavailable) {
        try {
          if (project.planAcceptance.documentId !== documentId)
            throw new Error('presentation_response_invalid')
          if (project.plan) {
            const [planDigest, styleDigest] = await Promise.all([
              presentationDigest(canonicalPresentationValue(project.plan.value)),
              presentationDigest(presentationPlanSnapshotInputs(project.plan.value).styleDigest),
            ])
            check()
            project.planAcceptanceCurrent = project.planAcceptance.records
              .filter(
                (record) =>
                  record.planRevision === project.plan!.revision &&
                  record.planDigest === planDigest &&
                  record.styleDigest === styleDigest,
              )
              .at(-1)
          }
        } catch {
          project.planAcceptanceUnavailable = true
          delete project.planAcceptance
          delete project.planAcceptanceCurrent
        }
      }
      let sourceAudit: PresentationProjectSnapshot['sourceAudit']
      try {
        sourceAudit = await restoreSourceAudit(project, documentId, controller.signal, check)
      } catch (error) {
        if (
          controller.signal.aborted ||
          captured !== epoch ||
          (error instanceof Error && error.message === 'presentation_document_changed')
        )
          throw error
        project.sourceAuditHistoryUnavailable = true
      }
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
      let deliveryBundles: PresentationDeliveryBundleReceipt[] | undefined
      let bundleNotice: string | undefined
      try {
        deliveryBundles = await readBundleList(project, documentId, controller.signal, check)
      } catch (error) {
        check()
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed', { cause: error })
        if (error instanceof Error && error.message === 'presentation_document_changed') throw error
        bundleNotice = '本机交付包暂时无法读取；已有成果保留，请刷新查看。'
      }
      check()
      publish({
        phase: 'idle',
        project,
        ...(sourceAudit ? { sourceAudit } : {}),
        ...(deliveryBundles ? { deliveryBundles } : {}),
        ...(bundleNotice ? { bundleNotice, deliveryBundlesUnavailable: true as const } : {}),
      })
      if (
        captured === epoch &&
        project.productionJob &&
        ['running', 'pausing', 'cancelling'].includes(project.productionJob.state)
      )
        poll = setTimeout(() => {
          void run('loading')
        }, 1500)
    } catch (error) {
      if (captured === epoch) {
        const code = error instanceof Error ? error.message : ''
        if (['presentation_document_changed', 'presentation_document_mismatch'].includes(code)) {
          selection = undefined
          projectDocument = undefined
        }
        const retain =
          previous &&
          boundDocument &&
          boundDocument === projectDocument &&
          [
            'presentation_service_unavailable',
            'presentation_unavailable',
            'presentation_busy',
            'presentation_selection_save_failed',
          ].includes(code) &&
          (await options.documentId().then(
            (id) => id === boundDocument,
            () => false,
          ))
        if (captured !== epoch) return
        publish({ phase: 'idle', ...(retain ? { project: previous } : {}), error: message(error) })
      }
    } finally {
      if (captured === epoch) active = undefined
    }
  }
  const readBoundResearch = async () => {
    const project = state.project
    const binding = project?.plan?.value.research
    if (active || !project || !binding || !projectDocument || !options.readResearchRecord) return
    const documentId = projectDocument
    const previous = state
    const controller = new AbortController()
    active = controller
    const captured = ++epoch
    const check = async () => {
      if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
      if (!options.available()) throw new Error('presentation_unavailable')
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
    }
    stopPolling()
    publish({ ...previous, phase: 'readingResearch', error: undefined })
    try {
      await check()
      await options.readResearchRecord(project.projectId, binding.ledgerId, controller.signal)
      await check()
      publish({
        ...previous,
        phase: 'idle',
        deliveryNotice: '计划绑定的指定研究已读取；研究整理与引用匹配不代表事实支持或 QA 通过。',
        error: undefined,
      })
    } catch (error) {
      if (captured !== epoch) return
      const changed = (await options.documentId().catch(() => undefined)) !== documentId
      if (captured !== epoch) return
      if (changed) {
        controller.abort()
        stop(message(new Error('presentation_document_changed')))
      } else publish({ ...previous, phase: 'idle', error: message(error) })
    } finally {
      if (captured === epoch) active = undefined
    }
  }
  const bundleAction = async (
    tool?: 'export_current_presentation_bundle' | 'restore_presentation_delivery_bundle',
    includePdf = false,
    bundleId?: string,
    deleteBundleId?: string,
    includePageScreenshots = false,
  ) => {
    const project = state.project
    const requestId = project?.production?.requestId
    if (
      active ||
      !project?.deliveryBundlesAvailable ||
      !requestId ||
      !projectDocument ||
      !options.available()
    )
      return
    if (tool === 'export_current_presentation_bundle' && !options.nativeDocumentExportAvailable?.())
      return
    if (
      tool === 'restore_presentation_delivery_bundle' &&
      (!bundleId || !/^[a-f0-9]{64}$/.test(bundleId))
    )
      return
    if (deleteBundleId !== undefined && !/^[a-f0-9]{64}$/.test(deleteBundleId)) return
    const documentId = projectDocument
    const previous = state
    stopPolling()
    const controller = new AbortController()
    active = controller
    const captured = ++epoch
    const check = () => {
      if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
      if (!options.available()) throw new Error('presentation_unavailable')
    }
    publish({ ...previous, phase: 'bundling', error: undefined, bundleNotice: undefined })
    try {
      check()
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
      let actionError: unknown
      let resultSummary: string | undefined
      if (deleteBundleId) {
        try {
          const response = await options.request(
            {
              operation: 'delivery_bundle_delete',
              documentId,
              projectId: project.projectId,
              requestId,
              bundleId: deleteBundleId,
            },
            controller.signal,
          )
          check()
          if ((await options.documentId()) !== documentId)
            throw new Error('presentation_document_changed')
          check()
          if (!response.ok) throw new Error('presentation_service_unavailable')
          const text = await response.text()
          check()
          if (new TextEncoder().encode(text).byteLength > 256 * 1024)
            throw new Error('presentation_response_invalid')
          const value = JSON.parse(text)
          if (
            !value ||
            Object.keys(value).sort().join(',') !== 'bundleId,deleted' ||
            value.bundleId !== deleteBundleId ||
            value.deleted !== true
          )
            throw new Error('presentation_response_invalid')
        } catch (error) {
          actionError = error
        }
      }
      if (tool) {
        try {
          const result = await options.executeTool(
            {
              id: `presentation-bundle-${captured}`,
              name: tool,
              input: {
                project_id: project.projectId,
                request_id: requestId,
                ...(tool === 'export_current_presentation_bundle'
                  ? {
                      include_pdf: includePdf,
                      ...(includePageScreenshots ? { include_page_screenshots: true } : {}),
                    }
                  : { bundle_id: bundleId }),
              },
            },
            controller.signal,
          )
          check()
          if (result.isError) throw new Error(result.output)
          resultSummary = result.summary
        } catch (error) {
          actionError = error
        }
      }
      check()
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      const deliveryBundles = await readBundleList(project, documentId, controller.signal, check)
      check()
      const settled = { ...previous }
      delete settled.deliveryBundlesUnavailable
      publish({
        ...settled,
        phase: 'idle',
        deliveryBundles,
        bundleNotice: actionError
          ? `${message(actionError)} 本次交付包操作未确认成功；本机可能已保存包，请刷新或恢复已有包，勿自动重复导出。`
          : deleteBundleId
            ? '本机交付包已删除；原始 PowerPoint 文稿不受影响。'
            : tool
              ? `${resultSummary ?? '交付包已放回会话附件。'} 历史 QA 与保存点不代表当前宿主验收，检查待完成。`
              : undefined,
        error: undefined,
      })
    } catch (error) {
      if (captured !== epoch) return
      const changed = (await options.documentId().catch(() => undefined)) !== documentId
      if (captured !== epoch) return
      publish(
        changed
          ? { phase: 'idle', error: '文档已改变，请在目标文档刷新项目。' }
          : {
              ...previous,
              phase: 'idle',
              deliveryBundlesUnavailable: true,
              bundleNotice:
                '交付包操作暂时无法确认；已有成果保留，请刷新本机包后恢复，勿自动重复导出。',
            },
      )
    } finally {
      if (captured === epoch) active = undefined
    }
  }
  const deliveryAction = async (tool: string, action?: PresentationIssueActionInput) => {
    if (active || !state.project?.production || !projectDocument) return
    const project = state.project
    if (
      tool === 'export_presentation_pdf' &&
      (project.production!.status !== 'compiled' || !options.productionPdfAvailable?.())
    )
      return
    const requestId = project.production!.requestId
    const documentId = projectDocument
    const report = state.deliveryReport
    if (action && (!report || report.requestId !== requestId)) return
    stopPolling()
    const controller = new AbortController()
    active = controller
    const captured = ++epoch
    publish({ phase: 'loading', project, ...(report ? { deliveryReport: report } : {}) })
    try {
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      if (captured !== epoch || controller.signal.aborted) return
      const result = await options.executeTool(
        {
          id: `presentation-delivery-${captured}`,
          name: tool,
          input: {
            project_id: project.projectId,
            request_id: requestId,
            ...(tool === 'export_presentation_pdf' ? { source: 'production' } : {}),
            ...(action ? { expected_revision: report!.issueLedger.revision, action } : {}),
          },
        },
        controller.signal,
      )
      if (captured !== epoch || controller.signal.aborted) return
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      if (captured !== epoch || controller.signal.aborted) return
      if (result.isError) throw new Error(result.output)
      if (tool === 'export_presentation_delivery_report' || tool === 'export_presentation_pdf') {
        publish({
          phase: 'idle',
          project,
          ...(report ? { deliveryReport: report } : {}),
          deliveryNotice:
            tool === 'export_presentation_pdf'
              ? 'PDF 预览已保存到会话附件，可在下方下载；它对应编译成果，不代表当前 PowerPoint 文档。'
              : 'JSON 和 Markdown 已保存到会话附件，可在附件区下载。',
        })
      } else {
        if (new TextEncoder().encode(result.output).byteLength > 8 * 1024 * 1024)
          throw new Error('presentation_response_invalid')
        const next = parsePresentationDeliveryReport(JSON.parse(result.output))
        if (
          next.documentId !== documentId ||
          next.projectId !== project.projectId ||
          next.requestId !== requestId
        )
          throw new Error('presentation_response_invalid')
        publish({ phase: 'idle', project, deliveryReport: next })
      }
    } catch (error) {
      if (captured !== epoch) return
      if (
        await options.documentId().then(
          (id) => id !== documentId,
          () => true,
        )
      ) {
        stop(message(new Error('presentation_document_changed')))
      } else if (captured === epoch) publish({ phase: 'idle', project, error: message(error) })
    } finally {
      if (captured === epoch) {
        active = undefined
        if (
          state.project?.productionJob &&
          ['running', 'pausing', 'cancelling'].includes(state.project.productionJob.state)
        )
          poll = setTimeout(() => {
            void run('loading')
          }, 1500)
      }
    }
  }
  const productionAction = (tool: string, requestId?: string, pageId?: string) => {
    const production = state.project?.production
    if (
      !production ||
      (requestId && requestId !== production.requestId) ||
      (pageId &&
        !production.pages.some((page) => page.id === pageId && page.state === 'compiled')) ||
      (tool === 'prepare_presentation_production_import' &&
        (production.status !== 'compiled' || production.revision !== undefined))
    )
      return Promise.resolve()
    return run('producing', production.requestId, tool, pageId)
  }
  const editPlan = async (expectedRevision: number, action: PresentationPlanEdit) => {
    if (active || !state.project?.plan || !projectDocument) return
    const project = state.project
    const documentId = projectDocument
    const saved = project.plan!
    if (expectedRevision !== saved.revision) {
      publish({ ...state, error: message(new Error('presentation_revision_conflict')) })
      return
    }
    stopPolling()
    const controller = new AbortController()
    active = controller
    const captured = ++epoch
    const check = () => {
      if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
      if (!options.available()) throw new Error('presentation_unavailable')
    }
    let committed = false
    let unchangedPlan: string | undefined
    publish({ phase: 'planning', project })
    try {
      check()
      let plan = structuredClone(saved.value)
      if (action.kind === 'restore') {
        const historical = saved.revisions?.find((event) => event.revision === action.revision)
        if (
          !Number.isSafeInteger(action.revision) ||
          action.revision < 1 ||
          action.revision >= saved.revision ||
          !historical
        )
          throw new Error('presentation_plan_revision_unavailable')
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        const read = await options.executeTool(
          {
            id: `presentation-plan-history-${captured}`,
            name: 'read_presentation_plan',
            input: { project_id: project.projectId, revision: action.revision },
          },
          controller.signal,
        )
        check()
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        if (read.isError) throw new Error(read.output)
        if (new TextEncoder().encode(read.output).byteLength > 512 * 1024)
          throw new Error('presentation_response_invalid')
        const value = JSON.parse(read.output)
        if (value.projectId !== project.projectId || value.revision !== action.revision)
          throw new Error('presentation_response_invalid')
        plan = parsePresentationPlan(value.plan)
        if ((await presentationDigest(canonicalPresentationValue(plan))) !== historical.inputDigest)
          throw new Error('presentation_response_invalid')
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        if (canonicalPresentationValue(plan) === canonicalPresentationValue(saved.value))
          unchangedPlan = canonicalPresentationValue(plan)
      } else {
        const index = plan.slides.findIndex((slide) => slide.id === action.pageId)
        if (index < 0) throw new Error('presentation_plan_page_missing')
        if (action.kind === 'lock') {
          if (typeof action.locked !== 'boolean') throw new Error('presentation_invalid_request')
          if ((plan.slides[index]!.locked === true) === action.locked) return
          if (action.locked) plan.slides[index]!.locked = true
          else delete plan.slides[index]!.locked
        } else if (plan.slides[index]!.locked) throw new Error('presentation_page_locked')
        else if (action.kind === 'delete') {
          if (plan.slides.length === 1) throw new Error('presentation_plan_last_page')
          plan.slides.splice(index, 1)
        } else if (action.kind === 'move' && ['up', 'down'].includes(action.direction)) {
          const target = index + (action.direction === 'up' ? -1 : 1)
          if (target < 0 || target >= plan.slides.length) return
          ;[plan.slides[index], plan.slides[target]] = [plan.slides[target]!, plan.slides[index]!]
        } else throw new Error('presentation_plan_page_missing')
      }
      const proposed = parsePresentationPlan(plan)
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
      if (!unchangedPlan) {
        const result =
          action.kind === 'lock'
            ? await (async () => {
                const response = await options.request(
                  {
                    operation: 'set_plan_page_lock',
                    documentId,
                    projectId: project.projectId,
                    expectedRevision,
                    pageId: action.pageId,
                    locked: action.locked,
                  },
                  controller.signal,
                )
                check()
                if (!response.ok) throw new Error('presentation_service_unavailable')
                const output = await response.text()
                if (new TextEncoder().encode(output).byteLength > 512 * 1024)
                  throw new Error('presentation_response_invalid')
                const value = JSON.parse(output)
                if (value.error)
                  throw new Error(
                    [
                      'revision_conflict',
                      'document_mismatch',
                      'page_locked',
                      'invalid_request',
                    ].includes(value.error)
                      ? `presentation_${value.error}`
                      : 'presentation_response_invalid',
                  )
                return { output, isError: false }
              })()
            : await options.executeTool(
                {
                  id: `presentation-plan-edit-${captured}`,
                  name: 'save_presentation_plan',
                  input: { expected_revision: expectedRevision, plan: proposed },
                },
                controller.signal,
              )
        check()
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        if (result.isError) throw new Error(result.output)
        if (new TextEncoder().encode(result.output).byteLength > 512 * 1024)
          throw new Error('presentation_response_invalid')
        const acknowledgement = JSON.parse(result.output)
        if (
          acknowledgement.projectId !== project.projectId ||
          acknowledgement.revision !== expectedRevision + 1 ||
          canonicalPresentationValue(parsePresentationPlan(acknowledgement.plan)) !==
            canonicalPresentationValue(proposed)
        )
          throw new Error('presentation_response_invalid')
        committed = true
      }
    } catch (error) {
      if (captured !== epoch) return
      if (
        await options.documentId().then(
          (id) => id !== documentId,
          () => true,
        )
      ) {
        if (captured === epoch) stop(message(new Error('presentation_document_changed')))
      } else if (captured === epoch) publish({ phase: 'idle', project, error: message(error) })
    } finally {
      if (captured === epoch) {
        active = undefined
        if (state.phase === 'planning') publish({ phase: 'idle', project })
        if (
          !committed &&
          project.productionJob &&
          ['running', 'pausing', 'cancelling'].includes(project.productionJob.state)
        )
          poll = setTimeout(() => {
            void run('loading')
          }, 1500)
      }
    }
    if ((!committed && !unchangedPlan) || captured !== epoch) return
    await run('loading')
    if (
      captured + 1 === epoch &&
      state.project?.projectId === project.projectId &&
      projectDocument === documentId
    ) {
      if (unchangedPlan) {
        if (state.error) return
        publish({
          ...state,
          ...(canonicalPresentationValue(state.project.plan?.value) === unchangedPlan
            ? { planNotice: '当前计划已与所选版本一致，无需新建修订；已有 PowerPoint 页面保留。' }
            : { error: message(new Error('presentation_revision_conflict')) }),
        })
        return
      }
      publish({
        ...state,
        planNotice: `调整已保存为计划第 ${expectedRevision + 1} 版；已有 PowerPoint 页面保留。${state.project.plan && state.project.plan.revision >= expectedRevision + 1 ? '继续制作时需使用当前计划；旧后台任务按原快照继续。' : '项目状态尚未刷新，请刷新后继续制作。'}`,
      })
    }
  }
  const acceptPlan = async (expectedRevision: number) => {
    if (active || !state.project?.plan || !projectDocument) return
    const project = state.project,
      plan = project.plan!,
      documentId = projectDocument
    if (expectedRevision !== plan.revision) {
      publish({ ...state, error: message(new Error('presentation_revision_conflict')) })
      return
    }
    if (!project.planAcceptance || project.planAcceptanceUnavailable) {
      publish({
        ...state,
        error: '接受决定记录暂不可用，请刷新项目或升级 PC 后重试；普通制作仍可继续。',
      })
      return
    }
    stopPolling()
    const controller = new AbortController(),
      captured = ++epoch
    active = controller
    publish({ phase: 'accepting', project })
    const check = () => {
      if (controller.signal.aborted || captured !== epoch) throw new Error('aborted')
    }
    let committed = false
    try {
      const planDigest = await presentationDigest(canonicalPresentationValue(plan.value)),
        decisionId = `accept-${await presentationDigest(canonicalPresentationValue({ projectId: project.projectId, documentId, planRevision: expectedRevision, planDigest }))}`
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
      const response = await options.request(
        {
          operation: 'accept_plan',
          projectId: project.projectId,
          documentId,
          decisionId,
          expectedRevision,
          planDigest,
        },
        controller.signal,
      )
      if (!response.ok) throw new Error('presentation_request_failed')
      const text = await response.text()
      check()
      if (new TextEncoder().encode(text).byteLength > 128 * 1024)
        throw new Error('presentation_response_invalid')
      const value = JSON.parse(text)
      if (value?.error) throw new Error(`presentation_${value.error}`)
      if (
        !value ||
        Object.keys(value).sort().join(',') !== 'acceptance,documentId,projectId' ||
        value.projectId !== project.projectId ||
        value.documentId !== documentId
      )
        throw new Error('presentation_response_invalid')
      const acceptance = parsePresentationPlanAcceptances({
        version: 1,
        projectId: project.projectId,
        documentId,
        records: [value.acceptance],
      }).records[0]!
      if (
        acceptance.decisionId !== decisionId ||
        acceptance.planRevision !== expectedRevision ||
        acceptance.planDigest !== planDigest ||
        acceptance.styleDigest !==
          (await presentationDigest(presentationPlanSnapshotInputs(plan.value).styleDigest))
      )
        throw new Error('presentation_response_invalid')
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
      committed = true
      publish({
        phase: 'idle',
        project,
        planNotice: `已记录对计划第 ${expectedRevision} 版与该版样式的接受决定；事实来源和页面质量仍需核验。`,
      })
    } catch (error) {
      if (captured === epoch) {
        if (error instanceof Error && error.message === 'presentation_document_changed')
          stop(message(error))
        else if (!controller.signal.aborted)
          publish({
            phase: 'idle',
            project,
            error: `${message(error)} 请刷新核对接受决定记录，不自动重复提交。`,
          })
      }
    } finally {
      if (captured === epoch) active = undefined
    }
    if (committed && captured === epoch) await run('loading')
  }
  const auditSources = async () => {
    if (active || !state.project?.plan || !projectDocument) return
    const project = state.project
    const plan = project.plan!
    const documentId = projectDocument
    const expected = plan.value.sources.flatMap((source) => {
      const attachmentId = presentationSourceAttachmentId(source)
      return attachmentId ? [{ sourceId: source.id, attachmentId }] : []
    })
    if (!expected.length) return
    stopPolling()
    const controller = new AbortController()
    active = controller
    const captured = ++epoch
    publish({ phase: 'auditing', project })
    try {
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      const result = await options.executeTool(
        {
          id: `presentation-source-audit-${captured}`,
          name: 'audit_presentation_sources',
          input: { project_id: project.projectId },
        },
        controller.signal,
      )
      if (captured !== epoch || controller.signal.aborted) return
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      if (result.isError) throw new Error(result.output)
      if (new TextEncoder().encode(result.output).byteLength > 128 * 1024)
        throw new Error('presentation_response_invalid')
      const value = JSON.parse(result.output) as {
        projectId?: unknown
        planRevision?: unknown
        sources?: {
          sourceId?: unknown
          attachmentId?: unknown
          status?: unknown
          offset?: unknown
        }[]
        checks?: unknown
      }
      if (
        value.projectId !== project.projectId ||
        value.planRevision !== plan.revision ||
        !Array.isArray(value.sources) ||
        value.sources.length !== expected.length ||
        JSON.stringify(value.checks) !==
          JSON.stringify({
            support: 'not_verified',
            sourceAuthority: 'not_verified',
            timeliness: 'not_verified',
          }) ||
        value.sources.some(
          (source, index) =>
            !source ||
            source.sourceId !== expected[index]!.sourceId ||
            source.attachmentId !== expected[index]!.attachmentId ||
            ![
              'found',
              'not_found',
              'empty_excerpt',
              'not_ready',
              'unsupported',
              'missing',
              'source_mismatch',
            ].includes(String(source.status)) ||
            (source.status === 'found'
              ? !Number.isSafeInteger(source.offset) ||
                Number(source.offset) < 0 ||
                Number(source.offset) > MAX_PRESENTATION_SOURCE_TEXT_CHARS
              : source.offset !== undefined),
        )
      )
        throw new Error('presentation_response_invalid')
      let updatedProject = project
      let sourceAudit: PresentationProjectSnapshot['sourceAudit'] = {
        planRevision: plan.revision,
        sources: value.sources as PresentationSourceAuditResult[],
      }
      if (project.sourceAuditHistory !== undefined) {
        const check = () => {
          if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
        }
        const response = await options.request(
          { operation: 'status', projectId: project.projectId, documentId },
          controller.signal,
        )
        check()
        if (!response.ok) throw new Error('presentation_service_unavailable')
        const text = await response.text()
        check()
        if (new TextEncoder().encode(text).byteLength > 256 * 1024)
          throw new Error('presentation_response_invalid')
        const refreshed = await parseStatus(JSON.parse(text), project.projectId)
        if (
          !refreshed.plan ||
          refreshed.plan.revision !== plan.revision ||
          canonicalPresentationValue(refreshed.plan.value) !==
            canonicalPresentationValue(plan.value)
        )
          throw new Error('presentation_revision_conflict')
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        if (refreshed.researchSummary && refreshed.researchSummary.documentId !== documentId) {
          delete refreshed.researchSummary
          refreshed.researchHistoryUnavailable = true
        }
        updatedProject = {
          ...project,
          sourceAuditHistory: refreshed.sourceAuditHistory,
          sourceAuditHistoryUnavailable: refreshed.sourceAuditHistoryUnavailable,
        }
        delete updatedProject.researchSummary
        delete updatedProject.researchHistoryUnavailable
        if (refreshed.researchSummary) updatedProject.researchSummary = refreshed.researchSummary
        if (refreshed.researchHistoryUnavailable) updatedProject.researchHistoryUnavailable = true
        try {
          sourceAudit =
            (await restoreSourceAudit(updatedProject, documentId, controller.signal, check)) ??
            sourceAudit
        } catch (error) {
          if (
            captured !== epoch ||
            controller.signal.aborted ||
            (error instanceof Error && error.message === 'presentation_document_changed')
          )
            throw error
          updatedProject.sourceAuditHistoryUnavailable = true
        }
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
      }
      publish({ phase: 'idle', project: updatedProject, sourceAudit })
    } catch (error) {
      if (captured !== epoch) return
      if (
        await options.documentId().then(
          (id) => id !== documentId,
          () => true,
        )
      )
        stop(message(new Error('presentation_document_changed')))
      else publish({ phase: 'idle', project, error: message(error) })
    } finally {
      if (captured === epoch) {
        active = undefined
        if (
          state.project?.productionJob &&
          ['running', 'pausing', 'cancelling'].includes(state.project.productionJob.state)
        )
          poll = setTimeout(() => {
            void run('loading')
          }, 1500)
      }
    }
  }
  const selectFeedbackComparisonBaseline = (requestId?: string) => {
    const project = state.project,
      production = project?.production
    if (active && !comparisonPending) return
    if (
      requestId !== undefined &&
      (!production ||
        production.status !== 'compiled' ||
        requestId === production.requestId ||
        !project?.productionTasks?.some(
          (task) =>
            task.requestId === requestId &&
            task.status === 'compiled' &&
            task.compiledCount === task.total,
        ))
    )
      return
    if (requestId === state.feedbackComparisonBaselineRequestId) return
    verifiedComparison = undefined
    if (comparisonPending) {
      epoch++
      active?.abort()
      active = undefined
      comparisonPending = false
    }
    publish({
      ...state,
      phase: 'idle',
      feedbackComparisonBaselineRequestId: requestId,
      feedbackComparison: undefined,
      feedbackComparisonUnavailable: undefined,
      error: undefined,
    })
  }
  const readFeedbackComparison = async () => {
    const previous = state,
      project = previous.project,
      production = project?.production,
      baselineRequestId = previous.feedbackComparisonBaselineRequestId,
      documentId = projectDocument
    const baseline = project?.productionTasks?.find((task) => task.requestId === baselineRequestId)
    if (
      active ||
      !options.available() ||
      !project ||
      !production ||
      production.status !== 'compiled' ||
      !baselineRequestId ||
      baselineRequestId === production.requestId ||
      !baseline ||
      baseline.status !== 'compiled' ||
      baseline.compiledCount !== baseline.total ||
      !documentId
    )
      return
    const fingerprint = JSON.stringify({ production, baseline })
    stopPolling()
    const controller = new AbortController(),
      captured = ++epoch
    active = controller
    comparisonPending = true
    publish({ ...previous, phase: 'loading', error: undefined })
    const check = () => {
      if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
      if (!options.available()) throw new Error('presentation_unavailable')
      if (
        state.project?.projectId !== project.projectId ||
        state.feedbackComparisonBaselineRequestId !== baselineRequestId ||
        JSON.stringify({
          production: state.project.production,
          baseline: state.project.productionTasks?.find(
            (task) => task.requestId === baselineRequestId,
          ),
        }) !== fingerprint
      )
        throw new Error('presentation_request_changed')
    }
    const current = async () => {
      check()
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
    }
    try {
      await current()
      check()
      const response = await options.request(
        {
          operation: 'production_feedback_compare',
          documentId,
          projectId: project.projectId,
          requestId: production.requestId,
          baselineRequestId,
        },
        controller.signal,
      )
      check()
      await current()
      check()
      const text = await response.text()
      check()
      await current()
      check()
      if (
        new TextEncoder().encode(text).byteLength >
        MAX_PRESENTATION_FEEDBACK_COMPARISON_RESPONSE_BYTES
      )
        throw new Error('presentation_response_invalid')
      const value = JSON.parse(text)
      if (value?.error)
        throw new Error(
          ['invalid_request', 'unsupported', 'unsupported_operation', 'upgrade_required'].includes(
            value.error,
          )
            ? 'presentation_upgrade_required'
            : 'presentation_service_unavailable',
        )
      if (
        !response.ok ||
        !value ||
        typeof value !== 'object' ||
        Array.isArray(value) ||
        Object.keys(value).join(',') !== 'comparison'
      )
        throw new Error('presentation_response_invalid')
      const comparison = parsePresentationFeedbackComparison(value.comparison)
      if (
        comparison.documentId !== documentId ||
        comparison.projectId !== project.projectId ||
        comparison.candidate.requestId !== production.requestId ||
        comparison.baseline.requestId !== baselineRequestId ||
        comparison.candidate.planRevision !== production.planRevision ||
        comparison.baseline.planRevision !== baseline.planRevision ||
        comparison.baseline.pages.length !== baseline.total ||
        JSON.stringify(comparison.candidate.pages.map((page) => page.pageId)) !==
          JSON.stringify(production.pages.map((page) => page.id))
      )
        throw new Error('presentation_response_invalid')
      if (
        verifiedFeedback?.documentId === documentId &&
        verifiedFeedback.projectId === project.projectId &&
        verifiedFeedback.requestId === production.requestId
      ) {
        const selected = comparison.candidate,
          latest = verifiedFeedback.snapshots.at(-1)!
        if (
          selected.inputDigest !== verifiedFeedback.inputDigest ||
          selected.planDigest !== verifiedFeedback.planDigest ||
          selected.feedbackRevision === null ||
          selected.feedbackRevision < verifiedFeedback.revision ||
          selected.feedbackRecordedAt === null ||
          selected.feedbackRecordedAt < latest.recordedAt ||
          (selected.feedbackRevision === verifiedFeedback.revision &&
            (selected.feedbackRecordedAt !== latest.recordedAt ||
              canonicalPresentationValue(selected.pages) !==
                canonicalPresentationValue(
                  latest.pages.map(({ pageId, status }) => ({ pageId, status })),
                )))
        )
          throw new Error('presentation_response_invalid')
      }
      if (
        verifiedComparison?.documentId === documentId &&
        verifiedComparison.projectId === project.projectId &&
        verifiedComparison.baseline.requestId === baselineRequestId &&
        verifiedComparison.candidate.requestId === production.requestId
      ) {
        for (const role of ['baseline', 'candidate'] as const) {
          const old = verifiedComparison[role],
            next = comparison[role]
          if (
            old.inputDigest !== next.inputDigest ||
            old.planDigest !== next.planDigest ||
            old.planRevision !== next.planRevision ||
            canonicalPresentationValue(old.plan) !== canonicalPresentationValue(next.plan) ||
            (old.feedbackRevision !== null &&
              (next.feedbackRevision === null ||
                next.feedbackRevision < old.feedbackRevision ||
                (next.feedbackRevision === old.feedbackRevision &&
                  (next.feedbackRecordedAt !== old.feedbackRecordedAt ||
                    canonicalPresentationValue(next.pages) !==
                      canonicalPresentationValue(old.pages))) ||
                (next.feedbackRecordedAt !== null &&
                  old.feedbackRecordedAt !== null &&
                  next.feedbackRecordedAt < old.feedbackRecordedAt)))
          )
            throw new Error('presentation_response_invalid')
        }
      }
      await current()
      check()
      verifiedComparison = structuredClone(comparison)
      publish({
        ...previous,
        phase: 'idle',
        feedbackComparison: comparison,
        feedbackComparisonUnavailable: undefined,
        error: undefined,
      })
    } catch (error) {
      if (captured !== epoch) return
      if (
        await options.documentId().then(
          (id) => id !== documentId,
          () => true,
        )
      ) {
        if (captured === epoch) stop(message(new Error('presentation_document_changed')))
        return
      }
      if (captured !== epoch) return
      publish({
        ...previous,
        phase: 'idle',
        feedbackComparison: undefined,
        feedbackComparisonUnavailable: true,
        error:
          error instanceof Error && error.message === 'presentation_upgrade_required'
            ? '当前 PC 不支持制作反馈对照，请升级后重新读取。'
            : '制作反馈对照暂时不可读取，请核对当前任务与基线版本后重试。',
      })
    } finally {
      if (captured === epoch) {
        active = undefined
        comparisonPending = false
      }
    }
  }
  const feedbackAction = async (patch?: PresentationProductionFeedbackPage[]) => {
    const previous = state,
      project = previous.project,
      production = project?.production,
      documentId = projectDocument
    if (
      active ||
      !options.available() ||
      !project ||
      !production ||
      production.status !== 'compiled' ||
      !documentId ||
      (patch && previous.productionFeedback === undefined)
    )
      return
    const fingerprint = JSON.stringify(production),
      saved =
        verifiedFeedback?.documentId === documentId &&
        verifiedFeedback.projectId === project.projectId &&
        verifiedFeedback.requestId === production.requestId
          ? structuredClone(verifiedFeedback)
          : previous.productionFeedback === undefined
            ? undefined
            : structuredClone(previous.productionFeedback)
    const feedbackFingerprint = JSON.stringify(previous.productionFeedback)
    stopPolling()
    const controller = new AbortController(),
      captured = ++epoch
    active = controller
    publish({ ...previous, phase: 'loading', error: undefined })
    const check = () => {
      if (captured !== epoch || controller.signal.aborted) throw new Error('cancelled')
      if (!options.available()) throw new Error('presentation_unavailable')
      if (JSON.stringify(state.productionFeedback) !== feedbackFingerprint)
        throw new Error('presentation_revision_conflict')
      if (
        state.project?.projectId !== project.projectId ||
        JSON.stringify(state.project.production) !== fingerprint
      )
        throw new Error('presentation_request_changed')
    }
    const current = async () => {
      check()
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
    }
    try {
      const pages = patch ? parsePresentationProductionFeedbackPages(patch) : undefined
      if (pages?.some((page) => !production.pages.some((item) => item.id === page.pageId)))
        throw new Error('presentation_response_invalid')
      await current()
      check()
      const response = await options.request(
        {
          operation: patch ? 'production_feedback_record' : 'production_feedback_read',
          documentId,
          projectId: project.projectId,
          requestId: production.requestId,
          ...(pages ? { expectedRevision: saved?.revision ?? 0, pages } : {}),
        },
        controller.signal,
      )
      check()
      await current()
      check()
      const text = await response.text()
      check()
      await current()
      check()
      if (
        new TextEncoder().encode(text).byteLength >
        MAX_PRESENTATION_PRODUCTION_FEEDBACK_BYTES +
          new TextEncoder().encode('{"feedback":}').byteLength
      )
        throw new Error('presentation_response_invalid')
      const value = JSON.parse(text)
      if (!response.ok || value?.error)
        throw new Error(
          value?.error === 'revision_conflict'
            ? 'presentation_revision_conflict'
            : 'presentation_service_unavailable',
        )
      if (
        !value ||
        typeof value !== 'object' ||
        Array.isArray(value) ||
        Object.keys(value).join(',') !== 'feedback' ||
        (patch && value.feedback === null)
      )
        throw new Error('presentation_response_invalid')
      const feedback =
        value.feedback === null ? null : parsePresentationProductionFeedbackLedger(value.feedback)
      if (
        feedback &&
        (feedback.documentId !== documentId ||
          feedback.projectId !== project.projectId ||
          feedback.requestId !== production.requestId ||
          feedback.planRevision !== production.planRevision ||
          JSON.stringify(feedback.pageIds) !==
            JSON.stringify(production.pages.map((page) => page.id)))
      )
        throw new Error('presentation_response_invalid')
      if (
        !pages &&
        saved &&
        (!feedback ||
          feedback.revision < saved.revision ||
          feedback.inputDigest !== saved.inputDigest ||
          feedback.planDigest !== saved.planDigest ||
          canonicalPresentationValue(feedback.snapshots.slice(0, saved.snapshots.length)) !==
            canonicalPresentationValue(saved.snapshots))
      )
        throw new Error('presentation_response_invalid')
      if (pages && feedback) {
        if (
          feedback.revision !== (saved?.revision ?? 0) + 1 ||
          (saved &&
            (feedback.inputDigest !== saved.inputDigest ||
              feedback.planDigest !== saved.planDigest ||
              JSON.stringify(feedback.snapshots.slice(0, -1)) !== JSON.stringify(saved.snapshots)))
        )
          throw new Error('presentation_response_invalid')
        const expected = production.pages.map(
          (page) =>
            pages.find((item) => item.pageId === page.id) ??
            saved?.snapshots.at(-1)?.pages.find((item) => item.pageId === page.id) ?? {
              pageId: page.id,
              status: 'not_evaluated',
            },
        )
        if (
          canonicalPresentationValue(expected) !==
          canonicalPresentationValue(feedback.snapshots.at(-1)?.pages)
        )
          throw new Error('presentation_response_invalid')
      }
      await current()
      check()
      verifiedFeedback = feedback ? structuredClone(feedback) : undefined
      publish({
        ...previous,
        phase: 'idle',
        productionFeedback: feedback,
        productionFeedbackUnavailable: undefined,
        error: undefined,
      })
    } catch (error) {
      if (captured !== epoch) return
      if (
        await options.documentId().then(
          (id) => id !== documentId,
          () => true,
        )
      ) {
        if (captured === epoch) stop(message(new Error('presentation_document_changed')))
        return
      }
      if (captured !== epoch) return
      publish({
        ...previous,
        phase: 'idle',
        productionFeedback: undefined,
        productionFeedbackUnavailable: true,
        error:
          error instanceof Error && error.message === 'presentation_revision_conflict'
            ? '反馈已有更新，请重新读取后保存。'
            : '人工修正反馈暂时不可读取或保存，请重新读取当前任务；页面与已有检查保留。',
      })
    } finally {
      if (captured === epoch) active = undefined
    }
  }
  return {
    selectFeedbackComparisonBaseline,
    readFeedbackComparison,
    readProductionFeedback: () => feedbackAction(),
    recordProductionFeedback: (pages) => feedbackAction(pages),
    readBoundResearch,
    currentBundleAvailable: () =>
      options.available() && options.nativeDocumentExportAvailable?.() === true,
    exportCurrentBundle: (includePdf = false, includePageScreenshots = false) =>
      bundleAction(
        'export_current_presentation_bundle',
        includePdf,
        undefined,
        undefined,
        includePageScreenshots,
      ),
    restoreDeliveryBundle: (bundleId) =>
      bundleAction('restore_presentation_delivery_bundle', false, bundleId),
    readDeliveryBundles: () => bundleAction(),
    deleteDeliveryBundle: (bundleId) => bundleAction(undefined, false, undefined, bundleId),
    acceptPlan,
    editPlan,
    pdfAvailable: () => options.available() && options.productionPdfAvailable?.() === true,
    exportProductionPdf: () => deliveryAction('export_presentation_pdf'),
    auditSources,
    readDeliveryReport: () => deliveryAction('read_presentation_delivery_report'),
    exportDeliveryReport: () => deliveryAction('export_presentation_delivery_report'),
    recordIssueAction: (action) => deliveryAction('record_presentation_issue_action', action),
    selectProduction: (requestId) => {
      if (
        !validId(requestId) ||
        !state.project?.productionTasks?.some((task) => task.requestId === requestId)
      )
        return Promise.resolve()
      return run('loading', requestId)
    },
    startProductionJob: (requestId) =>
      productionAction('start_presentation_production_job', requestId),
    pauseProductionJob: (requestId) =>
      productionAction('pause_presentation_production_job', requestId),
    resumeProductionJob: (requestId) =>
      productionAction('resume_presentation_production_job', requestId),
    cancelProductionJob: (requestId) =>
      productionAction('cancel_presentation_production_job', requestId),
    downloadProductionPage: (pageId) =>
      productionAction('read_presentation_page_artifact', undefined, pageId),
    prepareProduction: () => productionAction('prepare_presentation_production_import'),
    snapshot: () => state,
    subscribe(listener) {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    refresh: () => run('loading'),
    restore: () => run('restoring'),
    resume: (requestId) => run('resuming', requestId),
    runProduction: (requestId) => run('producing', requestId),
    cancel: () => stop(message(new Error('cancelled'))),
    clear: () => stop(undefined, true),
    prepareReconnect: () => {
      ignoreStoredSelection = false
      recoverInterruptedJob = true
    },
  }
}
