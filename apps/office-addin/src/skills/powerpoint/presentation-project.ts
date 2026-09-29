import { MAX_PRESENTATION_SOURCE_TEXT_CHARS } from '@wiswork/pptx-engine/presentation-source-limits'
import {
  parsePresentationDeliveryReport,
  type PresentationDeliveryReport,
} from '@wiswork/pptx-engine/presentation-delivery-report'
import type { PresentationIssueActionInput } from '@wiswork/project-store/presentation-issue'
import type { PresentationProductionJob } from '@wiswork/project-store/presentation-job'
import type { PresentationPlanRevisionSnapshot } from '@wiswork/project-store'
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
  sourceAudit?: {
    planRevision: number
    sources: {
      sourceId: string
      attachmentId: string
      status: 'found' | 'not_found' | 'empty_excerpt' | 'not_ready' | 'unsupported' | 'missing'
      offset?: number
    }[]
  }
  deliveryReport?: PresentationDeliveryReport
  deliveryNotice?: string
  phase: 'idle' | 'loading' | 'restoring' | 'resuming' | 'producing' | 'auditing'
  project?: PresentationProjectStatus
  error?: string
}
export interface PresentationProjectController {
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
async function parseStatus(value: unknown, projectId: string): Promise<PresentationProjectStatus> {
  const p = value as PresentationProjectStatus | undefined
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
      const digest = async (input: string) => {
        const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(input))
        return Array.from(new Uint8Array(bytes), (byte) => byte.toString(16).padStart(2, '0')).join(
          '',
        )
      }
      for (const key of Object.keys(inputs) as (keyof typeof inputs)[]) {
        const hash = await digest(inputs[key])
        if (hash !== latestSnapshot[key]) throw new Error('presentation_response_invalid')
      }
      if (
        (await digest(canonicalPresentationValue(p.plan.value))) !== revisions!.at(-1)!.inputDigest
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
  // Copy only the bounded public projection; never retain arbitrary server fields or binary data.
  return {
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
  if (
    ['presentation_upgrade_required', 'presentation_invalid_request', 'invalid_request'].includes(
      code,
    )
  )
    return '当前 PC 尚不支持此项目操作，请升级 WisWork PC 后重试。'
  if (code === 'presentation_revision_conflict') return '记录已有更新，请重新读取后继续。'
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
    Pick<AgentSkill, 'executeTool'>,
): PresentationProjectController {
  let state: PresentationProjectSnapshot = { phase: 'idle' }
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
  const stop = (error?: string, resetSelection = false) => {
    stopPolling()
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
      projectDocument = documentId
      publish({ phase: 'idle', project })
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
            ].includes(String(source.status)) ||
            (source.status === 'found'
              ? !Number.isSafeInteger(source.offset) ||
                Number(source.offset) < 0 ||
                Number(source.offset) > MAX_PRESENTATION_SOURCE_TEXT_CHARS
              : source.offset !== undefined),
        )
      )
        throw new Error('presentation_response_invalid')
      publish({
        phase: 'idle',
        project,
        sourceAudit: {
          planRevision: plan.revision,
          sources: value.sources as NonNullable<
            PresentationProjectSnapshot['sourceAudit']
          >['sources'],
        },
      })
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
  return {
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
