import {
  parsePresentationDeliveryReport,
  type PresentationDeliveryReport,
} from '@wiswork/pptx-engine/presentation-delivery-report'
import type { PresentationIssueActionInput } from '@wiswork/project-store/presentation-issue'
import type { PresentationProductionJob } from '@wiswork/project-store/presentation-job'
import { parsePresentationJobResponse } from './presentation-jobs.js'
import {
  parsePresentationProductionStatus,
  type PresentationProductionStatus,
} from './presentation-production.js'
import {
  parsePresentationPlan,
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
}
export interface PresentationProjectStatus {
  productionTasks?: PresentationProductionTask[]
  productionJob?: PresentationProductionJob | null
  jobsUnavailable?: boolean
  production?: PresentationProductionStatus
  projectId: string
  title: string
  status: 'planned' | 'pending' | 'compiled'
  plan?: { revision: number; value: PresentationPlan }
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
  deliveryReport?: PresentationDeliveryReport
  deliveryNotice?: string
  phase: 'idle' | 'loading' | 'restoring' | 'resuming' | 'producing'
  project?: PresentationProjectStatus
  error?: string
}
export interface PresentationProjectController {
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
}
const validId = (value: unknown): value is string =>
  typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const pageCount = (value: unknown): value is number =>
  Number.isSafeInteger(value) && Number(value) >= 1 && Number(value) <= 100
function parseStatus(value: unknown, projectId: string): PresentationProjectStatus {
  const p = value as PresentationProjectStatus | undefined
  let plan: PresentationProjectStatus['plan']
  if (p?.plan !== undefined) {
    if (!p.plan || !Number.isSafeInteger(p.plan.revision) || p.plan.revision < 1)
      throw new Error('presentation_response_invalid')
    plan = { revision: p.plan.revision, value: parsePresentationPlan(p.plan.value) }
    if (plan.value.projectId !== projectId) throw new Error('presentation_response_invalid')
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
  if (code === 'cancelled' || code === 'presentation_aborted')
    return '已停止等待；PC 可能已保存结果，可刷新查看。'
  return '暂时无法恢复项目，已有成果已保留。请检查 PC 连接后重试。'
}

export function createPresentationProjectController(
  options: Pick<
    PresentationGenerationOptions,
    'request' | 'available' | 'documentId' | 'lastProject'
  > &
    Pick<AgentSkill, 'executeTool'>,
): PresentationProjectController {
  let state: PresentationProjectSnapshot = { phase: 'idle' }
  const listeners = new Set<() => void>()
  let epoch = 0
  let projectDocument: string | undefined
  let selection: { documentId: string; projectId: string; requestId: string } | undefined
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
  const stop = (error?: string) => {
    stopPolling()
    projectDocument = undefined
    selection = undefined
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
    const previous = state.project
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
      const project = parseStatus(value, projectId)
      let selectedRequest = phase === 'loading' && requestId ? requestId : selection?.requestId
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
        }
      }
      if ((await options.documentId()) !== documentId)
        throw new Error('presentation_document_changed')
      check()
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
      if (tool === 'export_presentation_delivery_report') {
        publish({
          phase: 'idle',
          project,
          ...(report ? { deliveryReport: report } : {}),
          deliveryNotice: 'JSON 和 Markdown 已保存到会话附件，可在附件区下载。',
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
  return {
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
    clear: () => stop(),
  }
}
