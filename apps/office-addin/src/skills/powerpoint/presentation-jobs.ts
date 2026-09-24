import type { AgentSkill, AgentToolDef } from '@wiswork/agent-core'
import {
  parsePresentationProductionJob,
  type PresentationProductionJob,
} from '@wiswork/project-store/presentation-job'
import type { PresentationGenerationOptions } from './presentation-generation.js'
import { parsePresentationProductionStatus } from './presentation-production.js'
const operations = {
  start_presentation_production_job: 'production_job_start',
  read_presentation_production_job: 'production_job_status',
  pause_presentation_production_job: 'production_job_pause',
  resume_presentation_production_job: 'production_job_resume',
  cancel_presentation_production_job: 'production_job_cancel',
} as const
const validId = (v: unknown): v is string =>
  typeof v === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(v)
export function parsePresentationJobResponse(
  value: unknown,
  documentId: string,
  projectId: string,
  requestId: string,
) {
  const v = value as { job: unknown; production: unknown } | null
  if (!v || Object.keys(v).some((k) => !['job', 'production'].includes(k)))
    throw new Error('presentation_response_invalid')
  const rawProduction = v.production as Record<string, unknown>
  const { inputDigest, planDigest, ...summary } = rawProduction ?? {}
  if (
    ![inputDigest, planDigest].every(
      (digest) => typeof digest === 'string' && /^[a-f0-9]{64}$/.test(digest),
    )
  )
    throw new Error('presentation_response_invalid')
  const production = parsePresentationProductionStatus(summary)
  const job: PresentationProductionJob | null =
    v.job === null ? null : parsePresentationProductionJob(v.job)
  if (
    production.projectId !== projectId ||
    production.requestId !== requestId ||
    (job &&
      (job.projectId !== projectId ||
        job.documentId !== documentId ||
        job.requestId !== requestId ||
        job.planRevision !== production.planRevision ||
        (job.state === 'completed' && production.status !== 'compiled') ||
        job.inputDigest !== inputDigest ||
        job.planDigest !== planDigest ||
        job.events.some(
          (event) =>
            'pageId' in event && !production.pages.some((page) => page.id === event.pageId),
        )))
  )
    throw new Error('presentation_response_invalid')
  return { job, production }
}
export function createPresentationJobsSkill(
  options: Pick<PresentationGenerationOptions, 'request' | 'available' | 'documentId'>,
): AgentSkill & { clear(): void } {
  let epoch = 0
  const tools: AgentToolDef[] = Object.keys(operations).map((name) => ({
    name,
    description:
      'Control or read durable PC compilation of frozen presentation pages. Background work survives client disconnection. Pause/cancel take effect after the current page; preserved compiled pages are not imported or QA checked. Resume only paused, interrupted or failed work. Use saved project/request IDs.',
    inputSchema: {
      type: 'object',
      properties: { project_id: { type: 'string' }, request_id: { type: 'string' } },
      required: ['project_id', 'request_id'],
      additionalProperties: false,
    },
  }))
  return {
    id: 'office-presentation-jobs',
    systemPrompt:
      'Use background presentation production jobs for recoverable compilation of already frozen pages. Read status on reconnect. Pausing and cancelling may finish the current page first. Stopping a client wait does not cancel PC work. Events describe page compilation only, never research, host delivery or QA.',
    get tools() {
      return options.available() ? tools : []
    },
    clear() {
      epoch++
    },
    async executeTool(call, signal) {
      const captured = epoch
      const check = () => {
        if (captured !== epoch || signal?.aborted) throw new Error('cancelled')
        if (!options.available()) throw new Error('presentation_unavailable')
      }
      try {
        check()
        const operation = operations[call.name as keyof typeof operations],
          input = call.input
        if (
          !operation ||
          call.inputError ||
          call.truncated ||
          Object.keys(input).some((k) => !['project_id', 'request_id'].includes(k)) ||
          !validId(input.project_id) ||
          !validId(input.request_id)
        )
          throw new Error('invalid_tool_input')
        const documentId = await options.documentId()
        check()
        const response = await options.request(
          { operation, documentId, projectId: input.project_id, requestId: input.request_id },
          signal,
        )
        check()
        if (!response.ok) throw new Error('presentation_service_unavailable')
        const text = await response.text()
        check()
        if (new TextEncoder().encode(text).byteLength > 256 * 1024)
          throw new Error('presentation_response_invalid')
        const value = JSON.parse(text)
        if (typeof value?.error === 'string' && /^[a-z_]{1,80}$/.test(value.error))
          throw new Error(
            value.error === 'invalid_request'
              ? 'presentation_upgrade_required'
              : `presentation_${value.error}`,
          )
        const result = parsePresentationJobResponse(
          value,
          documentId,
          input.project_id,
          input.request_id,
        )
        if ((await options.documentId()) !== documentId)
          throw new Error('presentation_document_changed')
        check()
        return {
          output: JSON.stringify(result),
          mutated: false,
          summary: '已读取 PC 后台页编译任务；尚未导入或验收',
        }
      } catch (error) {
        const raw = error instanceof Error ? error.message : ''
        return {
          output:
            raw === 'cancelled' ||
            raw === 'invalid_tool_input' ||
            /^presentation_[a-z_]{1,80}$/.test(raw)
              ? raw
              : 'presentation_response_invalid',
          isError: true,
          mutated: false,
          summary: '后台任务操作未完成；已有成果保留，可刷新查看',
        }
      }
    },
  }
}
