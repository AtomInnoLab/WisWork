import { resolve } from 'node:path'
import type { PresentationStore } from '@wiswork/project-store'
import type { PresentationProductionJobEventInput } from '@wiswork/project-store/presentation-job'
import {
  handlePresentationProduction,
  presentationProductionSummary,
} from './presentation-production'

// One Electron main process owns these workers, including across service instances.
const workers = new Map<
  string,
  {
    requestId: string
    projectId: string
    documentId: string
    controller: AbortController
    settled: Promise<void>
  }
>()
export async function stopPresentationWorkers(scope: {
  root: string
  projectId: string
  documentId: string
}): Promise<void> {
  const key = `${resolve(scope.root)}\0${scope.projectId}`
  const entry = workers.get(key)
  if (!entry || entry.projectId !== scope.projectId || entry.documentId !== scope.documentId) return
  entry.controller.abort()
  await entry.settled
}
export const presentationJobOperations = [
  'production_job_start',
  'production_job_status',
  'production_job_pause',
  'production_job_resume',
  'production_job_cancel',
]
export function activePresentationRequest(key: string): string | undefined {
  return workers.get(key)?.requestId
}
export function hasPresentationWorker(key: string): boolean {
  return workers.has(key)
}
export function handlePresentationJob(
  key: string,
  request: Record<string, unknown>,
  options: Parameters<typeof handlePresentationProduction>[1] & { store: PresentationStore },
) {
  request = structuredClone(request)
  options = { ...options }
  const { store } = options
  const controller = new AbortController()
  const assertWritable = options.assertWritable
  const beforeWrite = () => {
    if (controller.signal.aborted) throw new Error('aborted')
    assertWritable?.()
    if (controller.signal.aborted) throw new Error('aborted')
  }
  const projectId = request.projectId as string,
    documentId = request.documentId as string,
    requestId = request.requestId as string
  const production = () => {
    const record = store.production(projectId, documentId, requestId)
    if (!record) throw new Error('not_found')
    return record
  }
  production() // Validate document binding before inspecting a process-wide worker.
  const read = () => store.productionJob(projectId, documentId, requestId)
  const append = (event: PresentationProductionJobEventInput) => {
    beforeWrite()
    return store.appendProductionJobEvent(
      projectId,
      documentId,
      requestId,
      read()?.revision ?? 0,
      event,
    )
  }
  let job = read()
  const worker = workers.get(key)
  if (!worker || worker.requestId !== requestId) {
    if (job?.state === 'running') job = append({ type: 'run.interrupted' })
    else if (job?.state === 'pausing') job = append({ type: 'run.paused' })
    else if (job?.state === 'cancelling') job = append({ type: 'run.cancelled' })
  }
  const operation = request.operation
  if (operation === 'production_job_start' || operation === 'production_job_resume') {
    if (worker && worker.requestId !== requestId) throw new Error('busy')
    if (!worker && job?.state !== 'completed' && job?.state !== 'cancelled') {
      if (
        operation === 'production_job_start'
          ? !!job
          : !job || !['paused', 'interrupted', 'failed'].includes(job.state)
      )
        throw new Error('invalid_state')
      job = append({ type: 'run.started' })
      const entry = { requestId, projectId, documentId, controller, settled: Promise.resolve() }
      workers.set(key, entry)
      // Schedule after acceptance; the client's AbortSignal never belongs to the worker.
      entry.settled = Promise.resolve().then(async () => {
        try {
          await handlePresentationProduction(
            { ...request, operation: 'production_run' },
            {
              ...options,
              shouldStop: () => read()?.state !== 'running',
              onPage: (_record, page) => {
                if (page.state === 'building')
                  append({ type: 'page.started', pageId: page.pageId, attempt: page.attempt })
                else if (page.state === 'compiled')
                  append({ type: 'page.compiled', pageId: page.pageId, attempt: page.attempt })
                else if (page.state === 'failed')
                  append({
                    type: 'page.failed',
                    pageId: page.pageId,
                    attempt: page.attempt,
                    error: page.error as Extract<
                      PresentationProductionJobEventInput,
                      { type: 'page.failed' }
                    >['error'],
                  })
              },
            },
            controller.signal,
          )
          const state = read()?.state
          if (state === 'pausing' && production().pages.every((page) => page.state === 'compiled'))
            append({ type: 'run.completed' })
          else if (state === 'pausing') append({ type: 'run.paused' })
          else if (state === 'cancelling') append({ type: 'run.cancelled' })
          else if (production().pages.every((page) => page.state === 'compiled'))
            append({ type: 'run.completed' })
          else append({ type: 'run.failed', error: 'compile_failed' })
        } catch {
          // If storage itself is unavailable, leave the receipt for recovery on the next read.
          try {
            append({ type: 'run.failed', error: 'invalid_state' })
          } catch {
            /* durable recovery */
          }
        } finally {
          if (workers.get(key) === entry) workers.delete(key)
        }
      })
    }
  } else if (operation === 'production_job_pause') {
    if (!job) throw new Error('invalid_state')
    if (job.state === 'running') job = append({ type: 'run.pause_requested' })
  } else if (operation === 'production_job_cancel') {
    if (!job) throw new Error('invalid_state')
    if (!['cancelled', 'completed', 'cancelling'].includes(job.state))
      job = append({ type: 'run.cancel_requested' })
    if ((!worker || worker.requestId !== requestId) && job.state === 'cancelling')
      job = append({ type: 'run.cancelled' })
  }
  const record = production()
  return {
    job: job ?? null,
    production: {
      ...presentationProductionSummary(record),
      inputDigest: record.inputDigest,
      planDigest: record.planDigest,
    },
  }
}
