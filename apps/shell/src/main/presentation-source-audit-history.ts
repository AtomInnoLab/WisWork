import { randomUUID } from 'node:crypto'
import { parsePresentationPlan } from '@wiswork/pptx-engine/presentation-plan'
import type { PresentationStore } from '@wiswork/project-store'
import type { PresentationSourceAuditRun } from '@wiswork/project-store/presentation-source-audit'
import { auditPresentationSources } from './presentation-source-audit'

export async function handlePresentationSourceAudit(
  request: Record<string, unknown>,
  store: PresentationStore,
  attachments: (body: Record<string, unknown>, signal: AbortSignal) => Promise<unknown>,
  signal: AbortSignal,
  assertWritable?: () => void,
): Promise<unknown> {
  const projectId = request.projectId as string,
    documentId = request.documentId as string
  const auditId = request.auditId as string | undefined
  if (request.operation === 'read_source_audit') {
    const audit = store.sourceAudit(projectId, documentId, auditId!)
    if (!audit) throw new Error('not_found')
    return { projectId, documentId, audit }
  }
  assertWritable?.()
  const run = store.beginSourceAudit(projectId, documentId, auditId ?? randomUUID())
  if (run.state === 'failed') throw new Error(run.error)
  let completed: PresentationSourceAuditRun = run
  if (run.state === 'running') {
    try {
      const saved = store.plan(projectId, documentId)
      if (!saved || saved.revision !== run.planRevision || saved.inputDigest !== run.planDigest)
        throw new Error('request_conflict')
      const sources = await auditPresentationSources(
        parsePresentationPlan(saved.plan),
        documentId,
        attachments,
        signal,
      )
      if (signal.aborted) throw new Error('aborted')
      assertWritable?.()
      completed = store.finishSourceAudit(projectId, documentId, run.id, { sources })
    } catch (error) {
      assertWritable?.()
      const code =
        signal.aborted || (error instanceof Error && error.message === 'aborted')
          ? 'aborted'
          : error instanceof Error && error.message === 'invalid_state'
            ? 'invalid_state'
            : 'source_unavailable'
      // Preserve an interrupted run bound to another revision; a fresh audit needs a fresh ID.
      if (error instanceof Error && error.message === 'request_conflict') throw error
      assertWritable?.()
      store.finishSourceAudit(projectId, documentId, run.id, { error: code })
      throw new Error(code, { cause: error })
    }
  }
  return {
    projectId,
    planRevision: completed.planRevision,
    sources: completed.sources,
    checks: {
      support: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
    },
  }
}
