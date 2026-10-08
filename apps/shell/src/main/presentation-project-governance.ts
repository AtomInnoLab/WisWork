import { resolve } from 'node:path'
import {
  PresentationStore,
  PresentationLifecycleStore,
  type PresentationRetentionPolicy,
} from '@wiswork/project-store'
import { createPresentationProjectDeletionService } from './presentation-project-deletion'
import { registerPresentationProjectWork } from './presentation-project-work'
export const presentationProjectGovernanceOperations = [
  'project_deletion_preview',
  'project_deletion_confirm',
  'project_deletion_resume',
  'project_lifecycle_initialize',
  'project_lifecycle_read',
  'project_lifecycle_set_policy',
  'project_lifecycle_export_audit',
] as const
const INPUT_BYTES = 32 * 1024,
  MODEL_BYTES = 2 * 1024 * 1024,
  OUTPUT_BYTES = MODEL_BYTES + 14
const errors = new Set([
  'invalid_request',
  'invalid_state',
  'aborted',
  'cancelled',
  'output_too_large',
  'document_mismatch',
  'revision_conflict',
  'project_not_found',
  'project_deleting',
  'project_deleted',
  'project_busy',
  'busy',
  'deletion_conflict',
  'confirmation_conflict',
  'deletion_incomplete',
  'presentation_inventory_invalid',
  'presentation_inventory_budget',
])
function invalid(): never {
  throw Error('invalid_request')
}
const check = (signal: AbortSignal) => {
  if (signal.aborted) throw Error('aborted')
}
function record(value: unknown): Record<string, unknown> {
  if (
    !value ||
    typeof value !== 'object' ||
    Array.isArray(value) ||
    Object.getPrototypeOf(value) !== Object.prototype
  )
    invalid()
  const result: Record<string, unknown> = {}
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== 'string' || ['__proto__', 'prototype', 'constructor'].includes(key))
      invalid()
    const d = Object.getOwnPropertyDescriptor(value, key)!
    if (!d.enumerable || !('value' in d)) invalid()
    result[key] = d.value
  }
  return result
}
function exact(value: Record<string, unknown>, keys: string[]) {
  if (Object.keys(value).length !== keys.length || keys.some((k) => !Object.hasOwn(value, k)))
    invalid()
}
function number(value: unknown, nullable = false): number | null {
  if (nullable && value === null) return null
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 0) invalid()
  return value
}
function identifier(value: unknown): string {
  if (typeof value !== 'string' || !/^[A-Za-z0-9_-]{1,128}$/.test(value)) invalid()
  return value
}
function parse(body: unknown) {
  const r = record(body),
    op = r.operation
  if (
    typeof op !== 'string' ||
    !(presentationProjectGovernanceOperations as readonly string[]).includes(op)
  )
    invalid()
  const extra =
    op === 'project_deletion_confirm'
      ? ['expectedRevision', 'confirmationToken', 'deletionId']
      : op === 'project_deletion_resume'
        ? ['expectedRevision', 'deletionId']
        : op === 'project_lifecycle_set_policy'
          ? ['expectedRevision', 'policy']
          : []
  exact(r, ['operation', 'documentId', 'projectId', ...extra])
  const projectId = identifier(r.projectId),
    documentId = r.documentId
  if (
    typeof documentId !== 'string' ||
    !documentId.trim() ||
    documentId.length > 4096 ||
    Array.from(documentId).some(
      (c) => c.charCodeAt(0) < 32 || (c.charCodeAt(0) >= 127 && c.charCodeAt(0) <= 159),
    )
  )
    invalid()
  if (op === 'project_deletion_confirm' || op === 'project_deletion_resume') {
    number(r.expectedRevision, op === 'project_deletion_confirm')
    identifier(r.deletionId)
    if (
      op === 'project_deletion_confirm' &&
      (typeof r.confirmationToken !== 'string' || !/^[a-f0-9]{64}$/.test(r.confirmationToken))
    )
      invalid()
  }
  if (op === 'project_lifecycle_set_policy') {
    number(r.expectedRevision)
    const p = record(r.policy)
    exact(p, ['contentRetentionDays', 'auditRetentionDays'])
    for (const v of Object.values(p))
      if (v !== null && (typeof v !== 'number' || !Number.isSafeInteger(v) || v < 1 || v > 36500))
        invalid()
    r.policy = p
  }
  if (Buffer.byteLength(JSON.stringify(r)) > INPUT_BYTES) invalid()
  return { request: structuredClone(r), scope: Object.freeze({ projectId, documentId }) }
}
function bounded<T>(value: T): T {
  const payload = Object.values(value as Record<string, unknown>)[0]
  // The largest exact wrapper is {"lifecycle":}; the inner model keeps its original 2 MiB bound.
  if (
    Buffer.byteLength(JSON.stringify(payload)) > MODEL_BYTES ||
    Buffer.byteLength(JSON.stringify(value)) > OUTPUT_BYTES
  )
    throw Error('output_too_large')
  return value
}
/** Internal governance adapter only; publication of these operations requires a separate entrypoint gate. */
export function createPresentationProjectGovernanceService(optionsValue: {
  userDataPath: string
  acquireProjectLock: (projectId: string) => Promise<() => void>
}) {
  const options = { ...optionsValue, userDataPath: resolve(optionsValue.userDataPath) },
    life = new PresentationLifecycleStore(options.userDataPath),
    store = new PresentationStore(options.userDataPath),
    deletion = createPresentationProjectDeletionService(options)
  return async (body: unknown, signal: AbortSignal): Promise<unknown> => {
    try {
      const { request, scope } = parse(body)
      check(signal)
      const op = request.operation
      if (op === 'project_deletion_preview')
        return bounded({ preview: await deletion.preview(scope, signal) })
      if (op === 'project_deletion_confirm')
        return bounded({
          deletion: await deletion.confirm(
            {
              scope,
              expectedRevision: request.expectedRevision as number | null,
              confirmationToken: request.confirmationToken as string,
              deletionId: request.deletionId as string,
            },
            signal,
          ),
        })
      if (op === 'project_deletion_resume')
        return bounded({
          deletion: await deletion.resume(
            {
              scope,
              expectedRevision: request.expectedRevision as number,
              deletionId: request.deletionId as string,
            },
            signal,
          ),
        })
      // Control reads are independent of production body existence and never grant write authorization.
      const initial = life.readControl(scope)
      if (op === 'project_lifecycle_read') return bounded({ lifecycle: initial ?? null })
      if (op === 'project_lifecycle_export_audit') {
        if (!initial) throw Error('project_not_found')
        return bounded({
          audit: {
            version: 1,
            anonymousProjectId: initial.anonymousProjectId,
            auditRetentionDays: initial.policy.auditRetentionDays,
            events: structuredClone(initial.audit),
          },
        })
      }
      if (initial && initial.state !== 'active') throw Error('project_' + initial.state)
      if (op === 'project_lifecycle_initialize') {
        if (!store.projectScope(scope.projectId, scope.documentId)) throw Error('project_not_found')
      } else if (!initial) throw Error('project_not_found')
      const work = registerPresentationProjectWork({
        scope: { root: options.userDataPath, ...scope },
        signal,
      })
      signal = work.signal
      const fixed = () => {
        check(signal)
        const current = life.readControl(scope)
        if (current?.revision !== initial?.revision) throw Error('revision_conflict')
        if (current && current.state !== 'active') throw Error('project_' + current.state)
        check(signal)
      }
      try {
        fixed()
        const release = await options.acquireProjectLock(scope.projectId)
        try {
          fixed()
          if (op === 'project_lifecycle_initialize') {
            if (!store.projectScope(scope.projectId, scope.documentId))
              throw Error('project_not_found')
            fixed()
            return bounded({ lifecycle: life.initialize(scope) })
          }
          // The caller's numeric CAS and the admission epoch must both hold; neither is refreshed after waiting.
          return bounded({
            lifecycle: life.setPolicy(
              scope,
              request.expectedRevision as number,
              request.policy as unknown as PresentationRetentionPolicy,
            ),
          })
        } finally {
          release()
        }
      } finally {
        work.finish()
      }
    } catch (error) {
      let code = 'invalid_state'
      try {
        const message = error instanceof Error ? error.message : undefined
        if (typeof message === 'string' && errors.has(message)) code = message
      } catch {
        /* Do not return private exception metadata. */
      }
      return { error: code }
    }
  }
}
