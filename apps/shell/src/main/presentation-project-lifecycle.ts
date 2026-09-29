import { PresentationLifecycleStore } from '@wiswork/project-store'
import type { PresentationRetentionPolicy } from '@wiswork/project-store'

export const presentationProjectLifecycleOperations = [
  'project_lifecycle_initialize',
  'project_lifecycle_read',
  'project_lifecycle_set_policy',
  'project_lifecycle_export_audit',
] as const
const errors = new Set([
  'invalid_request',
  'invalid_state',
  'output_too_large',
  'document_mismatch',
  'revision_conflict',
  'project_not_found',
  'project_deleting',
  'project_deleted',
  'busy',
])
function invalid(): never {
  throw Error('invalid_request')
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
    const descriptor = Object.getOwnPropertyDescriptor(value, key)!
    if (!descriptor.enumerable || !('value' in descriptor)) invalid()
    result[key] = descriptor.value
  }
  return result
}
/** Private paired-PC metadata service. Retention execution and deletion are separate transactions. */
export function createPresentationProjectLifecycleService(options: { userDataPath: string }) {
  const store = new PresentationLifecycleStore(options.userDataPath)
  return (body: unknown): unknown => {
    try {
      const request = record(body)
      const operation = request.operation
      if (
        typeof operation !== 'string' ||
        !(presentationProjectLifecycleOperations as readonly string[]).includes(operation)
      )
        invalid()
      const keys = [
        'operation',
        'projectId',
        'documentId',
        ...(operation === 'project_lifecycle_set_policy' ? ['expectedRevision', 'policy'] : []),
      ]
      if (
        Object.keys(request).length !== keys.length ||
        keys.some((key) => !Object.hasOwn(request, key))
      )
        invalid()
      if (
        typeof request.projectId !== 'string' ||
        !/^[A-Za-z0-9_-]{1,128}$/.test(request.projectId) ||
        typeof request.documentId !== 'string' ||
        !request.documentId.trim() ||
        request.documentId.length > 2048
      )
        invalid()
      const scope = { projectId: request.projectId, documentId: request.documentId }
      if (operation === 'project_lifecycle_initialize')
        return { lifecycle: store.initialize(scope) }
      if (operation === 'project_lifecycle_read') return { lifecycle: store.read(scope) ?? null }
      if (operation === 'project_lifecycle_export_audit') return { audit: store.exportAudit(scope) }
      if (
        typeof request.expectedRevision !== 'number' ||
        !Number.isSafeInteger(request.expectedRevision) ||
        request.expectedRevision < 0
      )
        invalid()
      const policy = record(request.policy)
      if (
        Object.keys(policy).length !== 2 ||
        !Object.hasOwn(policy, 'contentRetentionDays') ||
        !Object.hasOwn(policy, 'auditRetentionDays')
      )
        invalid()
      for (const value of Object.values(policy)) {
        if (
          value !== null &&
          (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 1 || value > 36500)
        )
          invalid()
      }
      return {
        lifecycle: store.setPolicy(
          scope,
          request.expectedRevision,
          policy as unknown as PresentationRetentionPolicy,
        ),
      }
    } catch (error) {
      let code = 'invalid_state'
      try {
        const message = error instanceof Error ? error.message : undefined
        if (typeof message === 'string' && errors.has(message)) code = message
      } catch {
        /* Unknown exception metadata never enters the response. */
      }
      return { error: code }
    }
  }
}
