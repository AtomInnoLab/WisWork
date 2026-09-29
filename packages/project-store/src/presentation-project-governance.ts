/** Browser-safe public governance contracts. No filesystem or Node dependencies. */
export interface PresentationRetentionPolicy {
  contentRetentionDays: number | null
  auditRetentionDays: number | null
}
export const DEFAULT_PRESENTATION_RETENTION_POLICY: Readonly<PresentationRetentionPolicy> =
  Object.freeze({ contentRetentionDays: null, auditRetentionDays: null })
export interface PresentationLifecycleScope {
  projectId: string
  documentId: string
}
export const PRESENTATION_LIFECYCLE_RESOURCE_KINDS = [
  'project',
  'research',
  'delivery_bundles',
  'page_backups',
  'preferences',
  'comments',
  'manual_observations',
  'teams',
  'attachments',
  'acquisition_history',
  'existing_page_backups',
  'master_backups',
  'package_backups',
  'brand_kits',
] as const
export type PresentationLifecycleResourceKind =
  (typeof PRESENTATION_LIFECYCLE_RESOURCE_KINDS)[number]
export type PresentationLifecycleResourceStatus =
  'pending' | 'removed' | 'reference_removed' | 'retained' | 'failed'
export type PresentationLifecycleResultCode =
  'ownership_unproven' | 'shared_resource' | 'io_failed' | 'resource_busy' | 'not_found'
export interface PresentationLifecycleResource {
  resourceId: string
  kind: PresentationLifecycleResourceKind
  ownership: 'project_exclusive' | 'shared_reference' | 'unproven'
  status: PresentationLifecycleResourceStatus
  code?: PresentationLifecycleResultCode
}
export interface PresentationDeletionIntent {
  deletionId: string
  reason: 'user' | 'retention'
  resources: Omit<PresentationLifecycleResource, 'status' | 'code'>[]
}
export interface PresentationDeletionResult {
  deletionId: string
  resourceId: string
  status: Exclude<PresentationLifecycleResourceStatus, 'pending'>
  code?: PresentationLifecycleResultCode
}
export interface PresentationLifecycleAuditEvent {
  sequence: number
  at: string
  action:
    'created' | 'policy_updated' | 'deletion_started' | 'resource_result' | 'deletion_finished'
  result: 'accepted' | 'partial' | 'complete'
  counts: {
    pending: number
    removed: number
    referenceRemoved: number
    retained: number
    failed: number
  }
}
export interface PresentationLifecycleRecord extends PresentationLifecycleScope {
  version: 1
  revision: number
  state: 'active' | 'deleting' | 'deleted'
  anonymousProjectId: string
  createdAt: string
  updatedAt: string
  policy: PresentationRetentionPolicy
  deletion?: {
    deletionId: string
    reason: 'user' | 'retention'
    resources: PresentationLifecycleResource[]
  }
  audit: PresentationLifecycleAuditEvent[]
}
export const MAX_PRESENTATION_LIFECYCLE_BYTES = 2 * 1024 * 1024
const MAX_RESOURCES = 4096,
  MAX_AUDIT_EVENTS = 8192
function invalid(_code = 'invalid_request'): never {
  throw Error('presentation_governance_response_invalid')
}
const id = (value: unknown) => typeof value === 'string' && /^[A-Za-z0-9_-]{1,128}$/.test(value)
const plain = (value: unknown): value is Record<string, unknown> =>
  Boolean(
    value &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    Object.getPrototypeOf(value) === Object.prototype,
  )
function owned<T>(value: T): T {
  const visit = (v: unknown, depth: number) => {
    if (depth > 12) invalid()
    if (
      v === null ||
      typeof v === 'string' ||
      typeof v === 'boolean' ||
      (typeof v === 'number' && Number.isFinite(v))
    )
      return
    if (!Array.isArray(v) && !plain(v)) invalid()
    for (const key of Reflect.ownKeys(v as object)) {
      if (typeof key !== 'string' || ['__proto__', 'constructor', 'prototype'].includes(key))
        invalid()
      const d = Object.getOwnPropertyDescriptor(v, key)!
      if (!('value' in d) || (!d.enumerable && !(Array.isArray(v) && key === 'length'))) invalid()
      if (key !== 'length' || !Array.isArray(v)) visit(d.value, depth + 1)
    }
    if (Array.isArray(v) && Object.keys(v).length !== v.length) invalid()
  }
  visit(value, 0)
  if (new TextEncoder().encode(JSON.stringify(value)).byteLength > MAX_PRESENTATION_LIFECYCLE_BYTES)
    invalid('output_too_large')
  return structuredClone(value)
}
const exact = (value: unknown, keys: string[]) =>
  plain(value) && Object.keys(value).sort().join(',') === keys.sort().join(',')
function scope(value: PresentationLifecycleScope) {
  if (
    !exact(value, ['projectId', 'documentId']) ||
    !id(value.projectId) ||
    typeof value.documentId !== 'string' ||
    !value.documentId.trim() ||
    value.documentId.length > 4096 ||
    Array.from(value.documentId).some(
      (char) => char.charCodeAt(0) < 32 || (char.charCodeAt(0) >= 127 && char.charCodeAt(0) <= 159),
    )
  )
    invalid()
}
function policy(value: PresentationRetentionPolicy) {
  if (
    !exact(value, ['contentRetentionDays', 'auditRetentionDays']) ||
    Object.values(value).some((v) => v !== null && (!Number.isSafeInteger(v) || v < 1 || v > 36500))
  )
    invalid()
}
const codes = ['ownership_unproven', 'shared_resource', 'io_failed', 'resource_busy', 'not_found']
function resource(value: PresentationLifecycleResource, stored: boolean) {
  if (
    !exact(
      value,
      stored
        ? [
            'resourceId',
            'kind',
            'ownership',
            'status',
            ...(value.code === undefined ? [] : ['code']),
          ]
        : ['resourceId', 'kind', 'ownership'],
    ) ||
    !id(value.resourceId) ||
    !PRESENTATION_LIFECYCLE_RESOURCE_KINDS.includes(value.kind) ||
    !['project_exclusive', 'shared_reference', 'unproven'].includes(value.ownership)
  )
    invalid()
  if (
    stored &&
    (!['pending', 'removed', 'reference_removed', 'retained', 'failed'].includes(value.status) ||
      (value.code !== undefined && !codes.includes(value.code)) ||
      (value.status === 'pending' && value.code !== undefined) ||
      (value.status === 'removed' && value.ownership !== 'project_exclusive') ||
      (value.status === 'reference_removed' && value.ownership !== 'shared_reference'))
  )
    invalid()
}
function date(value: unknown): value is string {
  return (
    typeof value === 'string' &&
    /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(value) &&
    Number.isFinite(Date.parse(value)) &&
    new Date(value).toISOString() === value
  )
}
function counts(record: PresentationLifecycleRecord): PresentationLifecycleAuditEvent['counts'] {
  const all = record.deletion?.resources ?? []
  return {
    pending: all.filter((r) => r.status === 'pending').length,
    removed: all.filter((r) => r.status === 'removed').length,
    referenceRemoved: all.filter((r) => r.status === 'reference_removed').length,
    retained: all.filter((r) => r.status === 'retained').length,
    failed: all.filter((r) => r.status === 'failed').length,
  }
}
export function parsePresentationGovernanceLifecycle(value: unknown): PresentationLifecycleRecord {
  try {
    const r = owned(value) as PresentationLifecycleRecord
    if (
      !exact(r, [
        'version',
        'projectId',
        'documentId',
        'revision',
        'state',
        'anonymousProjectId',
        'createdAt',
        'updatedAt',
        'policy',
        'audit',
        ...(r?.deletion === undefined ? [] : ['deletion']),
      ]) ||
      r.version !== 1 ||
      !Number.isSafeInteger(r.revision) ||
      r.revision < 0 ||
      !['active', 'deleting', 'deleted'].includes(r.state) ||
      typeof r.anonymousProjectId !== 'string' ||
      !/^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/.test(
        r.anonymousProjectId,
      ) ||
      !date(r.createdAt) ||
      !date(r.updatedAt) ||
      r.updatedAt < r.createdAt
    )
      invalid()
    scope({ projectId: r.projectId, documentId: r.documentId })
    policy(r.policy)
    if ((r.state === 'active') !== (r.deletion === undefined)) invalid()
    if (r.deletion) {
      const d = r.deletion
      if (
        !exact(d, ['deletionId', 'reason', 'resources']) ||
        !id(d.deletionId) ||
        !['user', 'retention'].includes(d.reason) ||
        !Array.isArray(d.resources) ||
        d.resources.length < 1 ||
        d.resources.length > MAX_RESOURCES
      )
        invalid()
      d.resources.forEach((v) => resource(v, true))
      if (new Set(d.resources.map((v) => v.resourceId)).size !== d.resources.length) invalid()
      if (
        r.state === 'deleted' &&
        d.resources.some((v) => !['removed', 'reference_removed'].includes(v.status))
      )
        invalid()
    }
    if (
      !Array.isArray(r.audit) ||
      r.audit.length !== r.revision + 1 ||
      r.audit.length > MAX_AUDIT_EVENTS
    )
      invalid()
    let last = r.createdAt
    let auditState: PresentationLifecycleRecord['state'] = 'active'
    let previousCounts: PresentationLifecycleAuditEvent['counts'] = {
      pending: 0,
      removed: 0,
      referenceRemoved: 0,
      retained: 0,
      failed: 0,
    }
    const sameCounts = (
      a: PresentationLifecycleAuditEvent['counts'],
      b: PresentationLifecycleAuditEvent['counts'],
    ) => (Object.keys(a) as (keyof typeof a)[]).every((key) => a[key] === b[key])
    for (const [i, e] of r.audit.entries()) {
      if (
        !exact(e, ['sequence', 'at', 'action', 'result', 'counts']) ||
        e.sequence !== i ||
        !date(e.at) ||
        e.at < last ||
        e.at > r.updatedAt ||
        ![
          'created',
          'policy_updated',
          'deletion_started',
          'resource_result',
          'deletion_finished',
        ].includes(e.action) ||
        !['accepted', 'partial', 'complete'].includes(e.result) ||
        !exact(e.counts, ['pending', 'removed', 'referenceRemoved', 'retained', 'failed']) ||
        Object.values(e.counts).some((n) => !Number.isSafeInteger(n) || n < 0 || n > MAX_RESOURCES)
      )
        invalid()
      if (i === 0 && (e.action !== 'created' || e.at !== r.createdAt || e.result !== 'accepted'))
        invalid()
      if (i === 0 && !sameCounts(e.counts, previousCounts)) invalid()
      if (i > 0) {
        if (e.action === 'policy_updated') {
          if (
            auditState !== 'active' ||
            e.result !== 'accepted' ||
            !sameCounts(e.counts, previousCounts)
          )
            invalid()
        } else if (e.action === 'deletion_started') {
          if (
            auditState !== 'active' ||
            e.result !== 'accepted' ||
            !r.deletion ||
            e.counts.pending !== r.deletion.resources.length ||
            e.counts.removed ||
            e.counts.referenceRemoved ||
            e.counts.retained ||
            e.counts.failed
          )
            invalid()
          auditState = 'deleting'
        } else if (e.action === 'resource_result') {
          const deltas = (Object.keys(e.counts) as (keyof typeof e.counts)[]).map(
            (key) => e.counts[key] - previousCounts[key],
          )
          if (
            auditState !== 'deleting' ||
            deltas.reduce((a, b) => a + b, 0) !== 0 ||
            deltas.reduce((a, b) => a + Math.abs(b), 0) > 2 ||
            e.counts.removed < previousCounts.removed ||
            e.counts.referenceRemoved < previousCounts.referenceRemoved ||
            e.result !==
              (e.counts.pending || e.counts.retained || e.counts.failed ? 'partial' : 'complete')
          )
            invalid()
        } else if (e.action === 'deletion_finished') {
          if (
            auditState !== 'deleting' ||
            previousCounts.pending ||
            previousCounts.retained ||
            previousCounts.failed ||
            !sameCounts(e.counts, previousCounts) ||
            e.result !== 'complete'
          )
            invalid()
          auditState = 'deleted'
        } else invalid()
      }
      previousCounts = e.counts
      last = e.at
    }
    if (
      last !== r.updatedAt ||
      auditState !== r.state ||
      !sameCounts(r.audit.at(-1)!.counts, counts(r))
    )
      invalid()
    return r
  } catch {
    invalid('invalid_state')
  }
}

export const MAX_PRESENTATION_GOVERNANCE_BYTES = MAX_PRESENTATION_LIFECYCLE_BYTES
export const MAX_PRESENTATION_GOVERNANCE_RESPONSE_BYTES = MAX_PRESENTATION_GOVERNANCE_BYTES + 14
export const MAX_PRESENTATION_DELETION_ATTEMPT_BYTES = 32 * 1024
const inventoryKinds = [
  ...PRESENTATION_LIFECYCLE_RESOURCE_KINDS,
  'existing_page_releases',
  'brand_reference',
  'unbound_page_staging',
  'unbound_existing_staging',
] as const
export interface PresentationGovernanceResource {
  kind: (typeof inventoryKinds)[number]
  ownership: 'project_exclusive' | 'document_shared' | 'global_shared' | 'unproven'
  resourceId: string
  fileCount: number
  bytes: number
}
export interface PresentationGovernancePreview {
  expectedRevision: number | null
  confirmationToken: string
  resources: (PresentationGovernanceResource & { disposition: 'candidate' | 'retained' })[]
  governanceRetained: true
}
export interface PresentationGovernanceDeletionReport {
  state: 'partial' | 'deleted'
  revision: number
  deletionId: string
  projectContentRetained: boolean
  counts: { removed: number; pending: number; failed: number; retained: number }
  retained: PresentationGovernanceResource[]
}
export interface PresentationProjectDeletionAttempt {
  version: 1
  scope: PresentationLifecycleScope
  expectedRevision: number | null
  confirmationToken: string
  deletionId: string
}
export interface PresentationGovernanceAudit {
  version: 1
  anonymousProjectId: string
  auditRetentionDays: number | null
  events: PresentationLifecycleAuditEvent[]
}
const nonnegative = (v: unknown): v is number => Number.isSafeInteger(v) && (v as number) >= 0
const revision = (v: unknown) => v === null || nonnegative(v)
const token = (v: unknown) => typeof v === 'string' && /^[a-f0-9]{64}$/.test(v)
function inventoryResource(v: PresentationGovernanceResource, preview: boolean) {
  if (
    !exact(v, [
      'kind',
      'ownership',
      'resourceId',
      'fileCount',
      'bytes',
      ...(preview ? ['disposition'] : []),
    ]) ||
    !inventoryKinds.includes(v.kind) ||
    !['project_exclusive', 'document_shared', 'global_shared', 'unproven'].includes(v.ownership) ||
    !id(v.resourceId) ||
    !nonnegative(v.fileCount) ||
    v.fileCount > 32768 ||
    !nonnegative(v.bytes) ||
    v.bytes > 32 * 1024 ** 3
  )
    invalid()
  if (
    preview &&
    (v as PresentationGovernancePreview['resources'][number]).disposition !==
      (v.ownership === 'project_exclusive' ? 'candidate' : 'retained')
  )
    invalid()
}
function inventory(v: PresentationGovernanceResource[], preview: boolean) {
  if (!Array.isArray(v) || v.length > 4096) invalid()
  v.forEach((r) => inventoryResource(r, preview))
  if (new Set(v.map((r) => r.resourceId)).size !== v.length) invalid()
}
function parsePresentationGovernancePreviewInner(value: unknown): PresentationGovernancePreview {
  const v = owned(value) as PresentationGovernancePreview
  if (
    !exact(v, ['expectedRevision', 'confirmationToken', 'resources', 'governanceRetained']) ||
    !revision(v.expectedRevision) ||
    !token(v.confirmationToken) ||
    v.governanceRetained !== true
  )
    invalid()
  inventory(v.resources, true)
  return v
}
function parsePresentationGovernanceDeletionReportInner(
  value: unknown,
): PresentationGovernanceDeletionReport {
  const v = owned(value) as PresentationGovernanceDeletionReport
  if (
    !exact(v, [
      'state',
      'revision',
      'deletionId',
      'projectContentRetained',
      'counts',
      'retained',
    ]) ||
    !['partial', 'deleted'].includes(v.state) ||
    !nonnegative(v.revision) ||
    !id(v.deletionId) ||
    !exact(v.counts, ['removed', 'pending', 'failed', 'retained']) ||
    Object.values(v.counts).some((n) => !nonnegative(n) || n > 4096)
  )
    invalid()
  inventory(v.retained, false)
  if (
    v.counts.retained !== v.retained.length ||
    v.projectContentRetained !== v.retained.some((r) => r.kind === 'project') ||
    (v.state === 'deleted' && (v.counts.pending || v.counts.failed || v.counts.retained))
  )
    invalid()
  return v
}
function parsePresentationProjectDeletionAttemptInner(
  value: unknown,
): PresentationProjectDeletionAttempt {
  const v = owned(value) as PresentationProjectDeletionAttempt
  if (
    new TextEncoder().encode(JSON.stringify(v)).byteLength >
      MAX_PRESENTATION_DELETION_ATTEMPT_BYTES ||
    !exact(v, ['version', 'scope', 'expectedRevision', 'confirmationToken', 'deletionId']) ||
    v.version !== 1 ||
    !revision(v.expectedRevision) ||
    !token(v.confirmationToken) ||
    !id(v.deletionId)
  )
    invalid()
  scope(v.scope)
  return v
}
function parsePresentationGovernanceAuditInner(value: unknown): PresentationGovernanceAudit {
  const v = owned(value) as PresentationGovernanceAudit
  if (
    !exact(v, ['version', 'anonymousProjectId', 'auditRetentionDays', 'events']) ||
    v.version !== 1 ||
    !Array.isArray(v.events) ||
    !v.events.length
  )
    invalid()
  policy({ contentRetentionDays: null, auditRetentionDays: v.auditRetentionDays })
  const last = v.events.at(-1)!,
    start = v.events.find((e) => e.action === 'deletion_started'),
    finished = last.action === 'deletion_finished'
  const resources: PresentationLifecycleResource[] = []
  if (start) {
    for (const [key, status] of [
      ['pending', 'pending'],
      ['removed', 'removed'],
      ['referenceRemoved', 'reference_removed'],
      ['retained', 'retained'],
      ['failed', 'failed'],
    ] as const) {
      if (!nonnegative(last.counts?.[key]) || last.counts[key] > 4096) invalid()
      for (let i = 0; i < last.counts[key]; i++)
        resources.push({
          resourceId: 'r' + resources.length,
          kind: 'project',
          ownership: status === 'reference_removed' ? 'shared_reference' : 'project_exclusive',
          status,
        })
    }
  }
  parsePresentationGovernanceLifecycle({
    version: 1,
    projectId: 'audit',
    documentId: 'audit',
    revision: v.events.length - 1,
    state: start ? (finished ? 'deleted' : 'deleting') : 'active',
    anonymousProjectId: v.anonymousProjectId,
    createdAt: v.events[0]!.at,
    updatedAt: last.at,
    policy: { contentRetentionDays: null, auditRetentionDays: v.auditRetentionDays },
    audit: v.events,
    ...(start ? { deletion: { deletionId: 'audit', reason: 'user', resources } } : {}),
  })
  return v
}

export function parsePresentationGovernancePreview(value: unknown): PresentationGovernancePreview {
  try {
    return parsePresentationGovernancePreviewInner(value)
  } catch {
    invalid()
  }
}

export function parsePresentationGovernanceDeletionReport(
  value: unknown,
): PresentationGovernanceDeletionReport {
  try {
    return parsePresentationGovernanceDeletionReportInner(value)
  } catch {
    invalid()
  }
}

export function parsePresentationProjectDeletionAttempt(
  value: unknown,
): PresentationProjectDeletionAttempt {
  try {
    return parsePresentationProjectDeletionAttemptInner(value)
  } catch {
    invalid()
  }
}

export function parsePresentationGovernanceAudit(value: unknown): PresentationGovernanceAudit {
  try {
    return parsePresentationGovernanceAuditInner(value)
  } catch {
    invalid()
  }
}
export function parsePresentationGovernancePolicy(value: unknown): PresentationRetentionPolicy {
  try {
    const cloned = owned(value) as PresentationRetentionPolicy
    policy(cloned)
    return cloned
  } catch {
    invalid()
  }
}
