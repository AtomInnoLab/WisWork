import { createHash, randomUUID } from 'node:crypto'
import {
  closeSync,
  constants,
  fstatSync,
  fsyncSync,
  lstatSync,
  mkdirSync,
  openSync,
  readSync,
  readdirSync,
  renameSync,
  unlinkSync,
  writeFileSync,
  type BigIntStats,
} from 'node:fs'
import { dirname, join, parse, resolve } from 'node:path'

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
const busy = new Set<string>()
function invalid(code = 'invalid_request'): never {
  throw Error(code)
}
const hash = (value: string) => createHash('sha256').update(value).digest('hex')
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
  if (Buffer.byteLength(JSON.stringify(value)) > MAX_PRESENTATION_LIFECYCLE_BYTES)
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
    value.documentId.length > 2048 ||
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
export function parsePresentationLifecycle(value: unknown): PresentationLifecycleRecord {
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
type Identity = { path: string; dev: bigint; ino: bigint }
const stat = (path: string) => lstatSync(path, { bigint: true, throwIfNoEntry: false })
const same = (a: BigIntStats, b: { dev: bigint; ino: bigint }) => a.dev === b.dev && a.ino === b.ino
function ancestors(path: string): Identity[] {
  const paths: string[] = []
  for (let p = path; ; p = dirname(p)) {
    paths.unshift(p)
    if (p === parse(p).root) break
  }
  return paths.map((path) => {
    const s = stat(path)
    if (!s || !s.isDirectory() || s.isSymbolicLink()) invalid('invalid_state')
    return { path, dev: s.dev, ino: s.ino }
  })
}
function guard(parents: Identity[]) {
  for (const p of parents) {
    const s = stat(p.path)
    if (!s || !s.isDirectory() || s.isSymbolicLink() || !same(s, p)) invalid('invalid_state')
  }
}
function readJson(path: string, parents: Identity[]): unknown | undefined {
  guard(parents)
  const before = stat(path)
  if (!before) return undefined
  if (
    !before.isFile() ||
    before.isSymbolicLink() ||
    before.size > BigInt(MAX_PRESENTATION_LIFECYCLE_BYTES)
  )
    invalid('invalid_state')
  let fd: number | undefined
  try {
    fd = openSync(path, constants.O_RDONLY | constants.O_NOFOLLOW)
    guard(parents)
    const held = fstatSync(fd, { bigint: true })
    if (
      !held.isFile() ||
      !same(held, before) ||
      held.size > BigInt(MAX_PRESENTATION_LIFECYCLE_BYTES)
    )
      invalid('invalid_state')
    const bytes = Buffer.alloc(MAX_PRESENTATION_LIFECYCLE_BYTES + 1)
    let length = 0,
      n = 0
    do {
      n = readSync(fd, bytes, length, bytes.length - length, length)
      length += n
    } while (n && length < bytes.length)
    guard(parents)
    if (
      length > MAX_PRESENTATION_LIFECYCLE_BYTES ||
      !same(stat(path) ?? held, before) ||
      !stat(path)
    )
      invalid('invalid_state')
    return JSON.parse(bytes.subarray(0, length).toString('utf8'))
  } catch {
    invalid('invalid_state')
  } finally {
    if (fd !== undefined) closeSync(fd)
  }
}
function writeJson(path: string, value: unknown, parents: Identity[]) {
  const text = JSON.stringify(value)
  if (Buffer.byteLength(text) > MAX_PRESENTATION_LIFECYCLE_BYTES) invalid('output_too_large')
  guard(parents)
  const previous = stat(path)
  if (previous && (!previous.isFile() || previous.isSymbolicLink())) invalid('invalid_state')
  const temporary = path + '.' + randomUUID() + '.tmp'
  let fd: number | undefined, created: BigIntStats | undefined
  try {
    fd = openSync(
      temporary,
      constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW,
      0o600,
    )
    guard(parents)
    created = fstatSync(fd, { bigint: true })
    if (!created.isFile() || !stat(temporary) || !same(stat(temporary)!, created))
      invalid('invalid_state')
    writeFileSync(fd, text, 'utf8')
    fsyncSync(fd)
    guard(parents)
    const current = stat(path)
    if (previous ? !current || !same(current, previous) : current !== undefined)
      invalid('revision_conflict')
    if (!stat(temporary) || !same(stat(temporary)!, created)) invalid('invalid_state')
    renameSync(temporary, path)
    guard(parents)
    if (!stat(path) || !same(stat(path)!, created)) invalid('invalid_state')
    const directory = openSync(dirname(path), constants.O_RDONLY | constants.O_NOFOLLOW)
    try {
      guard(parents)
      if (!same(fstatSync(directory, { bigint: true }), parents.at(-1)!)) invalid('invalid_state')
      fsyncSync(directory)
    } finally {
      closeSync(directory)
    }
  } finally {
    if (fd !== undefined) closeSync(fd)
    if (created) {
      guard(parents)
      const temp = stat(temporary)
      if (temp && same(temp, created)) unlinkSync(temporary)
    }
  }
}
/** Metadata only. Policy scheduling, resource removal and all-entry write fences are separate integrations. */
export class PresentationLifecycleStore {
  private readonly root: string
  constructor(userDataPath: string) {
    this.root = resolve(userDataPath)
  }
  private checkOwnership(s: PresentationLifecycleScope): void {
    const paths = [
      join(this.root, 'projects'),
      join(this.root, 'projects', 'presentations'),
      join(this.root, 'projects', 'presentations', hash(s.projectId)),
    ]
    for (const path of paths) {
      const found = stat(path)
      if (!found) return
      if (!found.isDirectory() || found.isSymbolicLink()) invalid('invalid_state')
    }
    const directory = paths.at(-1)!,
      parents = ancestors(directory)
    const binding = readJson(join(directory, 'project.json'), parents) as
      Record<string, unknown> | undefined
    guard(parents)
    if (binding === undefined) {
      if (readdirSync(directory).length) invalid('invalid_state')
      guard(parents)
      return
    }
    if (
      !exact(binding, ['version', 'projectId', 'documentId']) ||
      binding.version !== 1 ||
      binding.projectId !== s.projectId ||
      typeof binding.documentId !== 'string'
    )
      invalid('invalid_state')
    if (binding.documentId !== s.documentId) invalid('document_mismatch')
  }
  private directory(s: PresentationLifecycleScope, create = false): string | undefined {
    scope(s)
    ancestors(this.root)
    this.checkOwnership(s)
    const paths = [
      join(this.root, 'presentation-project-lifecycles'),
      join(this.root, 'presentation-project-lifecycles', hash(s.projectId)),
    ]
    for (const path of paths) {
      const old = stat(path)
      if (!old) {
        if (!create) return undefined
        const parent = ancestors(dirname(path))
        guard(parent)
        mkdirSync(path, { mode: 0o700 })
        guard(parent)
      }
      const found = stat(path)
      if (!found || !found.isDirectory() || found.isSymbolicLink()) invalid('invalid_state')
    }
    return paths.at(-1)!
  }
  read(input: PresentationLifecycleScope): PresentationLifecycleRecord | undefined {
    const s = owned(input)
    scope(s)
    const directory = this.directory(s)
    if (!directory) return undefined
    const raw = readJson(join(directory, 'lifecycle.json'), ancestors(directory))
    if (raw === undefined) return undefined
    const record = parsePresentationLifecycle(raw)
    if (record.projectId !== s.projectId) invalid('invalid_state')
    if (record.documentId !== s.documentId) invalid('document_mismatch')
    return record
  }
  private transaction<T>(
    input: PresentationLifecycleScope,
    work: (scope: PresentationLifecycleScope) => T,
  ): T {
    const s = owned(input)
    scope(s)
    const key = this.root + '\0' + s.projectId
    if (busy.has(key)) invalid('project_busy')
    busy.add(key)
    try {
      return work(s)
    } finally {
      busy.delete(key)
    }
  }
  private persist(s: PresentationLifecycleScope, r: PresentationLifecycleRecord) {
    const directory = this.directory(s, true)!
    writeJson(
      join(directory, 'lifecycle.json'),
      parsePresentationLifecycle(r),
      ancestors(directory),
    )
    return structuredClone(r)
  }
  private current(s: PresentationLifecycleScope, revision: number) {
    if (!Number.isSafeInteger(revision) || revision < 0) invalid()
    const r = this.read(s)
    if (!r) invalid('project_not_found')
    if (r.revision !== revision) invalid('revision_conflict')
    return r
  }
  private append(
    r: PresentationLifecycleRecord,
    action: PresentationLifecycleAuditEvent['action'],
    result: PresentationLifecycleAuditEvent['result'],
  ) {
    if (r.audit.length >= MAX_AUDIT_EVENTS) invalid('output_too_large')
    r.revision++
    r.updatedAt = new Date(Math.max(Date.now(), Date.parse(r.updatedAt))).toISOString()
    r.audit.push({ sequence: r.revision, at: r.updatedAt, action, result, counts: counts(r) })
  }
  initialize(
    input: PresentationLifecycleScope,
    inputPolicy: PresentationRetentionPolicy = DEFAULT_PRESENTATION_RETENTION_POLICY,
  ) {
    const p = owned(inputPolicy)
    policy(p)
    return this.transaction(input, (s) => {
      const existing = this.read(s)
      if (existing) return existing
      const now = new Date().toISOString(),
        r: PresentationLifecycleRecord = {
          version: 1,
          ...s,
          revision: 0,
          state: 'active',
          anonymousProjectId: randomUUID(),
          createdAt: now,
          updatedAt: now,
          policy: p,
          audit: [
            {
              sequence: 0,
              at: now,
              action: 'created',
              result: 'accepted',
              counts: { pending: 0, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 },
            },
          ],
        }
      return this.persist(s, r)
    })
  }
  assertActive(input: PresentationLifecycleScope, revision?: number) {
    if (revision !== undefined && (!Number.isSafeInteger(revision) || revision < 0)) invalid()
    const r = this.read(input)
    if (!r) invalid('project_not_found')
    if (revision !== undefined && r.revision !== revision) invalid('revision_conflict')
    if (r.state !== 'active') invalid('project_' + r.state)
    return r
  }
  setPolicy(
    input: PresentationLifecycleScope,
    revision: number,
    inputPolicy: PresentationRetentionPolicy,
  ) {
    const p = owned(inputPolicy)
    policy(p)
    return this.transaction(input, (s) => {
      const r = this.current(s, revision)
      this.assertActive(s)
      r.policy = p
      this.append(r, 'policy_updated', 'accepted')
      // Keep four receipts and 4 KiB for one maximal-ID resource/intent plus bounded audit events.
      if (
        r.audit.length + 4 > MAX_AUDIT_EVENTS ||
        Buffer.byteLength(JSON.stringify(r)) + 4096 > MAX_PRESENTATION_LIFECYCLE_BYTES
      )
        invalid('output_too_large')
      return this.persist(s, r)
    })
  }
  beginDeletion(
    input: PresentationLifecycleScope,
    revision: number,
    inputIntent: PresentationDeletionIntent,
  ) {
    const intent = owned(inputIntent)
    if (
      !exact(intent, ['deletionId', 'reason', 'resources']) ||
      !id(intent.deletionId) ||
      !['user', 'retention'].includes(intent.reason) ||
      !Array.isArray(intent.resources) ||
      intent.resources.length < 1 ||
      intent.resources.length > MAX_RESOURCES
    )
      invalid()
    intent.resources.forEach((r) => resource(r as PresentationLifecycleResource, false))
    if (new Set(intent.resources.map((r) => r.resourceId)).size !== intent.resources.length)
      invalid()
    return this.transaction(input, (s) => {
      const r = this.current(s, revision)
      this.assertActive(s)
      if (r.audit.length + 2 * intent.resources.length + 2 > MAX_AUDIT_EVENTS)
        invalid('output_too_large')
      r.state = 'deleting'
      r.deletion = {
        ...intent,
        resources: intent.resources.map((v) => ({ ...v, status: 'pending' })),
      }
      this.append(r, 'deletion_started', 'accepted')
      // Reserve an initial failure, successful retry and final checkpoint before execution.
      this.reserveDeletion(r, 2 * intent.resources.length + 1)
      return this.persist(s, r)
    })
  }
  private reserveDeletion(r: PresentationLifecycleRecord, futureEvents: number) {
    if (r.audit.length + futureEvents > MAX_AUDIT_EVENTS) invalid('output_too_large')
    const worst = structuredClone(r)
    worst.deletion!.resources = worst.deletion!.resources.map((v) => ({
      ...v,
      status: v.ownership === 'shared_reference' ? 'reference_removed' : 'retained',
      code: 'ownership_unproven',
    }))
    const futureEvent = {
      sequence: MAX_AUDIT_EVENTS - 1,
      at: r.updatedAt,
      action: 'deletion_finished',
      result: 'accepted',
      counts: {
        pending: MAX_RESOURCES,
        removed: MAX_RESOURCES,
        referenceRemoved: MAX_RESOURCES,
        retained: MAX_RESOURCES,
        failed: MAX_RESOURCES,
      },
    }
    const reserved =
      Buffer.byteLength(JSON.stringify(worst)) +
      futureEvents * (Buffer.byteLength(JSON.stringify(futureEvent)) + 1) +
      128
    if (reserved > MAX_PRESENTATION_LIFECYCLE_BYTES) invalid('output_too_large')
  }
  recordDeletionResult(
    input: PresentationLifecycleScope,
    revision: number,
    inputResult: PresentationDeletionResult,
  ) {
    const result = owned(inputResult)
    if (
      !exact(result, [
        'deletionId',
        'resourceId',
        'status',
        ...(result.code === undefined ? [] : ['code']),
      ]) ||
      !id(result.deletionId) ||
      !id(result.resourceId) ||
      !['removed', 'reference_removed', 'retained', 'failed'].includes(result.status) ||
      (result.code !== undefined && !codes.includes(result.code))
    )
      invalid()
    return this.transaction(input, (s) => {
      const r = this.current(s, revision)
      if (r.state !== 'deleting' || r.deletion?.deletionId !== result.deletionId)
        invalid('deletion_conflict')
      const entry = r.deletion.resources.find((v) => v.resourceId === result.resourceId)
      if (!entry) invalid()
      const next = {
        ...entry,
        status: result.status,
        ...(result.code === undefined ? {} : { code: result.code }),
      }
      if (result.code === undefined) delete next.code
      resource(next, true)
      if (
        ['removed', 'reference_removed'].includes(entry.status) &&
        JSON.stringify(entry) !== JSON.stringify(next)
      )
        invalid()
      if (JSON.stringify(entry) === JSON.stringify(next)) return r
      Object.assign(entry, next)
      if (result.code === undefined) delete entry.code
      this.append(
        r,
        'resource_result',
        r.deletion.resources.some((v) => !['removed', 'reference_removed'].includes(v.status))
          ? 'partial'
          : 'complete',
      )
      // A retry must never consume the remaining successful receipts and final checkpoint.
      this.reserveDeletion(
        r,
        r.deletion.resources.filter((v) => !['removed', 'reference_removed'].includes(v.status))
          .length + 1,
      )
      return this.persist(s, r)
    })
  }
  finishDeletion(input: PresentationLifecycleScope, revision: number, deletionId: string) {
    if (!id(deletionId)) invalid()
    return this.transaction(input, (s) => {
      const r = this.current(s, revision)
      if (r.state !== 'deleting' || r.deletion?.deletionId !== deletionId)
        invalid('deletion_conflict')
      if (r.deletion.resources.some((v) => !['removed', 'reference_removed'].includes(v.status)))
        invalid('deletion_incomplete')
      r.state = 'deleted'
      this.append(r, 'deletion_finished', 'complete')
      return this.persist(s, r)
    })
  }
  exportAudit(input: PresentationLifecycleScope) {
    const r = this.read(input)
    if (!r) invalid('project_not_found')
    return {
      version: 1 as const,
      anonymousProjectId: r.anonymousProjectId,
      auditRetentionDays: r.policy.auditRetentionDays,
      events: structuredClone(r.audit),
    }
  }
}
