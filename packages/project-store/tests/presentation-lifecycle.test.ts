import { afterEach, expect, it, vi } from 'vitest'
import {
  mkdtempSync,
  mkdirSync,
  rmSync,
  readdirSync,
  symlinkSync,
  readFileSync,
  writeFileSync,
  openSync,
  renameSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import {
  PresentationLifecycleStore,
  DEFAULT_PRESENTATION_RETENTION_POLICY,
  parsePresentationLifecycle,
} from '../src/presentation-lifecycle.js'
vi.mock('node:fs', async (importOriginal) => {
  const fs = await importOriginal<typeof import('node:fs')>()
  return { ...fs, openSync: vi.fn(fs.openSync), renameSync: vi.fn(fs.renameSync) }
})
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'presentation-lifecycle-'))
  roots.push(root)
  return {
    root,
    store: new PresentationLifecycleStore(root),
    scope: { projectId: 'p', documentId: 'opaque:/sensitive-document' },
  }
}
const resource = (resourceId = 'own') => ({
  resourceId,
  kind: 'project' as const,
  ownership: 'project_exclusive' as const,
})
it('persists explicit preserving defaults and exact document/project binding across reopen', () => {
  const f = fixture(),
    record = f.store.initialize(f.scope)
  expect(record).toMatchObject({
    version: 1,
    revision: 0,
    state: 'active',
    policy: { contentRetentionDays: null, auditRetentionDays: null },
  })
  expect(DEFAULT_PRESENTATION_RETENTION_POLICY).toEqual(record.policy)
  expect(new PresentationLifecycleStore(f.root).read(f.scope)).toEqual(record)
  expect(() => f.store.read({ ...f.scope, documentId: 'foreign' })).toThrow('document_mismatch')
  expect(() => f.store.initialize({ ...f.scope, projectId: '../escape' })).toThrow(
    'invalid_request',
  )
  expect(f.store.read({ ...f.scope, projectId: 'other' })).toBeUndefined()
})
it('uses disk revision CAS across two actual store instances without lost policies', async () => {
  const f = fixture(),
    second = new PresentationLifecycleStore(f.root)
  f.store.initialize(f.scope)
  const results = await Promise.allSettled([
    Promise.resolve().then(() =>
      f.store.setPolicy(f.scope, 0, { contentRetentionDays: 30, auditRetentionDays: 365 }),
    ),
    Promise.resolve().then(() =>
      second.setPolicy(f.scope, 0, { contentRetentionDays: 60, auditRetentionDays: null }),
    ),
  ])
  expect(results.filter((r) => r.status === 'fulfilled')).toHaveLength(1)
  expect(results.filter((r) => r.status === 'rejected')).toHaveLength(1)
  expect(second.read(f.scope)).toMatchObject({
    revision: 1,
    policy: { contentRetentionDays: 30, auditRetentionDays: 365 },
  })
})
it('persists deletion intent and every partial resource result without deleting any project content', () => {
  const f = fixture()
  f.store.initialize(f.scope)
  const begin = f.store.beginDeletion(f.scope, 0, {
    deletionId: 'delete-1',
    reason: 'user',
    resources: [
      resource(),
      { resourceId: 'shared', kind: 'attachments', ownership: 'shared_reference' },
      { resourceId: 'unknown', kind: 'master_backups', ownership: 'unproven' },
    ],
  })
  expect(begin.state).toBe('deleting')
  expect(() => f.store.assertActive(f.scope)).toThrow('project_deleting')
  expect(() =>
    f.store.setPolicy(f.scope, 1, { contentRetentionDays: 1, auditRetentionDays: 1 }),
  ).toThrow('project_deleting')
  const fresh = new PresentationLifecycleStore(f.root)
  const one = fresh.recordDeletionResult(f.scope, 1, {
    deletionId: 'delete-1',
    resourceId: 'own',
    status: 'removed',
  })
  expect(one.revision).toBe(2)
  expect(() =>
    fresh.recordDeletionResult(f.scope, 2, {
      deletionId: 'delete-1',
      resourceId: 'shared',
      status: 'removed',
    }),
  ).toThrow('invalid_request')
  fresh.recordDeletionResult(f.scope, 2, {
    deletionId: 'delete-1',
    resourceId: 'shared',
    status: 'reference_removed',
  })
  fresh.recordDeletionResult(f.scope, 3, {
    deletionId: 'delete-1',
    resourceId: 'unknown',
    status: 'retained',
    code: 'ownership_unproven',
  })
  expect(() => fresh.finishDeletion(f.scope, 4, 'delete-1')).toThrow('deletion_incomplete')
  expect(new PresentationLifecycleStore(f.root).read(f.scope)).toMatchObject({
    state: 'deleting',
    revision: 4,
  })
  expect(readdirSync(join(f.root, 'presentation-project-lifecycles'))).toHaveLength(1)
})
it('allows deleted only after all authorized resources have resolved and remains terminal on reopen', () => {
  const f = fixture()
  f.store.initialize(f.scope)
  f.store.beginDeletion(f.scope, 0, {
    deletionId: 'delete',
    reason: 'retention',
    resources: [
      resource(),
      { resourceId: 'shared', kind: 'brand_kits', ownership: 'shared_reference' },
    ],
  })
  f.store.recordDeletionResult(f.scope, 1, {
    deletionId: 'delete',
    resourceId: 'own',
    status: 'failed',
    code: 'io_failed',
  })
  expect(() => f.store.finishDeletion(f.scope, 2, 'delete')).toThrow('deletion_incomplete')
  f.store.recordDeletionResult(f.scope, 2, {
    deletionId: 'delete',
    resourceId: 'own',
    status: 'removed',
  })
  f.store.recordDeletionResult(f.scope, 3, {
    deletionId: 'delete',
    resourceId: 'shared',
    status: 'reference_removed',
  })
  const final = f.store.finishDeletion(f.scope, 4, 'delete')
  expect(final.state).toBe('deleted')
  expect(() => new PresentationLifecycleStore(f.root).assertActive(f.scope)).toThrow(
    'project_deleted',
  )
  expect(() =>
    f.store.beginDeletion(f.scope, 5, {
      deletionId: 'again',
      reason: 'user',
      resources: [resource()],
    }),
  ).toThrow('project_deleted')
})
it('exports bounded minimal anonymous audit without private identifiers, paths, URLs, body or raw error', () => {
  const f = fixture()
  f.store.initialize(f.scope)
  f.store.beginDeletion(f.scope, 0, {
    deletionId: 'private-delete-id',
    reason: 'user',
    resources: [resource('private-resource-id')],
  })
  f.store.recordDeletionResult(f.scope, 1, {
    deletionId: 'private-delete-id',
    resourceId: 'private-resource-id',
    status: 'failed',
    code: 'io_failed',
  })
  const audit = f.store.exportAudit(f.scope),
    serialized = JSON.stringify(audit)
  expect(audit.events).toHaveLength(3)
  expect(audit.anonymousProjectId).toMatch(/^[a-f0-9-]{36}$/)
  for (const value of [
    f.scope.projectId + '"',
    f.scope.documentId,
    'private-delete-id',
    'private-resource-id',
    f.root,
    'http',
    'sha256',
    'body',
    'error',
  ])
    expect(serialized).not.toContain(value)
  expect(audit.events[2]).toMatchObject({
    action: 'resource_result',
    result: 'partial',
    counts: { failed: 1 },
  })
})
it('expires a completed anonymous audit while retaining an irreversible identity tombstone', () => {
  const f = fixture()
  const initial = f.store.initialize(f.scope)
  const policy = f.store.setPolicy(f.scope, initial.revision, {
    contentRetentionDays: null,
    auditRetentionDays: 1,
  })
  const started = f.store.beginDeletion(f.scope, policy.revision, {
    deletionId: 'delete',
    reason: 'user',
    resources: [resource()],
  })
  const result = f.store.recordDeletionResult(f.scope, started.revision, {
    deletionId: 'delete',
    resourceId: 'own',
    status: 'removed',
  })
  const deleted = f.store.finishDeletion(f.scope, result.revision, 'delete')
  const directory = join(
    f.root,
    'presentation-project-lifecycles',
    createHash('sha256').update(f.scope.projectId).digest('hex'),
  )
  expect(
    f.store.pruneExpiredAudit(f.scope, new Date(Date.parse(deleted.updatedAt) + 86400000 - 1)),
  ).toBe(false)
  expect(
    f.store.pruneExpiredAudit(f.scope, new Date(Date.parse(deleted.updatedAt) + 86400000)),
  ).toBe(true)
  expect(readdirSync(directory)).toEqual(['lifecycle.json'])
  const marker = readFileSync(join(directory, 'lifecycle.json'), 'utf8')
  expect(marker).not.toContain(f.scope.projectId + '"')
  expect(marker).not.toContain(f.scope.documentId)
  expect(marker).not.toContain('"deletionId"')
  expect(marker).not.toContain('"resourceId"')
  expect(() => parsePresentationLifecycle(JSON.parse(marker))).toThrow('invalid_state')
  expect(() => f.store.readControl(f.scope)).toThrow('project_deleted')
  expect(() => f.store.initialize(f.scope)).toThrow('project_deleted')
  expect(() => f.store.exportAudit(f.scope)).toThrow('project_deleted')
  expect(() => f.store.readControl({ ...f.scope, documentId: 'other' })).toThrow(
    'document_mismatch',
  )
  expect(
    f.store.pruneExpiredAudit(f.scope, new Date(Date.parse(deleted.updatedAt) + 86400000)),
  ).toBe(false)
})
it('retains the old audit when atomic tombstone publication fails and retries cleanly', () => {
  const f = fixture()
  const initial = f.store.initialize(f.scope, {
    contentRetentionDays: null,
    auditRetentionDays: 1,
  })
  const started = f.store.beginDeletion(f.scope, initial.revision, {
    deletionId: 'delete',
    reason: 'user',
    resources: [resource()],
  })
  const result = f.store.recordDeletionResult(f.scope, started.revision, {
    deletionId: 'delete',
    resourceId: 'own',
    status: 'removed',
  })
  const deleted = f.store.finishDeletion(f.scope, result.revision, 'delete')
  const now = new Date(Date.parse(deleted.updatedAt) + 86400000)
  vi.mocked(renameSync).mockImplementationOnce(() => {
    throw Error('simulated_rename_failure')
  })
  expect(() => f.store.pruneExpiredAudit(f.scope, now)).toThrow('simulated_rename_failure')
  expect(f.store.exportAudit(f.scope).events.at(-1)?.action).toBe('deletion_finished')
  expect(f.store.pruneExpiredAudit(f.scope, now)).toBe(true)
  expect(() => f.store.initialize(f.scope)).toThrow('project_deleted')
})
it.each([
  { contentRetentionDays: 0, auditRetentionDays: null },
  { contentRetentionDays: null, auditRetentionDays: 36501 },
  { contentRetentionDays: 1, auditRetentionDays: null, body: 'secret' },
])('rejects unknown or invalid policies without writing', (policy) => {
  const f = fixture()
  f.store.initialize(f.scope)
  expect(() => f.store.setPolicy(f.scope, 0, policy)).toThrow('invalid_request')
  expect(f.store.read(f.scope)!.revision).toBe(0)
})
it('rejects mismatched existing PresentationStore ownership and symlinked ancestors or metadata', () => {
  const f = fixture()
  f.store.initialize(f.scope)
  const directory = join(
    f.root,
    'projects',
    'presentations',
    createHash('sha256').update('p').digest('hex'),
  )
  mkdirSync(directory, { recursive: true })
  writeFileSync(
    join(directory, 'project.json'),
    JSON.stringify({ version: 1, projectId: 'p', documentId: 'another' }),
  )
  expect(() => f.store.read(f.scope)).toThrow('document_mismatch')
  rmSync(join(directory, 'project.json'))
  const foreign = fixture()
  const lifecycle = join(
    f.root,
    'presentation-project-lifecycles',
    createHash('sha256').update('p').digest('hex'),
    'lifecycle.json',
  )
  rmSync(lifecycle)
  symlinkSync(join(foreign.root, 'foreign.json'), lifecycle)
  expect(() => f.store.read(f.scope)).toThrow('invalid_state')
  const unsafe = fixture()
  symlinkSync(f.root, join(unsafe.root, 'projects'))
  expect(() => unsafe.store.initialize(unsafe.scope)).toThrow('invalid_state')
})
it('rejects corrupted or over-budget state on reopen without silently truncating or resetting it', () => {
  const f = fixture()
  f.store.initialize(f.scope)
  const path = join(
    f.root,
    'presentation-project-lifecycles',
    createHash('sha256').update('p').digest('hex'),
    'lifecycle.json',
  )
  const original = readFileSync(path, 'utf8')
  writeFileSync(path, original.replace('"version":1', '"version":99'))
  expect(() => new PresentationLifecycleStore(f.root).read(f.scope)).toThrow('invalid_state')
  expect(readFileSync(path, 'utf8')).toContain('"version":99')
  writeFileSync(path, ' '.repeat(2 * 1024 * 1024 + 1))
  expect(() => f.store.read(f.scope)).toThrow('invalid_state')
})
it('rejects duplicate, unbounded and accessor-bearing deletion inputs without changing active state', () => {
  const f = fixture()
  f.store.initialize(f.scope)
  for (const resources of [
    [resource(), resource()],
    Array.from({ length: 4097 }, (_, i) => resource(`r${i}`)),
    [
      {
        ...resource(),
        get body() {
          throw Error('getter executed')
        },
      },
    ],
  ])
    expect(() =>
      f.store.beginDeletion(f.scope, 0, { deletionId: 'd', reason: 'user', resources }),
    ).toThrow('invalid_request')
  expect(f.store.read(f.scope)!.state).toBe('active')
})
it('keeps lifecycle tombstones outside content and remains compatible with actual PresentationStore complete-history reads', async () => {
  const { PresentationStore } = await import('../src/presentation-store.js')
  const f = fixture()
  f.store.initialize(f.scope)
  expect(() =>
    new PresentationStore(f.root).plan(f.scope.projectId, f.scope.documentId, true),
  ).not.toThrow()
  expect(
    new PresentationStore(f.root).plan(f.scope.projectId, f.scope.documentId, true),
  ).toBeUndefined()
  new PresentationStore(f.root).begin(f.scope.projectId, f.scope.documentId, 'request', {
    title: 'protected content',
  })
  expect(f.store.read(f.scope)!.state).toBe('active')
  expect(() =>
    new PresentationLifecycleStore(f.root).read({ ...f.scope, documentId: 'foreign' }),
  ).toThrow('document_mismatch')
})
it('rejects preexisting orphan project content without assigning lifecycle ownership', () => {
  const f = fixture(),
    directory = join(
      f.root,
      'projects',
      'presentations',
      createHash('sha256').update('p').digest('hex'),
    )
  mkdirSync(directory, { recursive: true })
  writeFileSync(
    join(directory, 'plan.json'),
    JSON.stringify({ projectId: 'p', documentId: 'foreign' }),
  )
  expect(() => f.store.initialize(f.scope)).toThrow('invalid_state')
})
it('rejects impossible persisted audit action order but accepts reordered JSON count keys', () => {
  const f = fixture(),
    record = f.store.initialize(f.scope),
    corrupt = structuredClone(record)
  corrupt.revision = 1
  corrupt.audit.push({
    sequence: 1,
    at: record.updatedAt,
    action: 'resource_result',
    result: 'partial',
    counts: { pending: 0, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 },
  })
  expect(() => parsePresentationLifecycle(corrupt)).toThrow('invalid_state')
  record.audit[0]!.counts = { failed: 0, retained: 0, referenceRemoved: 0, removed: 0, pending: 0 }
  expect(() => parsePresentationLifecycle(record)).not.toThrow()
})
it('keeps authoritative prior state after failed publication and reconciles lost publish ACK by fresh disk CAS', async () => {
  const f = fixture()
  f.store.initialize(f.scope)
  vi.mocked(renameSync).mockImplementationOnce(() => {
    throw Error('disk_write_failed')
  })
  expect(() =>
    f.store.setPolicy(f.scope, 0, { contentRetentionDays: 30, auditRetentionDays: 90 }),
  ).toThrow('disk_write_failed')
  expect(new PresentationLifecycleStore(f.root).read(f.scope)).toMatchObject({
    revision: 0,
    policy: DEFAULT_PRESENTATION_RETENTION_POLICY,
  })
  const directory = join(
    f.root,
    'presentation-project-lifecycles',
    createHash('sha256').update('p').digest('hex'),
  )
  expect(readdirSync(directory)).toEqual(['lifecycle.json'])
  const real = await vi.importActual<typeof import('node:fs')>('node:fs')
  vi.mocked(renameSync).mockImplementationOnce((from, to) => {
    real.renameSync(from, to)
    throw Error('lost_publish_ack')
  })
  expect(() =>
    f.store.setPolicy(f.scope, 0, { contentRetentionDays: 30, auditRetentionDays: 90 }),
  ).toThrow('lost_publish_ack')
  expect(new PresentationLifecycleStore(f.root).read(f.scope)).toMatchObject({
    revision: 1,
    policy: { contentRetentionDays: 30, auditRetentionDays: 90 },
  })
  expect(() =>
    f.store.setPolicy(f.scope, 0, { contentRetentionDays: 60, auditRetentionDays: null }),
  ).toThrow('revision_conflict')
})
it('fails closed when a trusted parent is substituted at a metadata leaf open', async () => {
  const f = fixture()
  f.store.initialize(f.scope)
  const directory = join(
      f.root,
      'presentation-project-lifecycles',
      createHash('sha256').update('p').digest('hex'),
    ),
    path = join(directory, 'lifecycle.json'),
    foreign = fixture()
  writeFileSync(join(foreign.root, 'lifecycle.json'), readFileSync(path))
  const real = await vi.importActual<typeof import('node:fs')>('node:fs')
  vi.mocked(openSync).mockImplementationOnce((input, flags, mode) => {
    expect(input).toBe(path)
    real.renameSync(directory, directory + '-old')
    real.symlinkSync(foreign.root, directory)
    return real.openSync(input, flags, mode)
  })
  expect(() => f.store.read(f.scope)).toThrow('invalid_state')
  expect(readFileSync(join(foreign.root, 'lifecycle.json'), 'utf8')).toContain('"revision":0')
})
it('reserves bounded result and audit space before accepting a large deletion intent', () => {
  const f = fixture(),
    record = f.store.initialize(f.scope)
  for (let i = 1; i <= 3000; i++)
    record.audit.push({
      sequence: i,
      at: record.updatedAt,
      action: 'policy_updated',
      result: 'accepted',
      counts: { pending: 0, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 },
    })
  record.revision = 3000
  const path = join(
    f.root,
    'presentation-project-lifecycles',
    createHash('sha256').update('p').digest('hex'),
    'lifecycle.json',
  )
  writeFileSync(path, JSON.stringify(record))
  expect(() =>
    f.store.beginDeletion(f.scope, 3000, {
      deletionId: 'large',
      reason: 'user',
      resources: Array.from({ length: 4096 }, (_, i) => resource(`r${i}`.padEnd(128, 'x'))),
    }),
  ).toThrow('output_too_large')
  expect(f.store.read(f.scope)).toMatchObject({ state: 'active', revision: 3000 })
})
it('rejects an intent whose first failure and successful retry receipts cannot fit', () => {
  const f = fixture()
  f.store.initialize(f.scope)
  expect(() =>
    f.store.beginDeletion(f.scope, 0, {
      deletionId: 'large',
      reason: 'user',
      resources: Array.from({ length: 4096 }, (_, i) => resource(`r${i}`)),
    }),
  ).toThrow('output_too_large')
  expect(f.store.read(f.scope)?.state).toBe('active')
})
it('preserves remaining success and final checkpoint space when retries exhaust audit capacity', () => {
  const f = fixture(),
    r = f.store.initialize(f.scope)
  for (let i = 1; i <= 8185; i++)
    r.audit.push({
      sequence: i,
      at: r.updatedAt,
      action: 'policy_updated',
      result: 'accepted',
      counts: { pending: 0, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 },
    })
  r.revision = 8185
  const path = join(
    f.root,
    'presentation-project-lifecycles',
    createHash('sha256').update('p').digest('hex'),
    'lifecycle.json',
  )
  writeFileSync(path, JSON.stringify(r))
  let current = f.store.beginDeletion(f.scope, r.revision, {
    deletionId: 'd',
    reason: 'user',
    resources: [resource('a'), resource('b')],
  })
  for (const resourceId of ['a', 'b'])
    current = f.store.recordDeletionResult(f.scope, current.revision, {
      deletionId: 'd',
      resourceId,
      status: 'failed',
      code: 'io_failed',
    })
  expect(() =>
    f.store.recordDeletionResult(f.scope, current.revision, {
      deletionId: 'd',
      resourceId: 'a',
      status: 'failed',
      code: 'resource_busy',
    }),
  ).toThrow('output_too_large')
  for (const resourceId of ['a', 'b'])
    current = f.store.recordDeletionResult(f.scope, current.revision, {
      deletionId: 'd',
      resourceId,
      status: 'removed',
    })
  expect(f.store.finishDeletion(f.scope, current.revision, 'd').state).toBe('deleted')
})
it('keeps minimal deletion failure, retry and closure capacity when setting active policy', () => {
  const f = fixture(),
    r = f.store.initialize(f.scope)
  for (let i = 1; i <= 8187; i++)
    r.audit.push({
      sequence: i,
      at: r.updatedAt,
      action: 'policy_updated',
      result: 'accepted',
      counts: { pending: 0, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 },
    })
  r.revision = 8187
  const path = join(
    f.root,
    'presentation-project-lifecycles',
    createHash('sha256').update('p').digest('hex'),
    'lifecycle.json',
  )
  writeFileSync(path, JSON.stringify(r))
  expect(() =>
    f.store.setPolicy(f.scope, r.revision, { contentRetentionDays: 30, auditRetentionDays: null }),
  ).toThrow('output_too_large')
  expect(f.store.read(f.scope)?.revision).toBe(r.revision)
  let current = f.store.beginDeletion(f.scope, r.revision, {
    deletionId: 'minimal',
    reason: 'user',
    resources: [resource()],
  })
  current = f.store.recordDeletionResult(f.scope, current.revision, {
    deletionId: 'minimal',
    resourceId: 'own',
    status: 'failed',
    code: 'io_failed',
  })
  current = f.store.recordDeletionResult(f.scope, current.revision, {
    deletionId: 'minimal',
    resourceId: 'own',
    status: 'removed',
  })
  expect(f.store.finishDeletion(f.scope, current.revision, 'minimal').state).toBe('deleted')
})
