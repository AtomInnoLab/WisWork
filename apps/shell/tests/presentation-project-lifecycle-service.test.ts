import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createPresentationProjectLifecycleService } from '../src/main/presentation-project-lifecycle'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'presentation-lifecycle-service-'))
  roots.push(root)
  return { root, service: createPresentationProjectLifecycleService({ userDataPath: root }) }
}
const scope = { projectId: 'private-project', documentId: 'private-document' }
const request = (operation: string, extra: Record<string, unknown> = {}) => ({
  operation,
  ...scope,
  ...extra,
})
it('reads missing policy without creating a project and explicitly initializes durable defaults', () => {
  const f = fixture()
  expect(f.service(request('project_lifecycle_read'))).toEqual({ lifecycle: null })
  const first = f.service(request('project_lifecycle_initialize')) as any
  expect(first.lifecycle).toMatchObject({
    state: 'active',
    revision: 0,
    policy: { contentRetentionDays: null, auditRetentionDays: null },
  })
  const reopened = createPresentationProjectLifecycleService({ userDataPath: f.root })
  expect(reopened(request('project_lifecycle_read'))).toEqual(first)
  expect(reopened(request('project_lifecycle_initialize'))).toEqual(first)
})
it('updates policy through exact revision CAS and retains the prior record on stale writes', () => {
  const f = fixture()
  f.service(request('project_lifecycle_initialize'))
  const policy = { contentRetentionDays: 30, auditRetentionDays: 365 }
  expect(
    f.service(request('project_lifecycle_set_policy', { expectedRevision: 0, policy })),
  ).toMatchObject({ lifecycle: { revision: 1, policy } })
  expect(
    f.service(request('project_lifecycle_set_policy', { expectedRevision: 0, policy })),
  ).toEqual({ error: 'revision_conflict' })
  expect(f.service(request('project_lifecycle_read'))).toMatchObject({
    lifecycle: { revision: 1, policy },
  })
})
it('exports only anonymous durable audit and rejects a different document binding', () => {
  const f = fixture()
  f.service(request('project_lifecycle_initialize'))
  const exported = f.service(request('project_lifecycle_export_audit')) as any
  expect(exported.audit).toMatchObject({
    version: 1,
    auditRetentionDays: null,
    events: [{ action: 'created' }],
  })
  expect(JSON.stringify(exported)).not.toContain(scope.projectId)
  expect(JSON.stringify(exported)).not.toContain(scope.documentId)
  expect(JSON.stringify(exported)).not.toContain(f.root)
  expect(
    f.service({ ...request('project_lifecycle_read'), documentId: 'different-document' }),
  ).toEqual({ error: 'document_mismatch' })
})
for (const extra of [
  { unknown: true },
  { expectedRevision: 0 },
  { policy: { contentRetentionDays: 30, auditRetentionDays: null } },
]) {
  it(`rejects extra initialization fields ${Object.keys(extra)[0]} without creating metadata`, () => {
    const f = fixture()
    expect(f.service(request('project_lifecycle_initialize', extra))).toEqual({
      error: 'invalid_request',
    })
    expect(f.service(request('project_lifecycle_read'))).toEqual({ lifecycle: null })
  })
}
it('rejects malformed budgets and policies without raw error text', () => {
  const f = fixture()
  f.service(request('project_lifecycle_initialize'))
  for (const policy of [
    { contentRetentionDays: 0, auditRetentionDays: null },
    { contentRetentionDays: 30, auditRetentionDays: null, rawPath: 'private' },
  ]) {
    expect(
      f.service(request('project_lifecycle_set_policy', { expectedRevision: 0, policy })),
    ).toEqual({ error: 'invalid_request' })
  }
  const bad = Object.defineProperty({}, 'operation', {
    enumerable: true,
    get: () => {
      throw Error('private-token=secret')
    },
  })
  expect(f.service(bad)).toEqual({ error: 'invalid_request' })
  expect(f.service(request('project_lifecycle_delete'))).toEqual({ error: 'invalid_request' })
})

it('reads hostile error messages once and suppresses a throwing accessor', () => {
  const f = fixture()
  let reads = 0
  const error = new Error()
  Object.defineProperty(error, 'message', {
    get: () => (++reads === 1 ? 'invalid_state' : 'private-token=secret'),
  })
  const input = new Proxy(
    {},
    {
      getPrototypeOf: () => {
        throw error
      },
    },
  )
  expect(f.service(input)).toEqual({ error: 'invalid_state' })
  expect(reads).toBe(1)
  const throwing = new Error()
  Object.defineProperty(throwing, 'message', {
    get: () => {
      throw Error('private-token=secret')
    },
  })
  expect(
    f.service(
      new Proxy(
        {},
        {
          getPrototypeOf: () => {
            throw throwing
          },
        },
      ),
    ),
  ).toEqual({ error: 'invalid_state' })
})

it('supports all governance operations at the research document boundary without leaking audit identity', () => {
  const f = fixture(),
    documentId = 'd'.repeat(4096)
  const body = (operation: string, extra: Record<string, unknown> = {}) =>
    request(operation, { documentId, ...extra })
  expect(f.service(body('project_lifecycle_read'))).toEqual({ lifecycle: null })
  expect(f.service(body('project_lifecycle_initialize'))).toMatchObject({
    lifecycle: { documentId, revision: 0 },
  })
  const policy = { contentRetentionDays: 30, auditRetentionDays: null }
  const updated = f.service(body('project_lifecycle_set_policy', { expectedRevision: 0, policy }))
  expect(updated).toMatchObject({ lifecycle: { documentId, revision: 1, policy } })
  const reopened = createPresentationProjectLifecycleService({ userDataPath: f.root })
  expect(reopened(body('project_lifecycle_read'))).toEqual(updated)
  const audit = reopened(body('project_lifecycle_export_audit'))
  expect(audit).toMatchObject({
    audit: { version: 1, events: [{ action: 'created' }, { action: 'policy_updated' }] },
  })
  expect(JSON.stringify(audit)).not.toContain(documentId)
  expect(JSON.stringify(audit)).not.toContain(scope.projectId)
  expect(JSON.stringify(audit)).not.toContain(f.root)
  for (const operation of [
    'project_lifecycle_initialize',
    'project_lifecycle_read',
    'project_lifecycle_set_policy',
    'project_lifecycle_export_audit',
  ]) {
    expect(
      reopened(
        body(operation, {
          documentId: 'd'.repeat(4097),
          ...(operation === 'project_lifecycle_set_policy' ? { expectedRevision: 1, policy } : {}),
        }),
      ),
    ).toEqual({ error: 'invalid_request' })
  }
  expect(reopened(body('project_lifecycle_read'))).toEqual(updated)
})
