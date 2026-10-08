import { afterEach, expect, it } from 'vitest'
import {
  mkdtempSync,
  rmSync,
  readdirSync,
  existsSync,
  unlinkSync,
  mkdirSync,
  writeFileSync,
} from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import {
  PresentationStore,
  PresentationLifecycleStore,
  parsePresentationLifecycle,
} from '@wiswork/project-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import {
  registerPresentationProjectWork,
  stopPresentationProjectWork,
} from '../src/main/presentation-project-work'
import { createPresentationProjectGovernanceService } from '../src/main/presentation-project-governance'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture(anchor = true) {
  const root = mkdtempSync(join(tmpdir(), 'governance-adapter-'))
  roots.push(root)
  const plan = benchmarkPlan(),
    scope = { projectId: plan.projectId, documentId: 'private-doc' },
    store = new PresentationStore(root),
    life = new PresentationLifecycleStore(root)
  if (anchor) store.savePlan(scope.projectId, scope.documentId, 0, plan)
  let tail = Promise.resolve(),
    held = false,
    admissions = 0
  const lock = async () => {
    admissions++
    const previous = tail
    let release!: () => void
    tail = new Promise<void>((r) => {
      release = r
    })
    await previous
    expect(held).toBe(false)
    held = true
    return () => {
      held = false
      release()
    }
  }
  const service = createPresentationProjectGovernanceService({
      userDataPath: root,
      acquireProjectLock: lock,
    }),
    call = (operation: string, fields = {}) =>
      service({ operation, ...scope, ...fields }, new AbortController().signal)
  return {
    root,
    scope,
    store,
    life,
    lock,
    service,
    call,
    held: () => held,
    admissions: () => admissions,
  }
}
it('pure read and preview never initialize controls; explicit confirmation deletes and control/audit remain readable', async () => {
  const f = fixture()
  expect(await f.call('project_lifecycle_read')).toEqual({ lifecycle: null })
  const p = (await f.call('project_deletion_preview')) as any
  expect(p.preview.resources).toEqual([
    expect.objectContaining({ kind: 'project', ownership: 'project_exclusive' }),
  ])
  expect(f.life.readControl(f.scope)).toBeUndefined()
  expect(
    await f.call('project_deletion_confirm', {
      expectedRevision: p.preview.expectedRevision,
      confirmationToken: p.preview.confirmationToken,
      deletionId: 'delete',
    }),
  ).toMatchObject({ deletion: { state: 'deleted' } })
  expect(await f.call('project_lifecycle_read')).toMatchObject({ lifecycle: { state: 'deleted' } })
  const audit = await f.call('project_lifecycle_export_audit')
  expect(audit).toMatchObject({
    audit: {
      events: [expect.any(Object), expect.any(Object), expect.any(Object), expect.any(Object)],
    },
  })
  for (const secret of [f.scope.projectId, f.scope.documentId, f.root])
    expect(JSON.stringify(audit)).not.toContain(secret)
})
it('explicit initialization requires real project anchor and never grants nonactive writes', async () => {
  const f = fixture(false)
  expect(await f.call('project_lifecycle_initialize')).toEqual({ error: 'project_not_found' })
  expect(readdirSync(f.root)).toEqual([])
  f.store.savePlan(f.scope.projectId, f.scope.documentId, 0, benchmarkPlan())
  expect(await f.call('project_lifecycle_initialize')).toMatchObject({
    lifecycle: {
      revision: 0,
      state: 'active',
      policy: { contentRetentionDays: null, auditRetentionDays: null },
    },
  })
  const p = (await f.call('project_deletion_preview')) as any
  f.life.beginDeletion(f.scope, 0, {
    deletionId: 'd',
    reason: 'user',
    resources: p.preview.resources.map((r: any) => ({
      resourceId: r.resourceId,
      kind: r.kind,
      ownership: r.ownership,
    })),
  })
  expect(await f.call('project_lifecycle_initialize')).toEqual({ error: 'project_deleting' })
  expect(
    await f.call('project_lifecycle_set_policy', {
      expectedRevision: 1,
      policy: { contentRetentionDays: 1, auditRetentionDays: null },
    }),
  ).toEqual({ error: 'project_deleting' })
})
it('set policy needs existing numeric revision and preserves original CAS', async () => {
  const f = fixture()
  const policy = { contentRetentionDays: 30, auditRetentionDays: 365 }
  expect(await f.call('project_lifecycle_set_policy', { expectedRevision: 0, policy })).toEqual({
    error: 'project_not_found',
  })
  await f.call('project_lifecycle_initialize')
  expect(
    await f.call('project_lifecycle_set_policy', { expectedRevision: 0, policy }),
  ).toMatchObject({ lifecycle: { revision: 1, policy } })
  expect(
    await f.call('project_lifecycle_set_policy', {
      expectedRevision: 0,
      policy: { contentRetentionDays: 1, auditRetentionDays: null },
    }),
  ).toEqual({ error: 'revision_conflict' })
  expect(f.life.readControl(f.scope)?.policy).toEqual(policy)
})
it('readControl works with missing body binding but does not authorize initialize or foreign audit', async () => {
  const f = fixture()
  await f.call('project_lifecycle_initialize')
  const path = join(
    f.root,
    'projects',
    'presentations',
    createHash('sha256').update(f.scope.projectId).digest('hex'),
    'project.json',
  )
  unlinkSync(path)
  expect(await f.call('project_lifecycle_read')).toMatchObject({ lifecycle: { state: 'active' } })
  expect(await f.call('project_lifecycle_export_audit')).toMatchObject({
    audit: { anonymousProjectId: expect.any(String) },
  })
  expect(await f.call('project_lifecycle_initialize')).toEqual({ error: 'project_not_found' })
  expect(await f.call('project_lifecycle_export_audit', { documentId: 'foreign' })).toEqual({
    error: 'document_mismatch',
  })
})
it('queued initialize owns original absence and rejects intervening control creation', async () => {
  const f = fixture(),
    release = await f.lock(),
    pending = f.call('project_lifecycle_initialize')
  await Promise.resolve()
  f.life.initialize(f.scope)
  release()
  expect(await pending).toEqual({ error: 'revision_conflict' })
  expect(f.life.readControl(f.scope)?.revision).toBe(0)
})
it('queued policy snapshots aliases and rejects changed original version', async () => {
  const f = fixture()
  await f.call('project_lifecycle_initialize')
  const release = await f.lock(),
    body = {
      operation: 'project_lifecycle_set_policy',
      ...f.scope,
      expectedRevision: 0,
      policy: { contentRetentionDays: 30, auditRetentionDays: 365 },
    },
    pending = f.service(body, new AbortController().signal)
  body.projectId = 'foreign'
  body.documentId = 'foreign'
  body.policy.contentRetentionDays = 999
  f.life.setPolicy(f.scope, 0, { contentRetentionDays: 2, auditRetentionDays: null })
  release()
  expect(await pending).toEqual({ error: 'revision_conflict' })
  expect(f.life.readControl(f.scope)?.policy.contentRetentionDays).toBe(2)
})
it.each(['project_lifecycle_initialize', 'project_lifecycle_set_policy'])(
  'Work drains queued cancellation without changing state or retaining lock %s',
  async (operation) => {
    const f = fixture()
    if (operation === 'project_lifecycle_set_policy') await f.call('project_lifecycle_initialize')
    const release = await f.lock(),
      controller = new AbortController(),
      pending = f.service(
        {
          operation,
          ...f.scope,
          ...(operation === 'project_lifecycle_set_policy'
            ? { expectedRevision: 0, policy: { contentRetentionDays: 1, auditRetentionDays: null } }
            : {}),
        },
        controller.signal,
      )
    let drained = false
    const drain = stopPresentationProjectWork({ root: f.root, ...f.scope }).then(() => {
      drained = true
    })
    await Promise.resolve()
    expect(drained).toBe(false)
    release()
    expect(await pending).toEqual({ error: 'aborted' })
    await drain
    expect(f.held()).toBe(false)
    expect(f.life.readControl(f.scope)?.revision).toBe(
      operation === 'project_lifecycle_initialize' ? undefined : 0,
    )
  },
)
it('delete confirmation avoids self-registration while draining actual other work', async () => {
  const f = fixture(),
    p = (await f.call('project_deletion_preview')) as any,
    w = registerPresentationProjectWork({ scope: { root: f.root, ...f.scope } }),
    pending = f.call('project_deletion_confirm', {
      expectedRevision: p.preview.expectedRevision,
      confirmationToken: p.preview.confirmationToken,
      deletionId: 'd',
    })
  while (!w.signal.aborted) await Promise.resolve()
  expect(f.held()).toBe(false)
  w.finish()
  expect(await pending).toMatchObject({ deletion: { state: 'deleted' } })
})
it.each([
  { operation: 'unknown' },
  { operation: 'project_lifecycle_read', extra: true },
  {
    operation: 'project_lifecycle_set_policy',
    expectedRevision: null,
    policy: { contentRetentionDays: 1, auditRetentionDays: null },
  },
  { operation: 'project_lifecycle_read', documentId: 'a\u0000b' },
  {
    operation: 'project_deletion_confirm',
    expectedRevision: null,
    confirmationToken: 'a'.repeat(32768),
    deletionId: 'd',
  },
])(
  'strict fields/bounds reject before Work or project lock admission $operation',
  async (extra) => {
    const f = fixture(false)
    expect(await f.service({ ...f.scope, ...extra }, new AbortController().signal)).toEqual({
      error: 'invalid_request',
    })
    expect(f.admissions()).toBe(0)
    expect(readdirSync(f.root)).toEqual([])
  },
)
it('rejects getter-shaped input without running it or admitting work', async () => {
  const f = fixture(false),
    body = { operation: 'project_lifecycle_initialize', ...f.scope }
  let reads = 0
  Object.defineProperty(body, 'documentId', {
    enumerable: true,
    get() {
      reads++
      return 'private-secret'
    },
  })
  expect(await f.service(body, new AbortController().signal)).toEqual({ error: 'invalid_request' })
  expect(reads).toBe(0)
  expect(f.admissions()).toBe(0)
})
it.each(['dynamic', 'throws'])(
  'returns finite unknown errors despite hostile message getter %s',
  async (mode) => {
    const f = fixture(),
      service = createPresentationProjectGovernanceService({
        userDataPath: f.root,
        acquireProjectLock: async () => {
          const e = new Error()
          let reads = 0
          Object.defineProperty(e, 'message', {
            get() {
              if (mode === 'throws') throw Error('private-secret')
              return ++reads === 1 ? 'revision_conflict' : 'private-secret'
            },
          })
          throw e
        },
      })
    expect(
      await service(
        { operation: 'project_lifecycle_initialize', ...f.scope },
        new AbortController().signal,
      ),
    ).toEqual({ error: mode === 'dynamic' ? 'revision_conflict' : 'invalid_state' })
    expect(existsSync(join(f.root, 'presentation-project-lifecycles'))).toBe(false)
  },
)

it('unchanged queued policy uses owned values despite caller mutation', async () => {
  const f = fixture()
  await f.call('project_lifecycle_initialize')
  const release = await f.lock(),
    body = {
      operation: 'project_lifecycle_set_policy',
      ...f.scope,
      expectedRevision: 0,
      policy: { contentRetentionDays: 30, auditRetentionDays: 365 },
    },
    pending = f.service(body, new AbortController().signal)
  body.projectId = 'foreign'
  body.documentId = 'foreign'
  body.policy.contentRetentionDays = 999
  release()
  expect(await pending).toMatchObject({
    lifecycle: {
      projectId: f.scope.projectId,
      documentId: f.scope.documentId,
      revision: 1,
      policy: { contentRetentionDays: 30, auditRetentionDays: 365 },
    },
  })
  expect(f.life.readControl({ projectId: 'foreign', documentId: 'foreign' })).toBeUndefined()
})

it('resume avoids registering itself in drain after an explicitly cancelled confirmed intent', async () => {
  const f = fixture(),
    p = (await f.call('project_deletion_preview')) as {
      preview: { expectedRevision: null | number; confirmationToken: string }
    },
    w = registerPresentationProjectWork({ scope: { root: f.root, ...f.scope } }),
    controller = new AbortController(),
    first = f.service(
      {
        operation: 'project_deletion_confirm',
        ...f.scope,
        expectedRevision: p.preview.expectedRevision,
        confirmationToken: p.preview.confirmationToken,
        deletionId: 'd',
      },
      controller.signal,
    )
  while (!w.signal.aborted) await Promise.resolve()
  controller.abort()
  w.finish()
  expect(await first).toEqual({ error: 'aborted' })
  const r = f.life.readControl(f.scope)!
  expect(r.state).toBe('deleting')
  const next = registerPresentationProjectWork({ scope: { root: f.root, ...f.scope } }),
    pending = f.call('project_deletion_resume', { expectedRevision: r.revision, deletionId: 'd' })
  while (!next.signal.aborted) await Promise.resolve()
  expect(f.held()).toBe(false)
  next.finish()
  expect(await pending).toMatchObject({ deletion: { state: 'deleted' } })
  expect(await f.call('project_lifecycle_initialize')).toEqual({ error: 'project_deleted' })
  expect(
    await f.call('project_lifecycle_set_policy', {
      expectedRevision: f.life.readControl(f.scope)!.revision,
      policy: { contentRetentionDays: 1, auditRetentionDays: null },
    }),
  ).toEqual({ error: 'project_deleted' })
})
it('valid metadata-only4096scope initializes and reads;4097scope is denied without admission', async () => {
  const f = fixture(false),
    long = { ...f.scope, documentId: 'd'.repeat(4096) },
    dir = join(
      f.root,
      'projects',
      'presentations',
      createHash('sha256').update(f.scope.projectId).digest('hex'),
    )
  mkdirSync(dir, { recursive: true })
  writeFileSync(join(dir, 'project.json'), JSON.stringify({ version: 1, ...long }))
  expect(
    await f.service(
      { operation: 'project_lifecycle_initialize', ...long },
      new AbortController().signal,
    ),
  ).toMatchObject({ lifecycle: { documentId: long.documentId, state: 'active' } })
  const admissions = f.admissions()
  expect(
    await f.service(
      { operation: 'project_lifecycle_initialize', ...long, documentId: long.documentId + 'd' },
      new AbortController().signal,
    ),
  ).toEqual({ error: 'invalid_request' })
  expect(f.admissions()).toBe(admissions)
  expect(
    await f.service({ operation: 'project_lifecycle_read', ...long }, new AbortController().signal),
  ).toMatchObject({ lifecycle: { documentId: long.documentId } })
})

it('transports a strictly valid 2MiB lifecycle plus its exact14byte wrapper', async () => {
  const f = fixture(false),
    base = f.life.initialize(f.scope),
    limit = 2 * 1024 * 1024
  const build = (n: number) => {
    const resources = Array.from({ length: n }, (_, i) => ({
      resourceId: ('r' + i + '_').padEnd(128, 'x'),
      kind: 'research' as const,
      ownership: 'project_exclusive' as const,
      status: i === 0 ? ('failed' as const) : ('pending' as const),
      ...(i === 0 ? { code: 'io_failed' as const } : {}),
    }))
    return {
      ...base,
      documentId: 'd',
      state: 'deleting' as const,
      revision: 8191,
      deletion: { deletionId: 'd', reason: 'user' as const, resources },
      audit: Array.from({ length: 8192 }, (_, i) => ({
        sequence: i,
        at: base.createdAt,
        action:
          i === 0
            ? ('created' as const)
            : i === 1
              ? ('deletion_started' as const)
              : ('resource_result' as const),
        result: i < 2 ? ('accepted' as const) : ('partial' as const),
        counts:
          i === 0
            ? { pending: 0, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 }
            : i === 1
              ? { pending: n, removed: 0, referenceRemoved: 0, retained: 0, failed: 0 }
              : {
                  pending: n - 1,
                  removed: 0,
                  referenceRemoved: 0,
                  retained: i % 2 === 0 ? 1 : 0,
                  failed: i % 2 === 0 ? 0 : 1,
                },
      })),
    }
  }
  let low = 1,
    high = 4096
  while (low < high) {
    const mid = Math.ceil((low + high) / 2)
    if (Buffer.byteLength(JSON.stringify(build(mid))) <= limit) low = mid
    else high = mid - 1
  }
  const record = build(low),
    difference = limit - Buffer.byteLength(JSON.stringify(record))
  expect(difference).toBeGreaterThanOrEqual(0)
  expect(difference).toBeLessThan(4096)
  record.documentId += 'd'.repeat(difference)
  const valid = parsePresentationLifecycle(record)
  expect(Buffer.byteLength(JSON.stringify(valid))).toBe(limit)
  const path = join(
    f.root,
    'presentation-project-lifecycles',
    createHash('sha256').update(f.scope.projectId).digest('hex'),
    'lifecycle.json',
  )
  writeFileSync(path, JSON.stringify(valid))
  const result = await f.service(
    {
      operation: 'project_lifecycle_read',
      projectId: f.scope.projectId,
      documentId: valid.documentId,
    },
    new AbortController().signal,
  )
  expect(result).toEqual({ lifecycle: valid })
  expect(Buffer.byteLength(JSON.stringify(result))).toBe(limit + 14)
})
