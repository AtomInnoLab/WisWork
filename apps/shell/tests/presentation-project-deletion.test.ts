import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync, existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { tmpdir } from 'node:os'
import { join, dirname } from 'node:path'
import { PresentationStore, PresentationLifecycleStore } from '@wiswork/project-store'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { createPresentationService } from '../src/main/presentation-service'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { registerPresentationProjectWork } from '../src/main/presentation-project-work'
import { createPresentationProjectDeletionService } from '../src/main/presentation-project-deletion'
const faults = vi.hoisted(() => ({ unlink: false, receipt: false, armReceipt: false }))
vi.mock('node:fs', async (load) => {
  const actual = await load<typeof import('node:fs')>()
  return {
    ...actual,
    unlinkSync: (...args: Parameters<typeof actual.unlinkSync>) => {
      if (faults.unlink && String(args[0]).includes('presentation-project-deletion-work')) {
        faults.unlink = false
        throw Error('synthetic-private-io')
      }
      if (faults.armReceipt && String(args[0]).endsWith('proof.json')) {
        faults.armReceipt = false
        faults.receipt = true
      }
      return actual.unlinkSync(...args)
    },
    renameSync: (...args: Parameters<typeof actual.renameSync>) => {
      if (faults.receipt && String(args[1]).endsWith('lifecycle.json')) {
        faults.receipt = false
        throw Error('synthetic-receipt-failure')
      }
      return actual.renameSync(...args)
    },
  }
})
const roots: string[] = []
afterEach(() => {
  faults.unlink = false
  faults.receipt = false
  faults.armReceipt = false
  for (const r of roots.splice(0)) rmSync(r, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'deletion-executor-'))
  roots.push(root)
  const plan = benchmarkPlan(),
    scope = { projectId: plan.projectId, documentId: 'doc' }
  new PresentationStore(root).savePlan(scope.projectId, scope.documentId, 0, plan)
  const life = new PresentationLifecycleStore(root)
  let held = false
  let tail = Promise.resolve()
  const lock = async () => {
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
  const make = () =>
    createPresentationProjectDeletionService({
      userDataPath: root,
      acquireProjectLock: lock,
    })
  return { root, scope, life, make, lock, held: () => held }
}
it('preview does not initialize controls; actual confirm deletes content and retains tombstone', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope)
  expect(f.life.read(f.scope)).toBeUndefined()
  const r = await s.confirm({
    scope: f.scope,
    expectedRevision: p.expectedRevision,
    confirmationToken: p.confirmationToken,
    deletionId: 'delete',
  })
  expect(r.state).toBe('deleted')
  expect(f.life.read(f.scope)?.state).toBe('deleted')
  expect(
    new PresentationStore(f.root).projectScope(f.scope.projectId, f.scope.documentId),
  ).toBeUndefined()
  expect(f.held()).toBe(false)
})
it('freeze is durable and drain waits without acquiring project lock', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope),
    work = registerPresentationProjectWork({ scope: { root: f.root, ...f.scope } })
  let done = false
  const pending = s
    .confirm({
      scope: f.scope,
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'delete',
    })
    .then((r) => {
      done = true
      return r
    })
  while (!work.signal.aborted) await Promise.resolve()
  expect(f.life.read(f.scope)?.state).toBe('deleting')
  expect(f.held()).toBe(false)
  expect(done).toBe(false)
  work.finish()
  expect((await pending).state).toBe('deleted')
})
it('rejects cross-scope token and stale revision before freezing', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope)
  await expect(
    s.confirm({
      scope: { ...f.scope, documentId: 'other' },
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'd',
    }),
  ).rejects.toThrow()
  const l = f.life.initialize(f.scope)
  f.life.setPolicy(f.scope, l.revision, { contentRetentionDays: 2, auditRetentionDays: null })
  await expect(
    s.confirm({
      scope: f.scope,
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'd',
    }),
  ).rejects.toThrow('revision_conflict')
  expect(f.life.read(f.scope)?.state).toBe('active')
})
it('new unbound staging after freeze is retained and prevents deleted closure', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope),
    w = registerPresentationProjectWork({ scope: { root: f.root, ...f.scope } })
  const pending = s.confirm({
    scope: f.scope,
    expectedRevision: p.expectedRevision,
    confirmationToken: p.confirmationToken,
    deletionId: 'd',
  })
  while (!w.signal.aborted) await Promise.resolve()
  const dir = join(f.root, 'presentation-page-backups', '.tmp-11111111-1111-4111-8111-111111111111')
  mkdirSync(dir, { recursive: true })
  writeFileSync(join(dir, 'unknown'), 'keep')
  w.finish()
  const r = await pending
  expect(r.state).toBe('partial')
  expect(existsSync(dir)).toBe(true)
  expect(f.life.read(f.scope)?.state).toBe('deleting')
})
it('duplicate confirmation does not revive or repeat completed deletion', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope),
    body = {
      scope: f.scope,
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'd',
    }
  expect((await s.confirm(body)).state).toBe('deleted')
  await expect(f.make().confirm(body)).rejects.toThrow()
  expect(f.life.read(f.scope)?.state).toBe('deleted')
})
it('queued second executor cannot adopt first executor receipts', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope),
    w = registerPresentationProjectWork({ scope: { root: f.root, ...f.scope } })
  const a = s.confirm({
    scope: f.scope,
    expectedRevision: p.expectedRevision,
    confirmationToken: p.confirmationToken,
    deletionId: 'd',
  })
  while (!w.signal.aborted) await Promise.resolve()
  const rev = f.life.read(f.scope)!.revision
  const b = f.make().resume({ scope: f.scope, expectedRevision: rev, deletionId: 'd' })
  w.finish()
  const outcomes = await Promise.allSettled([a, b])
  expect(outcomes.filter((r) => r.status === 'fulfilled')).toHaveLength(1)
  expect(f.life.read(f.scope)?.state).toBe('deleted')
})
it('missing scope preview creates no directories', async () => {
  const f = fixture()
  await expect(f.make().preview({ projectId: 'missing', documentId: 'doc' })).rejects.toThrow(
    'project_not_found',
  )
  expect(existsSync(join(f.root, 'presentation-project-lifecycles'))).toBe(false)
})

it('reopens failed resource receipt and completes deterministic quarantine cleanup', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope)
  faults.unlink = true
  const first = await s.confirm({
    scope: f.scope,
    expectedRevision: p.expectedRevision,
    confirmationToken: p.confirmationToken,
    deletionId: 'd',
  })
  expect(first.state).toBe('partial')
  expect(first.counts.failed).toBe(1)
  expect(
    (await f.make().resume({ scope: f.scope, expectedRevision: first.revision, deletionId: 'd' }))
      .state,
  ).toBe('deleted')
})
it('confirmation owns caller body before awaiting project lock', async () => {
  const f = fixture(),
    p = await f.make().preview(f.scope)
  let go!: () => void
  const wait = new Promise<void>((r) => {
    go = r
  })
  const s = createPresentationProjectDeletionService({
    userDataPath: f.root,
    acquireProjectLock: async () => {
      await wait
      return () => {}
    },
  })
  const body = {
      scope: { ...f.scope },
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'd',
    },
    pending = s.confirm(body)
  body.scope.documentId = 'foreign'
  body.deletionId = 'foreign'
  body.confirmationToken = '0'.repeat(64)
  go()
  expect((await pending).state).toBe('deleted')
  expect(f.life.read(f.scope)?.deletion?.deletionId).toBe('d')
})
it('queued null revision cannot initialize over newly created controls', async () => {
  const f = fixture(),
    p = await f.make().preview(f.scope)
  let go!: () => void
  const wait = new Promise<void>((r) => {
    go = r
  })
  const s = createPresentationProjectDeletionService({
    userDataPath: f.root,
    acquireProjectLock: async () => {
      await wait
      return () => {}
    },
  })
  const pending = s.confirm({
    scope: f.scope,
    expectedRevision: null,
    confirmationToken: p.confirmationToken,
    deletionId: 'd',
  })
  f.life.initialize(f.scope)
  go()
  await expect(pending).rejects.toThrow('revision_conflict')
  expect(f.life.read(f.scope)?.state).toBe('active')
})
it('new known exclusive namespace invalidates preview token before freeze', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope)
  new PresentationStore(f.root).savePlan(f.scope.projectId, f.scope.documentId, 1, {
    ...benchmarkPlan(),
    title: 'changed',
  })
  await expect(
    s.confirm({
      scope: f.scope,
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'd',
    }),
  ).rejects.toThrow('confirmation_conflict')
  expect(f.life.read(f.scope)).toBeUndefined()
})

it('receipt publish failure after actual removal resumes the same durable intent without project metadata', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope)
  faults.armReceipt = true
  await expect(
    s.confirm({
      scope: f.scope,
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'd',
    }),
  ).rejects.toThrow('invalid_state')
  const r = f.life.read(f.scope)!
  expect(r.state).toBe('deleting')
  expect(
    new PresentationStore(f.root).projectScope(f.scope.projectId, f.scope.documentId),
  ).toBeUndefined()
  expect(
    (await f.make().resume({ scope: f.scope, expectedRevision: r.revision, deletionId: 'd' }))
      .state,
  ).toBe('deleted')
})
it('preserves frozen intent on client cancellation and resumes only explicit fixed revision', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope),
    w = registerPresentationProjectWork({ scope: { root: f.root, ...f.scope } }),
    controller = new AbortController()
  const pending = s.confirm(
    {
      scope: f.scope,
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'd',
    },
    controller.signal,
  )
  while (!w.signal.aborted) await Promise.resolve()
  controller.abort()
  w.finish()
  await expect(pending).rejects.toThrow('aborted')
  const r = f.life.read(f.scope)!
  expect(r.state).toBe('deleting')
  expect(
    (await f.make().resume({ scope: f.scope, expectedRevision: r.revision, deletionId: 'd' }))
      .state,
  ).toBe('deleted')
})
it('governance4096 boundary uses real stored project binding;4097 refuses without mutation', async () => {
  const f = fixture(),
    doc = 'x'.repeat(4096),
    long = { ...f.scope, documentId: doc },
    path = join(
      f.root,
      'projects',
      'presentations',
      createHash('sha256').update(f.scope.projectId).digest('hex'),
      'project.json',
    )
  // A real metadata-only governed project; compilation records retain their narrower contract.
  rmSync(dirname(path), { recursive: true, force: true })
  mkdirSync(dirname(path), { recursive: true })
  writeFileSync(path, JSON.stringify({ version: 1, ...long }))
  const s = f.make(),
    p = await s.preview(long)
  expect(f.life.read(long)).toBeUndefined()
  await expect(s.preview({ ...long, documentId: doc + 'x' })).rejects.toThrow('invalid_request')
  expect(
    (
      await s.confirm({
        scope: long,
        expectedRevision: p.expectedRevision,
        confirmationToken: p.confirmationToken,
        deletionId: 'd',
      })
    ).state,
  ).toBe('deleted')
})
it.each(['extra', 'nonstring'])(
  'rejects malformed confirmation before lock admission %s',
  async (mode) => {
    const f = fixture()
    let locks = 0
    const s = createPresentationProjectDeletionService({
      userDataPath: f.root,
      acquireProjectLock: async () => {
        locks++
        return () => {}
      },
    })
    const body = {
      scope: f.scope,
      expectedRevision: null,
      confirmationToken: 'a'.repeat(64),
      deletionId: 'd',
      ...(mode === 'extra' ? { privateBody: 'secret' } : {}),
    }
    if (mode === 'nonstring') (body as unknown as Record<string, unknown>).deletionId = 3
    await expect(s.confirm(body)).rejects.toThrow('invalid_request')
    expect(locks).toBe(0)
  },
)
it.each(['dynamic', 'throws'])(
  'normalizes hostile error message getters to finite error %s',
  async (mode) => {
    const f = fixture(),
      s = createPresentationProjectDeletionService({
        userDataPath: f.root,
        acquireProjectLock: async () => {
          const e = new Error()
          let reads = 0
          Object.defineProperty(e, 'message', {
            get() {
              if (mode === 'throws') throw Error('private-getter-secret')
              return ++reads === 1 ? 'revision_conflict' : 'private-transport-secret'
            },
          })
          throw e
        },
      }),
      p = await s.preview(f.scope)
    await expect(
      s.confirm({
        scope: f.scope,
        expectedRevision: p.expectedRevision,
        confirmationToken: p.confirmationToken,
        deletionId: 'd',
      }),
    ).rejects.toThrow(mode === 'dynamic' ? 'revision_conflict' : 'invalid_state')
  },
)

it('waiting resume preserves its old revision when another executor publishes a receipt', async () => {
  const f = fixture(),
    s = f.make(),
    p = await s.preview(f.scope),
    w = registerPresentationProjectWork({ scope: { root: f.root, ...f.scope } }),
    controller = new AbortController()
  const start = s.confirm(
    {
      scope: f.scope,
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'd',
    },
    controller.signal,
  )
  while (!w.signal.aborted) await Promise.resolve()
  controller.abort()
  w.finish()
  await expect(start).rejects.toThrow('aborted')
  const r = f.life.read(f.scope)!,
    release = await f.lock(),
    pending = f.make().resume({ scope: f.scope, expectedRevision: r.revision, deletionId: 'd' })
  await Promise.resolve()
  await Promise.resolve()
  f.life.recordDeletionResult(f.scope, r.revision, {
    deletionId: 'd',
    resourceId: r.deletion!.resources[0]!.resourceId,
    status: 'failed',
    code: 'io_failed',
  })
  release()
  await expect(pending).rejects.toThrow('revision_conflict')
  expect(new PresentationStore(f.root).projectScope(f.scope.projectId, f.scope.documentId)).toEqual(
    f.scope,
  )
})
it('actual background compiler must return before drain acquires project lock and removes content', async () => {
  const f = fixture()
  let entered!: () => void, resume!: () => void
  const arrived = new Promise<void>((r) => {
      entered = r
    }),
    wait = new Promise<void>((r) => {
      resume = r
    })
  const pc = createPresentationService({
    userDataPath: f.root,
    compile: async (...args) => {
      entered()
      await wait
      return compilePresentationDeck(...args)
    },
  })
  const call = (operation: string, body = {}) =>
    pc({ operation, ...f.scope, ...body }, new AbortController().signal)
  await call('production_begin', {
    requestId: 'job',
    planRevision: 1,
    deck: benchmarkPlannedDeck(),
  })
  await call('production_job_start', { requestId: 'job' })
  await arrived
  const s = f.make(),
    p = await s.preview(f.scope)
  let settled = false
  const pending = s
    .confirm({
      scope: f.scope,
      expectedRevision: p.expectedRevision,
      confirmationToken: p.confirmationToken,
      deletionId: 'd',
    })
    .then((r) => {
      settled = true
      return r
    })
  while (f.life.read(f.scope)?.state !== 'deleting') await Promise.resolve()
  await Promise.resolve()
  expect(f.held()).toBe(false)
  expect(settled).toBe(false)
  resume()
  expect((await pending).state).toBe('deleted')
  expect(f.held()).toBe(false)
})
