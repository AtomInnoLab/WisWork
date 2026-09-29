import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync, readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { PresentationStore, PresentationLifecycleStore } from '@wiswork/project-store'
import { stopPresentationProjectWork } from '../src/main/presentation-project-work'
import { PresentationTeamStore } from '../src/main/presentation-team'
import { createPresentationService } from '../src/main/presentation-service'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { researchDraft } from '../../../packages/project-store/tests/fixtures/presentation-research'
import {
  deliveryBundleFixture,
  cleanupDeliveryBundleFixtures,
} from './helpers/delivery-bundle-fixture'
const collision = vi.hoisted(() => ({ active: false, path: '' }))
vi.mock('node:fs', async (original) => {
  const actual = await original<typeof import('node:fs')>()
  return {
    ...actual,
    writeFileSync: (...args: Parameters<typeof actual.writeFileSync>) => {
      if (
        collision.active &&
        typeof args[0] === 'string' &&
        args[0].includes('/presentation-teams/') &&
        args[0].endsWith('.tmp')
      ) {
        collision.active = false
        collision.path = args[0]
        actual.writeFileSync(args[0], 'foreign', { flag: 'wx' })
      }
      return actual.writeFileSync(...args)
    },
  }
})
const roots: string[] = []
afterEach(() => {
  roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true }))
  cleanupDeliveryBundleFixtures()
})
const context = { version: 1 as const, actorSubject: 'a'.repeat(64), pcSubject: 'a'.repeat(64) }
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'service-factories-'))
  roots.push(root)
  const store = new PresentationStore(root),
    lifecycle = new PresentationLifecycleStore(root),
    plan = benchmarkPlan()
  const scope = { documentId: 'doc', projectId: plan.projectId }
  const service = createPresentationService({ userDataPath: root })
  const call = async (body: Record<string, unknown>) =>
    JSON.parse(Buffer.from(await service(body, new AbortController().signal, context)).toString())
  const freeze = () => {
    const record = lifecycle.read(scope) ?? lifecycle.initialize(scope)
    lifecycle.beginDeletion(scope, record.revision, {
      deletionId: 'delete',
      reason: 'user',
      resources: [{ resourceId: 'project', kind: 'project', ownership: 'project_exclusive' }],
    })
  }
  return { root, store, lifecycle, plan, scope, call, freeze }
}
it.each(['research', 'team', 'page'] as const)(
  'actual %s factory rejects a deleting project',
  async (kind) => {
    const f = fixture()
    f.store.savePlan(f.scope.projectId, 'doc', 0, f.plan)
    f.store.beginProduction(f.scope.projectId, 'doc', 'req', benchmarkPlannedDeck(), {
      revision: 1,
      plan: f.plan,
    })
    f.freeze()
    const body =
      kind === 'research'
        ? {
            operation: 'research_build',
            ...f.scope,
            ledgerId: 'ledger',
            expectedRevision: 0,
            draft: researchDraft(),
          }
        : kind === 'team'
          ? {
              operation: 'team_project_create',
              ...f.scope,
              planRevision: 1,
              expectedIdentity: context,
            }
          : { operation: 'page_backup_status', ...f.scope, backupId: 'missing' }
    expect(await f.call(body)).toEqual({ error: 'project_deleting' })
  },
)
it('actual delivery factory rejects a deleting project before beginning a bundle', async () => {
  const f = await deliveryBundleFixture(),
    lifecycle = new PresentationLifecycleStore(f.root),
    scope = { documentId: f.base.documentId, projectId: f.base.projectId }
  const record = lifecycle.initialize(scope)
  lifecycle.beginDeletion(scope, record.revision, {
    deletionId: 'delete',
    reason: 'user',
    resources: [{ resourceId: 'bundle', kind: 'delivery_bundles', ownership: 'project_exclusive' }],
  })
  const service = createPresentationService({ userDataPath: f.root })
  const result = JSON.parse(
    Buffer.from(
      await service(
        {
          ...f.base,
          operation: 'delivery_bundle_begin',
          sizeBytes: f.raw.length,
          sha256: f.base.bundleId,
          manifest: f.manifest,
        },
        new AbortController().signal,
      ),
    ).toString(),
  )
  expect(result).toEqual({ error: 'project_deleting' })
})
it('empty research reads retain absent legacy control and content', async () => {
  const f = fixture()
  const before = readdirSync(f.root, { recursive: true })
  expect(await f.call({ operation: 'research_list', ...f.scope })).not.toHaveProperty('error')
  expect(f.lifecycle.read(f.scope)).toBeUndefined()
  expect(readdirSync(f.root, { recursive: true })).toEqual(before)
})
it('legal pre-plan research creates only bound control and research content', async () => {
  const f = fixture()
  expect(
    await f.call({
      operation: 'research_build',
      ...f.scope,
      ledgerId: 'ledger',
      expectedRevision: 0,
      draft: researchDraft(),
    }),
  ).not.toHaveProperty('error')
  expect(f.lifecycle.read(f.scope)).toMatchObject({
    ...f.scope,
    state: 'active',
    policy: { contentRetentionDays: null, auditRetentionDays: null },
  })
  expect(f.store.projectScope(f.scope.projectId, 'doc')).toBeUndefined()
})

async function holdMainLock(root: string) {
  let entered!: () => void, resume!: () => void
  const arrived = new Promise<void>((resolve) => {
      entered = resolve
    }),
    gate = new Promise<void>((resolve) => {
      resume = resolve
    })
  const service = createPresentationService({
    userDataPath: root,
    compile: async () => {
      entered()
      await gate
      throw Error('compile_failed')
    },
  })
  const pending = service(
    {
      operation: 'compile',
      documentId: 'doc',
      projectId: benchmarkPlan().projectId,
      requestId: 'hold',
      ...(new PresentationStore(root).plan(benchmarkPlan().projectId, 'doc')
        ? { planRevision: 1 }
        : {}),
      deck: benchmarkPlannedDeck(),
    },
    new AbortController().signal,
  )
  await arrived
  return async () => {
    resume()
    await pending
  }
}
async function waitQueued() {
  await new Promise<void>((resolve) => setImmediate(resolve))
}
it.each(['research', 'team'] as const)(
  'captures %s revision before waiting on the actual main lock',
  async (kind) => {
    const f = fixture()
    f.store.savePlan(f.scope.projectId, 'doc', 0, f.plan)
    const release = await holdMainLock(f.root)
    const body =
      kind === 'research'
        ? {
            operation: 'research_build',
            ...f.scope,
            ledgerId: 'ledger',
            expectedRevision: 0,
            draft: researchDraft(),
          }
        : {
            operation: 'team_project_create',
            ...f.scope,
            planRevision: 1,
            expectedIdentity: context,
          }
    const pending = f.call(body)
    await waitQueued()
    f.freeze()
    await release()
    expect(await pending).toEqual({ error: 'revision_conflict' })
    expect(readdirSync(f.root)).not.toContain(
      kind === 'research' ? 'presentation-research' : 'presentation-teams',
    )
  },
)
it('team work drain waits for the actual lock waiter finally and does not fake interruption', async () => {
  const f = fixture()
  f.store.savePlan(f.scope.projectId, 'doc', 0, f.plan)
  const release = await holdMainLock(f.root)
  const pending = f.call({
    operation: 'team_project_create',
    ...f.scope,
    planRevision: 1,
    expectedIdentity: context,
  })
  await waitQueued()
  let settled = false
  const drained = stopPresentationProjectWork({ root: f.root, ...f.scope }).then(() => {
    settled = true
  })
  await new Promise<void>((resolve) => setImmediate(resolve))
  expect(settled).toBe(false)
  await release()
  expect(await pending).toEqual({ error: 'aborted' })
  await drained
  expect(settled).toBe(true)
})
it('delivery captures before the common project lock and rejects a later freeze', async () => {
  const f = await deliveryBundleFixture(),
    scope = { documentId: f.base.documentId, projectId: f.base.projectId },
    lifecycle = new PresentationLifecycleStore(f.root)
  const release = await holdMainLock(f.root),
    service = createPresentationService({ userDataPath: f.root })
  const pending = service(
    {
      ...f.base,
      operation: 'delivery_bundle_begin',
      sizeBytes: f.raw.length,
      sha256: f.base.bundleId,
      manifest: f.manifest,
    },
    new AbortController().signal,
  )
  await waitQueued()
  const record = lifecycle.read(scope)!
  expect(record).toBeDefined()
  lifecycle.beginDeletion(scope, record.revision, {
    deletionId: 'delete',
    reason: 'user',
    resources: [{ resourceId: 'bundle', kind: 'delivery_bundles', ownership: 'project_exclusive' }],
  })
  await release()
  expect(JSON.parse(Buffer.from(await pending).toString())).toEqual({ error: 'revision_conflict' })
  expect(readdirSync(f.root)).not.toContain('presentation-delivery-bundles')
})

it('team write retains a foreign temporary leaf when exclusive creation fails', async () => {
  const f = fixture()
  f.store.savePlan(f.scope.projectId, 'doc', 0, f.plan)
  const result = await f.call({
    operation: 'team_project_create',
    ...f.scope,
    planRevision: 1,
    expectedIdentity: context,
  })
  expect(result.team).toBeDefined()
  collision.active = true
  const store = new PresentationTeamStore(f.root)
  expect(() => store.write({ ...result.team, revision: 2 }, result.team)).toThrow()
  expect(readFileSync(collision.path, 'utf8')).toBe('foreign')
})
