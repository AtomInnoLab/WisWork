import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync, readdirSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationStore, PresentationLifecycleStore } from '@wiswork/project-store'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { stopPresentationProjectWork } from '../src/main/presentation-project-work'
const hooks = vi.hoisted(() => ({
  research: undefined as (() => Promise<void>) | undefined,
  audit: undefined as (() => Promise<void>) | undefined,
}))
vi.mock('../src/main/presentation-research-plan-binding', async (original) => {
  const actual = await original<typeof import('../src/main/presentation-research-plan-binding')>()
  return {
    ...actual,
    readBoundPresentationResearch: async (
      ...args: Parameters<typeof actual.readBoundPresentationResearch>
    ) => {
      const result = await actual.readBoundPresentationResearch(...args)
      await hooks.research?.()
      return result
    },
  }
})
vi.mock('../src/main/presentation-source-audit', async (original) => {
  const actual = await original<typeof import('../src/main/presentation-source-audit')>()
  return {
    ...actual,
    auditPresentationSources: async (
      ...args: Parameters<typeof actual.auditPresentationSources>
    ) => {
      const result = await actual.auditPresentationSources(...args)
      await hooks.audit?.()
      return result
    },
  }
})
const roots: string[] = []
afterEach(() => {
  hooks.research = undefined
  hooks.audit = undefined
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function deferred() {
  let resolve!: () => void
  const promise = new Promise<void>((done) => {
    resolve = done
  })
  return { promise, resolve }
}
function fixture(compile = compilePresentationDeck) {
  const root = mkdtempSync(join(tmpdir(), 'service-full-fences-'))
  roots.push(root)
  const service = createPresentationService({ userDataPath: root, compile }),
    store = new PresentationStore(root),
    lifecycle = new PresentationLifecycleStore(root),
    deck = benchmarkPlannedDeck(),
    plan = benchmarkPlan()
  deck.slides = deck.slides.slice(0, 2)
  plan.slides = plan.slides.slice(0, 2)
  const scope = { projectId: deck.id, documentId: 'doc' }
  const call = async (operation: string, fields: Record<string, unknown> = {}) =>
    JSON.parse(
      Buffer.from(
        await service({ operation, ...scope, ...fields }, new AbortController().signal),
      ).toString('utf8'),
    )
  const freeze = () => {
    const record = lifecycle.read(scope) ?? lifecycle.initialize(scope)
    return lifecycle.beginDeletion(scope, record.revision, {
      deletionId: 'delete',
      reason: 'user',
      resources: [{ resourceId: 'own', kind: 'project', ownership: 'project_exclusive' }],
    })
  }
  return { root, service, store, lifecycle, deck, plan, scope, call, freeze }
}
it.each([false, true])(
  'does not complete foreground compile after frozen success/failure (%s)',
  async (reject) => {
    const entered = deferred(),
      gate = deferred(),
      f = fixture(async (input) => {
        entered.resolve()
        await gate.promise
        if (reject) throw Error('compiler failed')
        return compilePresentationDeck(input)
      })
    const run = f.call('compile', { requestId: 'run', deck: f.deck })
    await entered.promise
    const before = f.store.request(f.scope.projectId, 'doc', 'run')
    f.freeze()
    gate.resolve()
    expect(await run).toEqual({ error: 'revision_conflict' })
    expect(f.store.request(f.scope.projectId, 'doc', 'run')).toEqual(before)
  },
)
it('does not save a new plan after its research read finishes after freeze', async () => {
  const f = fixture(),
    entered = deferred(),
    gate = deferred()
  hooks.research = async () => {
    entered.resolve()
    await gate.promise
  }
  const run = f.call('save_plan', { expectedRevision: 0, plan: f.plan })
  await entered.promise
  expect(f.lifecycle.read(f.scope)?.state).toBe('active')
  expect(f.store.projectScope(f.scope.projectId, 'doc')).toBeUndefined()
  f.freeze()
  gate.resolve()
  expect(await run).toEqual({ error: 'revision_conflict' })
  expect(f.store.plan(f.scope.projectId, 'doc')).toBeUndefined()
})
it('owns nested request input before a project lock await', async () => {
  const gate = deferred(),
    entered = deferred()
  let count = 0
  const f = fixture(async (input) => {
    if (++count === 1) {
      entered.resolve()
      await gate.promise
    }
    return compilePresentationDeck(input)
  })
  const first = f.call('compile', { requestId: 'first', deck: f.deck })
  await entered.promise
  const body = {
    operation: 'compile',
    ...f.scope,
    requestId: 'second',
    deck: structuredClone(f.deck),
  }
  const second = f.service(body, new AbortController().signal)
  body.requestId = 'foreign'
  body.documentId = 'foreign'
  body.deck.id = 'foreign'
  body.deck.title = 'mutated'
  body.deck.slides[0]!.title = 'mutated'
  gate.resolve()
  expect(await first).toMatchObject({ status: 'compiled' })
  expect(JSON.parse(Buffer.from(await second).toString('utf8'))).toMatchObject({
    requestId: 'second',
    status: 'compiled',
  })
  expect(f.store.request(f.scope.projectId, 'doc', 'second')?.deck).toEqual(f.deck)
  expect(f.store.request(f.scope.projectId, 'doc', 'foreign')).toBeUndefined()
})
it('registers foreground drain and waits for actual unsignalled compiler return', async () => {
  const gate = deferred(),
    entered = deferred(),
    f = fixture(async (input) => {
      entered.resolve()
      await gate.promise
      return compilePresentationDeck(input)
    })
  const run = f.call('compile', { requestId: 'run', deck: f.deck })
  await entered.promise
  const before = f.store.request(f.scope.projectId, 'doc', 'run')
  f.freeze()
  let drained = false
  const drain = stopPresentationProjectWork({ root: f.root, ...f.scope }).then(() => {
    drained = true
  })
  await Promise.resolve()
  expect(drained).toBe(false)
  gate.resolve()
  expect(await run).toEqual({ error: 'aborted' })
  await drain
  expect(f.store.request(f.scope.projectId, 'doc', 'run')).toEqual(before)
})
it.each([false, true])(
  'does not finish source audit after frozen success/failure (%s)',
  async (reject) => {
    const f = fixture()
    expect(await f.call('save_plan', { expectedRevision: 0, plan: f.plan })).toMatchObject({
      revision: 1,
    })
    const entered = deferred(),
      gate = deferred()
    hooks.audit = async () => {
      entered.resolve()
      await gate.promise
      if (reject) throw Error('source failed')
    }
    const run = f.call('audit_sources', { auditId: 'audit' })
    await entered.promise
    const before = f.store.sourceAudits(f.scope.projectId, 'doc')
    f.freeze()
    gate.resolve()
    expect(await run).toEqual({ error: 'revision_conflict' })
    expect(f.store.sourceAudits(f.scope.projectId, 'doc')).toEqual(before)
  },
)
it('retains genuinely read-only legacy and unknown-project filesystem behavior', async () => {
  const f = fixture()
  f.store.savePlan(f.scope.projectId, 'doc', 0, f.plan)
  const files = () =>
    readdirSync(f.root, { recursive: true, withFileTypes: true })
      .filter((v) => v.isFile())
      .map((v) => [
        join(v.parentPath, v.name),
        readFileSync(join(v.parentPath, v.name)).toString('base64'),
      ])
  const before = files()
  expect(await f.call('get_plan')).toMatchObject({ revision: 1 })
  expect(f.lifecycle.read(f.scope)).toBeUndefined()
  expect(await f.call('get_plan', { projectId: 'unknown' })).toEqual({ error: 'not_found' })
  expect(await f.call('get_plan', { projectId: 'unknown', revision: 1 })).toEqual({
    error: 'plan_revision_unavailable',
  })
  expect(files()).toEqual(before)
})
it('cannot revive a deleted project through creation or another document', async () => {
  const f = fixture()
  let r = f.freeze()
  r = f.lifecycle.recordDeletionResult(f.scope, r.revision, {
    deletionId: 'delete',
    resourceId: 'own',
    status: 'removed',
  })
  f.lifecycle.finishDeletion(f.scope, r.revision, 'delete')
  expect(await f.call('save_plan', { expectedRevision: 0, plan: f.plan })).toEqual({
    error: 'project_deleted',
  })
  expect(await f.call('compile', { requestId: 'run', deck: f.deck })).toEqual({
    error: 'project_deleted',
  })
  expect(
    await f.call('compile', { requestId: 'run', deck: f.deck, documentId: 'foreign' }),
  ).toEqual({ error: 'document_mismatch' })
  expect(f.store.projectScope(f.scope.projectId, 'doc')).toBeUndefined()
})

it('guards resumed historical receipts with no late completion', async () => {
  const gate = deferred(),
    entered = deferred(),
    f = fixture(async (input) => {
      entered.resolve()
      await gate.promise
      return compilePresentationDeck(input)
    })
  f.store.begin(f.scope.projectId, 'doc', 'run', f.deck)
  const run = f.call('resume', { requestId: 'run' })
  await entered.promise
  const before = f.store.request(f.scope.projectId, 'doc', 'run')
  f.freeze()
  gate.resolve()
  expect(await run).toEqual({ error: 'revision_conflict' })
  expect(f.store.request(f.scope.projectId, 'doc', 'run')).toEqual(before)
})
it('rejects invalid first-creation CAS before saving control or content', async () => {
  const f = fixture()
  expect(await f.call('save_plan', { expectedRevision: 1, plan: f.plan })).toEqual({
    error: 'revision_conflict',
  })
  expect(await f.call('compile', { requestId: 'run', deck: f.deck, planRevision: 1 })).toEqual({
    error: 'revision_conflict',
  })
  expect(readdirSync(f.root)).toEqual([])
})
it('blocks ordinary acceptance and page locks on a frozen existing plan', async () => {
  const f = fixture()
  expect(await f.call('save_plan', { expectedRevision: 0, plan: f.plan })).toMatchObject({
    revision: 1,
  })
  const before = f.store.plan(f.scope.projectId, 'doc')!
  f.freeze()
  expect(
    await f.call('accept_plan', {
      decisionId: 'decision',
      expectedRevision: 1,
      planDigest: before.inputDigest,
    }),
  ).toEqual({ error: 'project_deleting' })
  expect(
    await f.call('set_plan_page_lock', {
      expectedRevision: 1,
      pageId: f.plan.slides[0]!.id,
      locked: true,
    }),
  ).toEqual({ error: 'project_deleting' })
  expect(f.store.plan(f.scope.projectId, 'doc')).toEqual(before)
  expect(f.store.planAcceptances(f.scope.projectId, 'doc').records).toHaveLength(0)
})

it('keeps unknown existing-only writes absent and returns the original not_found result', async () => {
  const f = fixture()
  for (const [operation, fields] of [
    ['resume', { requestId: 'run' }],
    ['audit_sources', { auditId: 'audit' }],
    ['accept_plan', { decisionId: 'decision', expectedRevision: 1, planDigest: 'a'.repeat(64) }],
    ['set_plan_page_lock', { expectedRevision: 1, pageId: 'page', locked: true }],
  ] as const)
    expect(await f.call(operation, fields)).toEqual({ error: 'not_found' })
  expect(readdirSync(f.root)).toEqual([])
})
