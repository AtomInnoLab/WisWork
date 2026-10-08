import { createHash } from 'node:crypto'
import { mkdtempSync, rmSync, readdirSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationLifecycleStore, PresentationStore } from '@wiswork/project-store'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import { inspectPresentationImage } from '../src/main/presentation-image'
import { createPresentationService } from '../src/main/presentation-service'
import { hasPresentationWorker } from '../src/main/presentation-jobs'
const attachmentDelay = vi.hoisted(() => ({
  hook: undefined as undefined | ((operation: unknown) => Promise<void>),
}))
vi.mock('../src/main/presentation-attachments', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/main/presentation-attachments')>()
  return {
    ...actual,
    createPresentationAttachmentService: (
      options: Parameters<typeof actual.createPresentationAttachmentService>[0],
    ) => {
      const service = actual.createPresentationAttachmentService(options)
      return async (body: Record<string, unknown>, signal: AbortSignal) => {
        const result = await service(body, signal)
        await attachmentDelay.hook?.(body.operation)
        return result
      }
    },
  }
})
const roots: string[] = []
afterEach(() => {
  attachmentDelay.hook = undefined
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const decode = (bytes: Uint8Array) => JSON.parse(Buffer.from(bytes).toString('utf8'))
function deferred() {
  let resolve!: () => void
  const promise = new Promise<void>((r) => {
    resolve = r
  })
  return { promise, resolve }
}
async function fixture(compile = vi.fn(compilePresentationDeck), all = false) {
  const root = mkdtempSync(join(tmpdir(), 'service-write-lease-'))
  roots.push(root)
  const service = createPresentationService({
      userDataPath: root,
      compile,
      normalizeImage: async (bytes) => {
        const image = inspectPresentationImage(bytes)
        return { bytes, width: image.width, height: image.height }
      },
    }),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  if (!all) {
    plan.slides = plan.slides.slice(0, 2)
    deck.slides = deck.slides.slice(0, 2)
  }
  const scope = { projectId: deck.id, documentId: 'doc' },
    store = new PresentationStore(root),
    lifecycle = new PresentationLifecycleStore(root)
  const call = async (
    operation: string,
    extra: Record<string, unknown> = {},
    signal = new AbortController().signal,
  ) => decode(await service({ ...scope, operation, ...extra }, signal))
  expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
  expect(await call('production_begin', { requestId: 'run', planRevision: 1, deck })).toMatchObject(
    { total: deck.slides.length },
  )
  const freeze = () => {
    const r = lifecycle.read(scope) ?? lifecycle.initialize(scope)
    return lifecycle.beginDeletion(scope, r.revision, {
      deletionId: 'delete',
      reason: 'user',
      resources: [{ resourceId: 'project', kind: 'project', ownership: 'project_exclusive' }],
    })
  }
  return { root, service, scope, store, lifecycle, deck, plan, call, freeze }
}
it.each([false, true])(
  'blocks foreground compiler success/failure after project freeze (%s)',
  async (reject) => {
    const entered = deferred(),
      gate = deferred()
    const compile = vi.fn(async (input: unknown) => {
      entered.resolve()
      await gate.promise
      if (reject) throw Error('compiler failed')
      return compilePresentationDeck(input)
    })
    const f = await fixture(compile),
      run = f.call('production_run', { requestId: 'run' })
    await entered.promise
    const before = f.store.production(f.scope.projectId, 'doc', 'run')
    f.freeze()
    gate.resolve()
    expect(await run).toEqual({ error: 'revision_conflict' })
    expect(f.store.production(f.scope.projectId, 'doc', 'run')).toEqual(before)
  },
)
it('fixes the lease and request before a queued project lock await', async () => {
  const entered = deferred(),
    gate = deferred()
  const compile = vi.fn(async (input: unknown) => {
    entered.resolve()
    await gate.promise
    return compilePresentationDeck(input)
  })
  const f = await fixture(compile),
    run = f.call('production_run', { requestId: 'run' })
  await entered.promise
  const request = { ...f.scope, operation: 'production_status', requestId: 'run' }
  const queued = f.service(request, new AbortController().signal)
  request.documentId = 'foreign'
  request.projectId = 'foreign'
  request.operation = 'production_run'
  const before = f.store.production(f.scope.projectId, 'doc', 'run')
  f.freeze()
  gate.resolve()
  expect(await run).toEqual({ error: 'revision_conflict' })
  expect(decode(await queued)).toEqual({ error: 'revision_conflict' })
  expect(f.store.production(f.scope.projectId, 'doc', 'run')).toEqual(before)
  expect(f.lifecycle.read({ projectId: 'foreign', documentId: 'foreign' })).toBeUndefined()
})
it('rejects a reopened tombstone without changing production or job receipts', async () => {
  const f = await fixture()
  let r = f.freeze()
  r = f.lifecycle.recordDeletionResult(f.scope, r.revision, {
    deletionId: 'delete',
    resourceId: 'project',
    status: 'removed',
  })
  f.lifecycle.finishDeletion(f.scope, r.revision, 'delete')
  const before = f.store.production(f.scope.projectId, 'doc', 'run')
  const other = createPresentationService({ userDataPath: f.root })
  expect(
    decode(
      await other(
        { ...f.scope, operation: 'production_job_status', requestId: 'run' },
        new AbortController().signal,
      ),
    ),
  ).toEqual({ error: 'project_deleted' })
  expect(f.store.production(f.scope.projectId, 'doc', 'run')).toEqual(before)
  expect(f.store.productionJob(f.scope.projectId, 'doc', 'run')).toBeUndefined()
})
it('keeps background lifetime independent of acceptance signal and pause finishes current page', async () => {
  const entered = deferred(),
    gate = deferred(),
    compile = vi.fn(async (input: unknown) => {
      entered.resolve()
      await gate.promise
      return compilePresentationDeck(input)
    })
  const f = await fixture(compile),
    client = new AbortController()
  expect(await f.call('production_job_start', { requestId: 'run' }, client.signal)).toMatchObject({
    job: { state: 'running' },
  })
  client.abort()
  await entered.promise
  expect(await f.call('production_job_pause', { requestId: 'run' })).toMatchObject({
    job: { state: 'pausing' },
  })
  gate.resolve()
  await vi.waitFor(() =>
    expect(hasPresentationWorker(`${f.root}\0${f.scope.projectId}`)).toBe(false),
  )
  expect(await f.call('production_job_status', { requestId: 'run' })).toMatchObject({
    job: { state: 'paused' },
    production: { compiledCount: 1 },
  })
  expect(compile).toHaveBeenCalledTimes(1)
})
it.each([false, true])(
  'passes a fixed lease to background compiler success/failure across service instances (%s)',
  async (reject) => {
    const entered = deferred(),
      gate = deferred()
    const compile = vi.fn(async (input: unknown) => {
      entered.resolve()
      await gate.promise
      if (reject) throw Error('compiler failed')
      return compilePresentationDeck(input)
    })
    const f = await fixture(compile),
      client = new AbortController()
    expect(await f.call('production_job_start', { requestId: 'run' }, client.signal)).toMatchObject(
      { job: { state: 'running' } },
    )
    client.abort()
    await entered.promise
    const before = f.store.production(f.scope.projectId, 'doc', 'run'),
      job = f.store.productionJob(f.scope.projectId, 'doc', 'run')
    f.freeze()
    const other = createPresentationService({ userDataPath: f.root, compile })
    expect(
      decode(
        await other(
          { ...f.scope, operation: 'production_job_status', requestId: 'run' },
          new AbortController().signal,
        ),
      ),
    ).toEqual({ error: 'project_deleting' })
    gate.resolve()
    await vi.waitFor(() =>
      expect(hasPresentationWorker(`${f.root}\0${f.scope.projectId}`)).toBe(false),
    )
    expect(new PresentationStore(f.root).production(f.scope.projectId, 'doc', 'run')).toEqual(
      before,
    )
    expect(f.store.productionJob(f.scope.projectId, 'doc', 'run')).toEqual(job)
    expect(compile).toHaveBeenCalledTimes(1)
  },
)
async function upload(
  f: { service: ReturnType<typeof createPresentationService> },
  bytes: Buffer,
  name: string,
) {
  const attachmentId = createHash('sha256').update(bytes).digest('hex')
  const call = async (operation: string, extra: Record<string, unknown>) =>
    decode(
      await f.service(
        { operation, documentId: 'doc', attachmentId, ...extra },
        new AbortController().signal,
      ),
    )
  expect(
    await call('attachment_begin', { sha256: attachmentId, name, sizeBytes: bytes.length }),
  ).not.toHaveProperty('error')
  expect(
    await call('attachment_chunk', { offset: 0, base64: bytes.toString('base64') }),
  ).not.toHaveProperty('error')
  expect(await call('attachment_finish', {})).toMatchObject({ status: 'ready' })
  return attachmentId
}
it.each([false, true])(
  'refuses asset-ready/rejected receipts after a genuine PC image read finishes late (%s)',
  async (reject) => {
    const f = await fixture(undefined, true),
      asset = f.deck.assets[0]!
    if (!('base64' in asset)) throw Error('fixture requires inline image')
    const attachmentId = await upload(f, Buffer.from(asset.base64, 'base64'), 'source.png')
    const deck = structuredClone(f.deck)
    deck.assets = deck.assets.map((v) => ({ id: v.id, attachmentId }))
    expect(
      await f.call('production_begin', { requestId: 'asset', planRevision: 1, deck }),
    ).toMatchObject({ total: 8 })
    const entered = deferred(),
      gate = deferred()
    attachmentDelay.hook = async (operation) => {
      if (operation === 'attachment_asset') {
        entered.resolve()
        await gate.promise
        if (reject) throw Error('read ack lost')
      }
    }
    const run = f.call('production_run', { requestId: 'asset' })
    await entered.promise
    const before = f.store.production(f.scope.projectId, 'doc', 'asset'),
      events = f.store.productionAssets(f.scope.projectId, 'doc', 'asset')
    f.freeze()
    gate.resolve()
    expect(await run).toEqual({ error: 'revision_conflict' })
    expect(f.store.production(f.scope.projectId, 'doc', 'asset')).toEqual(before)
    expect(f.store.productionAssets(f.scope.projectId, 'doc', 'asset')).toEqual(events)
  },
)
it.each([false, true])(
  'guards fresh and exact-retry issue append after awaited source evidence (%s)',
  async (retry) => {
    const f = await fixture()
    const plan = structuredClone(f.plan),
      deck = structuredClone(f.deck)
    const source = plan.sources[0]!,
      attachmentId = await upload(f, Buffer.from(source.excerpt), 'source.txt')
    source.uri = `attachment:${attachmentId}`
    deck.claims = presentationPlanClaims(plan)
    expect(await f.call('save_plan', { expectedRevision: 1, plan })).toMatchObject({ revision: 2 })
    expect(
      await f.call('production_begin', { requestId: 'issue', planRevision: 2, deck }),
    ).toMatchObject({ total: 2 })
    const report = await f.call('production_delivery_report', { requestId: 'issue' })
    const issue = report.pages[0].issues[0]
    const action = {
      actionId: 'action',
      issueId: issue.id,
      issueDigest: issue.digest,
      state: 'deferred',
      note: 'bounded review',
    }
    if (retry)
      expect(
        await f.call('production_record_issue_action', {
          requestId: 'issue',
          expectedRevision: 0,
          action,
        }),
      ).not.toHaveProperty('error')
    const before = f.store.issueActions(f.scope.projectId, 'doc', 'issue')
    const entered = deferred(),
      gate = deferred()
    attachmentDelay.hook = async (operation) => {
      if (operation === 'attachment_metadata') {
        entered.resolve()
        await gate.promise
      }
    }
    const run = f.call('production_record_issue_action', {
      requestId: 'issue',
      expectedRevision: 0,
      action,
    })
    await entered.promise
    f.freeze()
    gate.resolve()
    expect(await run).toEqual({ error: 'revision_conflict' })
    expect(f.store.issueActions(f.scope.projectId, 'doc', 'issue')).toEqual(before)
  },
)
it('does not establish lifecycle ownership for an unanchored production project', async () => {
  const f = await fixture()
  expect(await f.call('production_status', { projectId: 'unknown', requestId: 'run' })).toEqual({
    error: 'project_not_found',
  })
  expect(f.lifecycle.read({ projectId: 'unknown', documentId: 'doc' })).toBeUndefined()
  expect(f.store.plan('unknown', 'doc')).toBeUndefined()
})
it('proves existing production ownership when a standalone plan file is absent', async () => {
  const root = mkdtempSync(join(tmpdir(), 'service-production-only-'))
  roots.push(root)
  const store = new PresentationStore(root),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  store.beginProduction(deck.id, 'doc', 'run', deck, { revision: 1, plan })
  expect(store.plan(deck.id, 'doc')).toBeUndefined()
  const service = createPresentationService({ userDataPath: root })
  expect(
    decode(
      await service(
        { operation: 'production_status', documentId: 'doc', projectId: deck.id, requestId: 'run' },
        new AbortController().signal,
      ),
    ),
  ).toMatchObject({ total: 8 })
  expect(
    new PresentationLifecycleStore(root).read({ projectId: deck.id, documentId: 'doc' }),
  ).toBeUndefined()
  expect(
    decode(
      await service(
        {
          operation: 'production_status',
          documentId: 'doc',
          projectId: deck.id,
          requestId: 'unknown',
        },
        new AbortController().signal,
      ),
    ),
  ).toEqual({ error: 'not_found' })
})
it('refuses an awaited report body after its project is frozen without writing ordinary receipts', async () => {
  const f = await fixture(),
    plan = structuredClone(f.plan),
    deck = structuredClone(f.deck)
  const source = plan.sources[0]!,
    attachmentId = await upload(f, Buffer.from(source.excerpt), 'report-source.txt')
  source.uri = `attachment:${attachmentId}`
  deck.claims = presentationPlanClaims(plan)
  expect(await f.call('save_plan', { expectedRevision: 1, plan })).toMatchObject({ revision: 2 })
  expect(
    await f.call('production_begin', { requestId: 'report', planRevision: 2, deck }),
  ).toMatchObject({ total: 2 })
  const entered = deferred(),
    gate = deferred()
  attachmentDelay.hook = async (operation) => {
    if (operation === 'attachment_metadata') {
      entered.resolve()
      await gate.promise
    }
  }
  const before = f.store.production(f.scope.projectId, 'doc', 'report'),
    issues = f.store.issueActions(f.scope.projectId, 'doc', 'report')
  const run = f.call('production_delivery_report', { requestId: 'report' })
  await entered.promise
  f.freeze()
  gate.resolve()
  expect(await run).toEqual({ error: 'revision_conflict' })
  expect(f.store.production(f.scope.projectId, 'doc', 'report')).toEqual(before)
  expect(f.store.issueActions(f.scope.projectId, 'doc', 'report')).toEqual(issues)
})
it('keeps a genuinely read-only production request byte-identical when lifecycle is absent', async () => {
  const root = mkdtempSync(join(tmpdir(), 'service-readonly-absence-'))
  roots.push(root)
  const store = new PresentationStore(root),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  store.beginProduction(deck.id, 'doc', 'run', deck, { revision: 1, plan })
  const service = createPresentationService({ userDataPath: root })
  const files = () =>
    readdirSync(root, { recursive: true, withFileTypes: true })
      .filter((entry) => entry.isFile())
      .map((entry) => [
        join(entry.parentPath, entry.name),
        readFileSync(join(entry.parentPath, entry.name)).toString('base64'),
      ])
  const before = files()
  expect(
    decode(
      await service(
        { operation: 'production_status', projectId: deck.id, documentId: 'doc', requestId: 'run' },
        new AbortController().signal,
      ),
    ),
  ).toMatchObject({ total: 8 })
  expect(
    new PresentationLifecycleStore(root).read({ projectId: deck.id, documentId: 'doc' }),
  ).toBeUndefined()
  expect(files()).toEqual(before)
})
it('blocks an absent read lease when any lifecycle control is published during source audit', async () => {
  const root = mkdtempSync(join(tmpdir(), 'service-readonly-publication-'))
  roots.push(root)
  const store = new PresentationStore(root),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck(),
    service = createPresentationService({ userDataPath: root })
  const source = plan.sources[0]!,
    attachmentId = await upload({ service }, Buffer.from(source.excerpt), 'readonly-source.txt')
  source.uri = `attachment:${attachmentId}`
  deck.claims = presentationPlanClaims(plan)
  store.beginProduction(deck.id, 'doc', 'run', deck, { revision: 1, plan })
  const scope = { projectId: deck.id, documentId: 'doc' },
    entered = deferred(),
    gate = deferred()
  attachmentDelay.hook = async (operation) => {
    if (operation === 'attachment_metadata') {
      entered.resolve()
      await gate.promise
    }
  }
  const before = store.production(deck.id, 'doc', 'run')
  const run = service(
    { ...scope, operation: 'production_delivery_report', requestId: 'run' },
    new AbortController().signal,
  )
  await entered.promise
  new PresentationLifecycleStore(root).initialize(scope)
  gate.resolve()
  expect(decode(await run)).toEqual({ error: 'revision_conflict' })
  expect(store.production(deck.id, 'doc', 'run')).toEqual(before)
  expect(store.issueActions(deck.id, 'doc', 'run').actions).toEqual([])
})
