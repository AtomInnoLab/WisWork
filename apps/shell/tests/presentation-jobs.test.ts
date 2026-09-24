import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationStore } from '@wiswork/project-store'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const decode = (bytes: Uint8Array) => JSON.parse(Buffer.from(bytes).toString('utf8'))
function deferred() {
  let resolve!: () => void
  const promise = new Promise<void>((done) => {
    resolve = done
  })
  return { promise, resolve }
}
async function setup(compile = vi.fn(compilePresentationDeck)) {
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-jobs-'))
  roots.push(userDataPath)
  const service = createPresentationService({ userDataPath, compile })
  const deck = benchmarkPlannedDeck(),
    plan = benchmarkPlan()
  deck.slides = deck.slides.slice(0, 2)
  plan.slides = plan.slides.slice(0, 2)
  const call = async (operation: string, extra = {}, signal = new AbortController().signal) =>
    decode(
      await service(
        {
          operation,
          documentId: 'doc',
          projectId: deck.id,
          ...(operation.startsWith('production_') ? { requestId: 'run' } : {}),
          ...extra,
        },
        signal,
      ),
    )
  expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
  expect(await call('production_begin', { planRevision: 1, deck })).toMatchObject({ total: 2 })
  return { call, userDataPath, compile, deck, plan }
}
async function settled(f: Awaited<ReturnType<typeof setup>>, state: string) {
  await vi.waitFor(
    async () => expect((await f.call('production_job_status')).job?.state).toBe(state),
    { timeout: 10000, interval: 100 },
  )
  return f.call('production_job_status')
}
it('accepts immediately, outlives client abort, shares workers across services and rejects foreground overlap', async () => {
  const gate = deferred(),
    entered = deferred()
  const compile = vi.fn(async (input: unknown) => {
    entered.resolve()
    await gate.promise
    return compilePresentationDeck(input)
  })
  const f = await setup(compile),
    controller = new AbortController()
  expect(await f.call('production_job_status')).toMatchObject({
    job: null,
    production: { total: 2 },
  })
  expect(await f.call('production_job_start', {}, controller.signal)).toMatchObject({
    job: { state: 'running' },
  })
  controller.abort()
  await entered.promise
  const other = createPresentationService({ userDataPath: f.userDataPath, compile })
  expect(
    decode(
      await other(
        {
          operation: 'production_job_status',
          documentId: 'doc',
          projectId: f.deck.id,
          requestId: 'run',
        },
        new AbortController().signal,
      ),
    ),
  ).toMatchObject({ job: { state: 'running' } })
  expect(await f.call('production_job_start')).toMatchObject({ job: { state: 'running' } })
  expect(await f.call('production_job_resume')).toMatchObject({ job: { state: 'running' } })
  expect(await f.call('production_run')).toEqual({ error: 'busy' })
  expect(await f.call('production_job_pause', { unexpected: true })).toEqual({
    error: 'invalid_request',
  })
  expect(await f.call('production_job_status', { documentId: 'other' })).toEqual({
    error: 'document_mismatch',
  })
  gate.resolve()
  const done = await settled(f, 'completed')
  expect(done.production.compiledCount).toBe(2)
  expect(compile).toHaveBeenCalledTimes(2)
  expect(done.job.events.filter((e: { type: string }) => e.type === 'page.compiled')).toHaveLength(
    2,
  )
})
it('pauses and cancels at page boundaries, preserving the in-flight page and successful receipts on resume', async () => {
  const gates = [deferred(), deferred()],
    entered = [deferred(), deferred()]
  const compile = vi.fn(async (input: unknown) => {
    const i = compile.mock.calls.length - 1
    if (i < 2) {
      entered[i]!.resolve()
      await gates[i]!.promise
    }
    return compilePresentationDeck(input)
  })
  const f = await setup(compile)
  expect(await f.call('production_job_start')).toMatchObject({ job: { state: 'running' } })
  await entered[0]!.promise
  expect(await f.call('production_job_pause')).toMatchObject({
    job: { state: 'pausing' },
    production: { compiledCount: 0 },
  })
  gates[0]!.resolve()
  expect((await settled(f, 'paused')).production.compiledCount).toBe(1)
  expect(compile).toHaveBeenCalledTimes(1)
  await f.call('production_job_resume')
  await entered[1]!.promise
  expect(await f.call('production_job_cancel')).toMatchObject({ job: { state: 'cancelling' } })
  gates[1]!.resolve()
  expect((await settled(f, 'cancelled')).production.compiledCount).toBe(2)
  expect(await f.call('production_job_resume')).toMatchObject({ job: { state: 'cancelled' } })
  expect(compile).toHaveBeenCalledTimes(2)
})
it('rejects a pre-aborted admission without a durable job', async () => {
  const f = await setup(),
    controller = new AbortController()
  controller.abort()
  expect(await f.call('production_job_start', {}, controller.signal)).toEqual({ error: 'aborted' })
  expect(await f.call('production_job_status')).toMatchObject({ job: null })
})
it('recovers persisted workers without a live process and retries each failed page only once per explicit resume', async () => {
  let fail = true
  const compile = vi.fn(async (input: unknown) => {
    if (fail) throw new Error('compiler')
    return compilePresentationDeck(input)
  })
  const f = await setup(compile),
    store = new PresentationStore(f.userDataPath)
  store.appendProductionJobEvent(f.deck.id, 'doc', 'run', 0, { type: 'run.started' })
  expect(await f.call('production_job_status')).toMatchObject({ job: { state: 'interrupted' } })
  expect(compile).not.toHaveBeenCalled()
  await f.call('production_job_resume')
  expect((await settled(f, 'failed')).production.compiledCount).toBe(0)
  expect(compile).toHaveBeenCalledTimes(2)
  fail = false
  await f.call('production_job_resume')
  await settled(f, 'completed')
  expect(compile).toHaveBeenCalledTimes(4)
})
it('keeps different projects independent and refuses another request in an occupied project', async () => {
  const gate = deferred(),
    entered = deferred()
  const compile = vi.fn(async (input: unknown) => {
    entered.resolve()
    await gate.promise
    return compilePresentationDeck(input)
  })
  const f = await setup(compile)
  expect(
    await f.call('production_begin', { requestId: 'second', planRevision: 1, deck: f.deck }),
  ).toMatchObject({ total: 2 })
  expect(await f.call('production_job_start')).toMatchObject({ job: { state: 'running' } })
  await entered.promise
  expect(await f.call('production_job_start', { requestId: 'second' })).toEqual({ error: 'busy' })
  const other = { projectId: 'other-project', documentId: 'other-doc' }
  expect(
    await f.call('save_plan', {
      ...other,
      expectedRevision: 0,
      plan: { ...f.plan, projectId: other.projectId },
    }),
  ).toMatchObject({ revision: 1 })
  expect(
    await f.call('production_begin', {
      ...other,
      planRevision: 1,
      deck: { ...f.deck, id: other.projectId },
    }),
  ).toMatchObject({ total: 2 })
  expect(await f.call('production_job_start', other)).toMatchObject({ job: { state: 'running' } })
  await vi.waitFor(() => expect(compile).toHaveBeenCalledTimes(2))
  await f.call('production_job_cancel')
  await f.call('production_job_cancel', other)
  gate.resolve()
  await settled(f, 'cancelled')
  await vi.waitFor(async () =>
    expect((await f.call('production_job_status', other)).job.state).toBe('cancelled'),
  )
})
it.each(['pause', 'cancel'] as const)(
  'recovers a missing worker with a persisted %s request without compiling',
  async (action) => {
    const f = await setup(),
      store = new PresentationStore(f.userDataPath)
    store.appendProductionJobEvent(f.deck.id, 'doc', 'run', 0, { type: 'run.started' })
    store.appendProductionJobEvent(f.deck.id, 'doc', 'run', 1, {
      type: action === 'pause' ? 'run.pause_requested' : 'run.cancel_requested',
    })
    expect(await f.call('production_job_status')).toMatchObject({
      job: { state: action === 'pause' ? 'paused' : 'cancelled' },
    })
    expect(f.compile).not.toHaveBeenCalled()
  },
)
it('coordinates queued admission with the foreground lock and rechecks abort before writing', async () => {
  const gate = deferred(),
    entered = deferred()
  const compile = vi.fn(async (input: unknown) => {
    entered.resolve()
    await gate.promise
    return compilePresentationDeck(input)
  })
  const f = await setup(compile)
  const foreground = f.call('production_run')
  await entered.promise
  const controller = new AbortController()
  const queued = f.call('production_job_start', {}, controller.signal)
  controller.abort()
  gate.resolve()
  expect(await foreground).toMatchObject({ compiledCount: 2 })
  expect(await queued).toEqual({ error: 'aborted' })
  expect(await f.call('production_job_status')).toMatchObject({ job: null })
  expect(await f.call('production_job_start')).toMatchObject({ job: { state: 'running' } })
  await settled(f, 'completed')
  expect(compile).toHaveBeenCalledTimes(2)
})
it.each(['paused', 'interrupted', 'failed'] as const)(
  'cancels a %s job immediately without changing receipts',
  async (state) => {
    const f = await setup(),
      store = new PresentationStore(f.userDataPath)
    let job = store.appendProductionJobEvent(f.deck.id, 'doc', 'run', 0, { type: 'run.started' })
    if (state === 'paused') {
      job = store.appendProductionJobEvent(f.deck.id, 'doc', 'run', job.revision, {
        type: 'run.pause_requested',
      })
      store.appendProductionJobEvent(f.deck.id, 'doc', 'run', job.revision, { type: 'run.paused' })
    } else
      store.appendProductionJobEvent(
        f.deck.id,
        'doc',
        'run',
        job.revision,
        state === 'failed'
          ? { type: 'run.failed', error: 'compile_failed' }
          : { type: 'run.interrupted' },
      )
    expect(await f.call('production_job_cancel')).toMatchObject({
      job: { state: 'cancelled' },
      production: { compiledCount: 0 },
    })
    expect(await f.call('production_job_start')).toMatchObject({ job: { state: 'cancelled' } })
    expect(f.compile).not.toHaveBeenCalled()
  },
)
it('keeps the real saved page authoritative when its event cannot be appended', async () => {
  const f = await setup()
  const original = PresentationStore.prototype.appendProductionJobEvent
  const append = vi
    .spyOn(PresentationStore.prototype, 'appendProductionJobEvent')
    .mockImplementation(function (...args) {
      if (args[4].type === 'page.compiled') {
        append.mockRestore()
        throw new Error('disk_unavailable')
      }
      return original.apply(this, args)
    })
  try {
    expect(await f.call('production_job_start')).toMatchObject({ job: { state: 'running' } })
    const failed = await settled(f, 'failed')
    expect(failed.production.compiledCount).toBe(1)
    expect(failed.job.events.at(-1)).toMatchObject({ type: 'run.failed', error: 'invalid_state' })
    expect(f.compile).toHaveBeenCalledTimes(1)
    await f.call('production_job_resume')
    expect((await settled(f, 'completed')).production.compiledCount).toBe(2)
    expect(f.compile).toHaveBeenCalledTimes(2)
  } finally {
    append.mockRestore()
  }
})
