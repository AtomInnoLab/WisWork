import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import { openPptx } from '@wiswork/pptx-engine'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
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
async function setup(compile = vi.fn(compilePresentationDeck)) {
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-page-production-'))
  roots.push(userDataPath)
  const service = createPresentationService({ userDataPath, compile }),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const call = async (operation: string, extra = {}, signal = new AbortController().signal) =>
    decode(await service({ operation, documentId: 'doc', projectId: deck.id, ...extra }, signal))
  expect((await call('save_plan', { expectedRevision: 0, plan })).revision).toBe(1)
  return { call, service, userDataPath, compile, deck, plan }
}
it('persists eight separate native pages, continues after one failure, and resumes only that page', async () => {
  let fail = true
  const compile = vi.fn(async (input: unknown) => {
    const deck = parsePresentationDeck(input)
    expect(deck.slides).toHaveLength(1)
    if (deck.slides[0]!.id === benchmarkPlannedDeck().slides[2]!.id && fail) {
      fail = false
      throw new Error('transient')
    }
    return compilePresentationDeck(deck)
  })
  const f = await setup(compile)
  const begin = { requestId: 'run', planRevision: 1, deck: f.deck }
  expect(await f.call('production_begin', begin)).toMatchObject({
    status: 'pending',
    total: 8,
    compiledCount: 0,
  })
  const first = await f.call('production_run', { requestId: 'run' })
  expect(first).toMatchObject({ status: 'partial', compiledCount: 7, total: 8 })
  expect(first.pages[2]).toMatchObject({ state: 'failed', attempt: 1, error: 'compile_failed' })
  expect(compile).toHaveBeenCalledTimes(8)
  const before = await f.call('production_page', { requestId: 'run', pageId: f.deck.slides[0]!.id })
  const restarted = createPresentationService({ userDataPath: f.userDataPath, compile })
  const run = {
    operation: 'production_run',
    documentId: 'doc',
    projectId: f.deck.id,
    requestId: 'run',
  }
  expect(decode(await restarted(run, new AbortController().signal))).toMatchObject({
    status: 'compiled',
    compiledCount: 8,
  })
  expect(compile).toHaveBeenCalledTimes(9)
  expect(
    await f.call('production_page', { requestId: 'run', pageId: f.deck.slides[0]!.id }),
  ).toEqual(before)
  expect(await f.call('production_begin', begin)).toMatchObject({ status: 'compiled' })
  expect(await f.call('production_run', { requestId: 'run' })).toMatchObject({ compiledCount: 8 })
  expect(compile).toHaveBeenCalledTimes(9)
  for (let i = 0; i < 8; i++) {
    const result = await f.call('production_page', {
      requestId: 'run',
      pageId: f.deck.slides[i]!.id,
    })
    expect(result.report.checks).toMatchObject({
      render: 'not_run',
      sources: 'not_verified',
      roundTrip: 'not_run',
    })
    const opened = await openPptx(Buffer.from(result.pptxBase64, 'base64'))
    expect(opened.deck.slides).toHaveLength(1)
    if (i === 5)
      expect(opened.deck.slides[0]!.elements.some((el) => el.type === 'table')).toBe(true)
    if (i === 6)
      expect(opened.deck.slides[0]!.elements.some((el) => el.type === 'chart')).toBe(true)
  }
  expect(await f.call('status')).toMatchObject({
    status: 'planned',
    production: { status: 'compiled', total: 8 },
  })
  expect(await f.call('get')).toEqual({ error: 'not_found' })
  expect(await f.call('production_status', { documentId: 'other', requestId: 'run' })).toEqual({
    error: 'document_mismatch',
  })
})
it('requires a matching saved plan and preserves an old job snapshot after plan edits', async () => {
  const f = await setup()
  expect(
    await f.call('production_begin', { requestId: 'run', planRevision: 2, deck: f.deck }),
  ).toEqual({ error: 'revision_conflict' })
  await f.call('production_begin', { requestId: 'run', planRevision: 1, deck: f.deck })
  expect(
    await f.call('production_page', { requestId: 'run', pageId: f.deck.slides[0]!.id }),
  ).toEqual({ error: 'page_not_ready' })
  const plan = structuredClone(f.plan)
  plan.brief.objective = 'New objective'
  expect((await f.call('save_plan', { expectedRevision: 1, plan })).revision).toBe(2)
  expect(
    await f.call('production_begin', { requestId: 'run', planRevision: 1, deck: f.deck }),
  ).toMatchObject({ planRevision: 1 })
  expect(
    await f.call('production_begin', { requestId: 'new', planRevision: 1, deck: f.deck }),
  ).toEqual({ error: 'revision_conflict' })
  const changed = structuredClone(f.deck)
  changed.slides[0]!.title = 'changed'
  expect(
    await f.call('production_begin', { requestId: 'run', planRevision: 1, deck: changed }),
  ).toEqual({ error: 'request_conflict' })
})
it('stops subsequent pages on cancellation and keeps earlier completed pages for resume', async () => {
  const controller = new AbortController()
  let calls = 0
  const compile = vi.fn(async (input: unknown) => {
    const result = await compilePresentationDeck(input)
    if (++calls === 2) controller.abort()
    return result
  })
  const f = await setup(compile)
  await f.call('production_begin', { requestId: 'run', planRevision: 1, deck: f.deck })
  expect(await f.call('production_run', { requestId: 'run' }, controller.signal)).toEqual({
    error: 'aborted',
  })
  const status = await f.call('production_status', { requestId: 'run' })
  expect(status).toMatchObject({ compiledCount: 1, status: 'partial' })
  expect(status.pages[1]).toMatchObject({ state: 'failed', error: 'aborted', attempt: 1 })
  expect(status.pages[2]).toMatchObject({ state: 'pending', attempt: 0 })
  expect(compile).toHaveBeenCalledTimes(2)
  expect(await f.call('production_run', { requestId: 'run' })).toMatchObject({ status: 'compiled' })
  expect(compile).toHaveBeenCalledTimes(9)
})
it('isolates an unavailable image to its referencing page and serializes duplicate runs', async () => {
  const f = await setup()
  const deck = structuredClone(f.deck)
  deck.assets = deck.assets.map((asset) => ({ id: asset.id, attachmentId: 'f'.repeat(64) }))
  const affected = deck.slides.filter((slide) => slide.elements.some((el) => el.kind === 'image'))
  expect(affected).toHaveLength(1)
  await f.call('production_begin', { requestId: 'run', planRevision: 1, deck })
  const result = await f.call('production_run', { requestId: 'run' })
  expect(result).toMatchObject({ status: 'partial', compiledCount: 7 })
  expect(result.pages.find((page: { id: string }) => page.id === affected[0]!.id)).toMatchObject({
    error: 'asset_unavailable',
  })
  await Promise.all([
    f.call('production_run', { requestId: 'run' }),
    f.call('production_run', { requestId: 'run' }),
  ])
  expect(f.compile).toHaveBeenCalledTimes(7)
})
it('downloads an independently produced page through the Taskpane and leaves whole-deck delivery untouched', async () => {
  const { createPresentationProductionSkill } =
    await import('../../office-addin/src/skills/powerpoint/presentation-production')
  const { InMemoryVfs } = await import('../../office-addin/src/skills/shared/vfs')
  const f = await setup(),
    vfs = new InMemoryVfs()
  const skill = createPresentationProductionSkill({
    available: () => true,
    assetsAvailable: () => true,
    documentId: async () => 'doc',
    lastProject: () => f.deck.id,
    rememberProject: async () => {},
    vfs,
    request: async (body, signal) =>
      new Response(Buffer.from(await f.service(body, signal ?? new AbortController().signal))),
  })
  for (const [name, input] of [
    ['start_presentation_production', { request_id: 'run', plan_revision: 1, deck: f.deck }],
    ['run_presentation_production', { project_id: f.deck.id, request_id: 'run' }],
  ] as const) {
    const result = await skill.executeTool({ id: name, name, input })
    expect(result.isError, result.output).not.toBe(true)
  }
  const artifact = await skill.executeTool({
    id: 'page',
    name: 'read_presentation_page_artifact',
    input: { project_id: f.deck.id, request_id: 'run', page_id: f.deck.slides[6]!.id },
  })
  expect(artifact.isError, artifact.output).not.toBe(true)
  const result = JSON.parse(artifact.output)
  const opened = await openPptx(vfs.readBytes(result.path))
  expect(opened.deck.slides).toHaveLength(1)
  expect(opened.deck.slides[0]!.elements.some((el) => el.type === 'chart')).toBe(true)
  expect(await f.call('get')).toEqual({ error: 'not_found' })
})
it('refreshes the project card after Agent production tools and supports its continue action', async () => {
  const { createOfficeHostRuntime } = await import('../../office-addin/src/agent/host-runtime')
  const f = await setup()
  let projectId: string | undefined
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      assetsAvailable: () => true,
      documentId: async () => 'doc',
      lastProject: () => projectId,
      rememberProject: async (value) => {
        projectId = value
      },
      request: async (body, signal) =>
        new Response(Buffer.from(await f.service(body, signal ?? new AbortController().signal))),
    },
  })
  const begun = await runtime.skill.executeTool({
    id: 'begin',
    name: 'start_presentation_production',
    input: { request_id: 'run', plan_revision: 1, deck: f.deck },
  })
  expect(begun.isError, begun.output).not.toBe(true)
  expect(runtime.presentation!.snapshot().project?.production).toMatchObject({
    status: 'pending',
    total: 8,
  })
  const run = await runtime.skill.executeTool({
    id: 'run',
    name: 'run_presentation_production',
    input: { request_id: 'run', project_id: f.deck.id },
  })
  expect(run.isError, run.output).not.toBe(true)
  expect(runtime.presentation!.snapshot().project?.production).toMatchObject({
    status: 'compiled',
    compiledCount: 8,
  })
  expect(runtime.proposals.pending()).toBeUndefined()
  const next = await runtime.skill.executeTool({
    id: 'begin-next',
    name: 'start_presentation_production',
    input: { request_id: 'next', plan_revision: 1, deck: f.deck },
  })
  expect(next.isError, next.output).not.toBe(true)
  await runtime.presentation!.runProduction('next')
  expect(runtime.presentation!.snapshot().project?.production).toMatchObject({
    status: 'compiled',
    requestId: 'next',
  })
  runtime.dispose()
})
it.each(['clearSession', 'abort'] as const)(
  'rejects a production result invalidated during project refresh by %s',
  async (mode) => {
    const { createOfficeHostRuntime } = await import('../../office-addin/src/agent/host-runtime')
    const f = await setup()
    let projectId: string | undefined
    const runtime = createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        documentId: async () => 'doc',
        lastProject: () => projectId,
        rememberProject: async (value) => {
          projectId = value
        },
        request: async (body, signal) =>
          new Response(Buffer.from(await f.service(body, signal ?? new AbortController().signal))),
      },
    })
    let release!: () => void, entered!: () => void
    const waiting = new Promise<void>((resolve) => {
      release = resolve
    })
    const started = new Promise<void>((resolve) => {
      entered = resolve
    })
    vi.spyOn(runtime.presentation!, 'refresh').mockImplementation(async () => {
      entered()
      await waiting
    })
    const controller = new AbortController()
    const result = runtime.skill.executeTool(
      {
        id: 'begin',
        name: 'start_presentation_production',
        input: { request_id: 'run', plan_revision: 1, deck: f.deck },
      },
      controller.signal,
    )
    await started
    if (mode === 'clearSession') runtime.clearSession()
    else controller.abort()
    release()
    expect(await result).toMatchObject({ output: 'cancelled', isError: true, mutated: false })
    expect(await f.call('production_status', { requestId: 'run' })).toMatchObject({
      status: 'pending',
    })
    runtime.dispose()
  },
)
