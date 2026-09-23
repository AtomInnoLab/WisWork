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

it('prepares real page files and resumes confirmed Office import without mixing whole-deck receipts', async () => {
  const { createOfficeHostRuntime } = await import('../../office-addin/src/agent/host-runtime')
  const { createPresentationDocumentBinding } =
    await import('../../office-addin/src/skills/powerpoint/presentation-document')
  const f = await setup()
  await f.call('production_begin', { requestId: 'run', planRevision: 1, deck: f.deck })
  await f.call('production_run', { requestId: 'run' })
  await f.call('compile', { requestId: 'whole', planRevision: 1, deck: f.deck })
  const { BrowserPowerPointAdapter } =
    await import('../../office-addin/src/skills/powerpoint/browser-powerpoint-adapter')
  const inspect = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'inspectPresentationPage')
    .mockImplementation(async (slideId) => ({
      slideId,
      slideWidth: 960,
      slideHeight: 540,
      shapes: [],
      shapesTruncated: false,
      overflows: [],
      overlaps: [],
      overlapsTruncated: false,
      screenshot: {
        mime: 'image/png',
        base64:
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl6LPsAAAAASUVORK5CYII=',
      },
    }))
  let pageText = 'before'
  const readText = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'readPresentationPageText')
    .mockImplementation(async (slideId, shapeId) => ({
      slideId,
      shapeId,
      text: pageText,
      paragraphs: [pageText],
    }))
  const editText = vi
    .spyOn(BrowserPowerPointAdapter.prototype, 'editPresentationPageText')
    .mockImplementation(async (_slide, _shape, text) => {
      pageText = text
    })
  const qaRecords = new Map<
    string,
    import('../../office-addin/src/skills/powerpoint/presentation-qa').PresentationQaRecord
  >()
  const settings = new Map<string, unknown>()
  const binding = createPresentationDocumentBinding(
    {
      get: (key) => settings.get(key),
      set: (key, value) => {
        settings.set(key, value)
      },
      save: async () => {},
      location: () => 'location',
    },
    () => 'doc',
  )
  const hostIds = ['original']
  const slides = { items: [] as { id: string }[], load: () => {} }
  const bytes: string[] = []
  let queued = false,
    stopAfterThree = true
  let runtime: ReturnType<typeof createOfficeHostRuntime>
  const insert = vi.fn((base64: string, options: { sourceSlideIds: string[] }) => {
    expect(options.sourceSlideIds).toEqual(['256#'])
    bytes.push(base64)
    queued = true
  })
  const context = {
    presentation: { slides, insertSlidesFromBase64: insert },
    sync: async () => {
      if (queued) {
        queued = false
        hostIds.push(`host-${hostIds.length}`)
        if (hostIds.length === 4 && stopAfterThree) runtime.proposals.newTurn()
      }
      slides.items = hostIds.map((id) => ({ id }))
    },
  }
  vi.stubGlobal('Office', {
    context: {
      host: 'PowerPoint',
      requirements: {
        isSetSupported: (_name: string, version: string) => ['1.2', '1.10'].includes(version),
      },
    },
  })
  vi.stubGlobal('PowerPoint', {
    run: async (action: (ctx: typeof context) => Promise<unknown>) => action(context),
  })
  const create = () =>
    createOfficeHostRuntime('powerpoint', {
      presentation: {
        ...binding,
        documentId: async () => 'doc',
        readQa: (key) => qaRecords.get(key),
        writeQa: async (key, value) => {
          qaRecords.set(key, structuredClone(value))
        },
        invalidateQa: async (ids) => {
          for (const record of qaRecords.values())
            for (const page of record.pages)
              if (!ids || ids.includes(page.hostSlideId)) page.recheckRequired = true
        },
        available: () => true,
        assetsAvailable: () => true,
        request: async (body, signal) =>
          new Response(Buffer.from(await f.service(body, signal ?? new AbortController().signal))),
      },
    })
  runtime = create()
  const prepare = {
    id: 'prepare',
    name: 'prepare_presentation_production_import',
    input: { project_id: f.deck.id, request_id: 'run' },
  }
  const importCall = {
    id: 'import',
    name: 'import_presentation_production',
    input: { project_id: f.deck.id },
  }
  try {
    const prepared = await runtime.skill.executeTool(prepare)
    expect(prepared.isError, prepared.output).not.toBe(true)
    expect(runtime.importProgress!.read()).toMatchObject({ total: 8, completed: 0 })
    expect((await runtime.skill.executeTool(importCall)).isError).not.toBe(true)
    await expect(runtime.proposals.confirm(runtime.proposals.pending()!.id)).rejects.toThrow(
      'cancelled',
    )
    expect(hostIds).toHaveLength(4)
    expect(binding.readReceipt(`${f.deck.id}/run`)).toBeUndefined()
    expect(binding.readReceipt(`production/${f.deck.id}/run`)?.checkpoint).toMatchObject({
      version: 2,
      pageIds: f.deck.slides.map((p) => p.id),
      completed: [{ slideId: 'host-1' }, { slideId: 'host-2' }, { slideId: 'host-3' }],
    })
    runtime.dispose()
    stopAfterThree = false
    runtime = create()
    const restored = await runtime.skill.executeTool({
      id: 'restore',
      name: 'restore_presentation_project',
      input: { project_id: f.deck.id },
    })
    expect(restored.isError, restored.output).not.toBe(true)
    expect((await runtime.skill.executeTool(prepare)).isError).not.toBe(true)
    expect(runtime.importProgress!.read()).toMatchObject({ completed: 3, status: 'partial' })
    expect((await runtime.skill.executeTool(importCall)).isError).not.toBe(true)
    await runtime.proposals.confirm(runtime.proposals.pending()!.id)
    expect(hostIds).toHaveLength(9)
    expect(insert).toHaveBeenCalledTimes(8)
    expect(runtime.importProgress!.read()).toMatchObject({ completed: 8, status: 'complete' })
    expect((await runtime.skill.executeTool(importCall)).output).toContain('already_imported')
    expect(insert).toHaveBeenCalledTimes(8)
    const qa = await runtime.skill.executeTool({
      id: 'qa',
      name: 'read_presentation_qa',
      input: { project_id: f.deck.id },
    })
    expect(qa.isError, qa.output).not.toBe(true)
    for (const page of f.deck.slides.slice(0, 2)) {
      const captured = await runtime.skill.executeTool({
        id: 'capture',
        name: 'capture_presentation_page_qa',
        input: { project_id: f.deck.id, page_id: page.id },
      })
      expect(captured.isError, captured.output).not.toBe(true)
    }
    expect(inspect).toHaveBeenLastCalledWith('host-2', undefined)
    expect(runtime.qa!.read()).toMatchObject({
      source: 'production',
      pages: [{ hostSlideId: 'host-1' }, { hostSlideId: 'host-2' }],
    })
    expect(qaRecords.has(`${f.deck.id}/run`)).toBe(false)
    const edit = await runtime.skill.executeTool({
      id: 'edit',
      name: 'edit_presentation_page_text',
      input: {
        project_id: f.deck.id,
        page_id: f.deck.slides[1]!.id,
        shape_id: 'title',
        text: 'after',
      },
    })
    expect(edit.isError, edit.output).not.toBe(true)
    await runtime.proposals.confirm(runtime.proposals.pending()!.id)
    expect(editText.mock.calls[0]?.slice(0, 2)).toEqual(['host-2', 'title'])
    expect(pageText).toBe('after')
    expect(runtime.qa!.read()!.pages.map((p) => p.recheckRequired)).toEqual([undefined, true])
    const pendingEdit = await runtime.skill.executeTool({
      id: 'pending-edit',
      name: 'edit_presentation_page_text',
      input: {
        project_id: f.deck.id,
        page_id: f.deck.slides[1]!.id,
        shape_id: 'title',
        text: 'must-not-write',
      },
    })
    expect(pendingEdit.isError, pendingEdit.output).not.toBe(true)
    const pendingId = runtime.proposals.pending()!.id
    const switchSource = await runtime.skill.executeTool({
      id: 'switch',
      name: 'restore_presentation_project',
      input: { project_id: f.deck.id },
    })
    expect(switchSource.isError, switchSource.output).not.toBe(true)
    await expect(runtime.proposals.confirm(pendingId)).rejects.toThrow('proposal_stale')
    expect(pageText).toBe('after')
    expect(runtime.qa!.read()).toBeUndefined()
    const old = await runtime.skill.executeTool({
      id: 'old',
      name: 'import_generated_presentation',
      input: { project_id: f.deck.id },
    })
    expect(old.isError, old.output).not.toBe(true)
    expect(old.output).toContain('awaiting_confirmation')
    runtime.proposals.reject()
    expect(insert).toHaveBeenCalledTimes(8)
    for (const [index, base64] of bytes.entries()) {
      const saved = await f.call('production_page', {
        requestId: 'run',
        pageId: f.deck.slides[index]!.id,
      })
      expect(base64).toBe(saved.pptxBase64)
      const parsed = await openPptx(Buffer.from(base64, 'base64'))
      expect(parsed.deck.slides).toHaveLength(1)
    }
  } finally {
    runtime.dispose()
    inspect.mockRestore()
    readText.mockRestore()
    editText.mockRestore()
    vi.unstubAllGlobals()
  }
})

it('keeps the latest explicitly selected import view when an earlier preparation finishes late', async () => {
  const { createOfficeHostRuntime } = await import('../../office-addin/src/agent/host-runtime')
  const f = await setup()
  await f.call('production_begin', { requestId: 'run', planRevision: 1, deck: f.deck })
  await f.call('production_run', { requestId: 'run' })
  await f.call('compile', { requestId: 'whole', planRevision: 1, deck: f.deck })
  let release!: () => void, entered!: () => void
  const waiting = new Promise<void>((resolve) => {
    release = resolve
  })
  const started = new Promise<void>((resolve) => {
    entered = resolve
  })
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      documentId: async () => 'doc',
      lastProject: () => f.deck.id,
      rememberProject: async () => {},
      readReceipt: () => undefined,
      writeReceipt: async () => {},
      request: async (body, signal) => {
        if ((body as { operation: string }).operation === 'production_status') {
          entered()
          await waiting
        }
        return new Response(
          Buffer.from(await f.service(body, signal ?? new AbortController().signal)),
        )
      },
    },
  })
  try {
    const earlier = runtime.skill.executeTool({
      id: 'p',
      name: 'prepare_presentation_production_import',
      input: { project_id: f.deck.id, request_id: 'run' },
    })
    await started
    const latest = await runtime.skill.executeTool({
      id: 'r',
      name: 'restore_presentation_project',
      input: { project_id: f.deck.id },
    })
    expect(latest.isError, latest.output).not.toBe(true)
    release()
    const prepared = await earlier
    expect(prepared.isError, prepared.output).not.toBe(true)
    expect(runtime.importProgress!.read()).toMatchObject({ total: 8, completed: 0 })
    expect(runtime.importProgress!.read()?.source).toBeUndefined()
  } finally {
    release()
    runtime.dispose()
  }
})

it('derives a frozen single-page revision, reuses seven exact artifacts, and retries only its failed target', async () => {
  const { createOfficeHostRuntime } = await import('../../office-addin/src/agent/host-runtime')
  let failTarget = true
  const compile = vi.fn(async (input: unknown) => {
    const deck = parsePresentationDeck(input)
    if (deck.slides[0]?.notes === 'Single page revision' && failTarget) {
      failTarget = false
      throw new Error('temporary')
    }
    return compilePresentationDeck(deck)
  })
  const f = await setup(compile)
  await f.call('production_begin', { requestId: 'parent', planRevision: 1, deck: f.deck })
  expect(
    await f.call('production_rebuild_page', {
      parentRequestId: 'parent',
      requestId: 'child',
      pageId: f.deck.slides[1]!.id,
      slide: f.deck.slides[1],
    }),
  ).toEqual({ error: 'page_not_ready' })
  await f.call('production_run', { requestId: 'parent' })
  const originals = await Promise.all(
    f.deck.slides.map((slide) =>
      f.call('production_page', { requestId: 'parent', pageId: slide.id }),
    ),
  )
  const slide = structuredClone(f.deck.slides[1]!)
  slide.notes = 'Single page revision'
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => true,
      assetsAvailable: () => true,
      documentId: async () => 'doc',
      lastProject: () => f.deck.id,
      rememberProject: async () => {},
      request: async (body, signal) =>
        new Response(Buffer.from(await f.service(body, signal ?? new AbortController().signal))),
    },
  })
  const input = {
    project_id: f.deck.id,
    parent_request_id: 'parent',
    request_id: 'child',
    page_id: slide.id,
    slide,
  }
  try {
    const created = await runtime.skill.executeTool({
      id: 'rebuild',
      name: 'rebuild_presentation_page',
      input,
    })
    expect(created.isError, created.output).not.toBe(true)
    expect(JSON.parse(created.output)).toMatchObject({
      requestId: 'child',
      compiledCount: 7,
      revision: { parentRequestId: 'parent', pageId: slide.id },
    })
    expect(runtime.presentation!.snapshot().project?.production).toMatchObject({
      requestId: 'child',
      compiledCount: 7,
    })
    expect(compile).toHaveBeenCalledTimes(8)
    expect(await f.call('production_run', { requestId: 'child' })).toMatchObject({
      status: 'partial',
      compiledCount: 7,
    })
    const restarted = createPresentationService({ userDataPath: f.userDataPath, compile })
    const result = decode(
      await restarted(
        {
          operation: 'production_run',
          documentId: 'doc',
          projectId: f.deck.id,
          requestId: 'child',
        },
        new AbortController().signal,
      ),
    )
    expect(result).toMatchObject({ status: 'compiled', compiledCount: 8 })
    expect(compile).toHaveBeenCalledTimes(10)
    for (const [index, page] of f.deck.slides.entries()) {
      expect(await f.call('production_page', { requestId: 'parent', pageId: page.id })).toEqual(
        originals[index],
      )
      const child = await f.call('production_page', { requestId: 'child', pageId: page.id })
      if (index !== 1) expect(child.pptxBase64).toBe(originals[index].pptxBase64)
      else expect(child.pptxBase64).not.toBe(originals[index].pptxBase64)
      expect((await openPptx(Buffer.from(child.pptxBase64, 'base64'))).deck.slides).toHaveLength(1)
    }
    const plan = structuredClone(f.plan)
    plan.brief.objective = 'Updated current plan'
    await f.call('save_plan', { expectedRevision: 1, plan })
    expect(
      (await runtime.skill.executeTool({ id: 'repeat', name: 'rebuild_presentation_page', input }))
        .isError,
    ).not.toBe(true)
    expect(compile).toHaveBeenCalledTimes(10)
    const blocked = await runtime.skill.executeTool({
      id: 'prepare',
      name: 'prepare_presentation_production_import',
      input: { project_id: f.deck.id, request_id: 'child' },
    })
    expect(blocked).toMatchObject({
      isError: true,
      output: 'presentation_page_replacement_required',
    })
    expect(
      await f.call('production_rebuild_page', {
        documentId: 'foreign',
        parentRequestId: 'parent',
        requestId: 'other',
        pageId: slide.id,
        slide,
      }),
    ).toEqual({ error: 'document_mismatch' })
    expect(
      await f.call('production_rebuild_page', {
        parentRequestId: 'parent',
        requestId: 'bad',
        pageId: slide.id,
        slide: { ...slide, title: 'Changed plan title' },
      }),
    ).toEqual({ error: 'plan_mismatch' })
  } finally {
    runtime.dispose()
  }
})
