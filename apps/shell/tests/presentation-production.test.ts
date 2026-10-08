import { confirmReviewed } from '../../office-addin/tests/presentation-lock-review-fixture.js'
import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import { openPptx } from '@wiswork/pptx-engine'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { PresentationStore } from '@wiswork/project-store'
import { parsePresentationDeck } from '@wiswork/pptx-engine/presentation'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
import { assertCitedPresentationSourcesReady } from '../src/main/presentation-production'
import {
  convertSinglePagePackageToPng,
  libreOfficeCommands,
} from '../src/main/presentation-page-render'
const sofficeAvailable = libreOfficeCommands().some(
  (command) => spawnSync(command, ['--version'], { timeout: 5_000 }).status === 0,
)
if (process.env.WISWORK_REQUIRE_LIBREOFFICE === '1' && !sofficeAvailable)
  throw new Error('LibreOffice is required for presentation visual comparison')
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const decode = (bytes: Uint8Array) => JSON.parse(Buffer.from(bytes).toString('utf8'))
it('blocks production when a fetched snapshot URL differs from the cited plan URL', async () => {
  const plan = benchmarkPlan()
  const deck = benchmarkPlannedDeck()
  const attachmentId = 'a'.repeat(64)
  const original = 'https://example.com/report?edition=1'
  plan.sources[0]!.uri = 'https://example.com/report?edition=2'
  plan.sources[0]!.snapshotAttachmentId = attachmentId
  const sourceUrlHash = createHash('sha256').update(original).digest('hex')
  const attachments = vi.fn(async (request: Record<string, unknown>) =>
    request.operation === 'attachment_metadata'
      ? { attachmentId, sourceUrlHash }
      : { attachmentId, status: 'found', offset: 0 },
  )
  await expect(
    assertCitedPresentationSourcesReady(
      plan,
      deck.slides[0]!,
      'doc',
      attachments,
      new AbortController().signal,
      new Map(),
    ),
  ).rejects.toThrow('source_unavailable')
  expect(attachments).toHaveBeenCalledTimes(1)
  plan.sources[0]!.uri = original
  await expect(
    assertCitedPresentationSourcesReady(
      plan,
      deck.slides[0]!,
      'doc',
      attachments,
      new AbortController().signal,
      new Map(),
    ),
  ).resolves.toBeUndefined()
})
it('blocks a cited PDF excerpt whose verified page differs from the plan locator', async () => {
  const plan = benchmarkPlan()
  const deck = benchmarkPlannedDeck()
  const attachmentId = 'a'.repeat(64)
  plan.sources[0]!.uri = `attachment:${attachmentId}`
  plan.sources[0]!.locator = '第 1 页'
  const attachments = vi.fn(async (request: Record<string, unknown>) =>
    request.operation === 'attachment_metadata'
      ? { attachmentId }
      : { attachmentId, status: 'found', offset: 10, locator: '第 2 页' },
  )
  await expect(
    assertCitedPresentationSourcesReady(
      plan,
      deck.slides[0]!,
      'doc',
      attachments,
      new AbortController().signal,
      new Map(),
    ),
  ).rejects.toThrow('source_unavailable')
  plan.sources[0]!.locator = '第 2 页'
  await expect(
    assertCitedPresentationSourcesReady(
      plan,
      deck.slides[0]!,
      'doc',
      attachments,
      new AbortController().signal,
      new Map(),
    ),
  ).resolves.toBeUndefined()
})
async function setup(compile = vi.fn(compilePresentationDeck), documentId = 'doc') {
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-page-production-'))
  roots.push(userDataPath)
  const service = createPresentationService({ userDataPath, compile }),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const call = async (operation: string, extra = {}, signal = new AbortController().signal) =>
    decode(await service({ operation, documentId, projectId: deck.id, ...extra }, signal))
  expect((await call('save_plan', { expectedRevision: 0, plan })).revision).toBe(1)
  return { call, service, userDataPath, compile, deck, plan }
}
it('reuses unchanged compiled pages across plan revisions and compiles only the affected page', async () => {
  const f = await setup()
  await f.call('production_begin', { requestId: 'original', planRevision: 1, deck: f.deck })
  await f.call('production_run', { requestId: 'original' })
  expect(f.compile).toHaveBeenCalledTimes(f.deck.slides.length)
  f.compile.mockClear()
  const plan = structuredClone(f.plan)
  plan.slides[2]!.purpose = '更新本页讲述重点'
  expect(await f.call('save_plan', { expectedRevision: 1, plan })).toMatchObject({ revision: 2 })
  const started = await f.call('production_begin', {
    requestId: 'revised',
    planRevision: 2,
    deck: f.deck,
  })
  expect(started.compiledCount).toBe(f.deck.slides.length - 1)
  expect(started.pages[2]).toMatchObject({ state: 'pending', attempt: 0 })
  expect(started.pages[0]).toMatchObject({ state: 'compiled', reusedFromRequestId: 'original' })
  const finished = await f.call('production_run', { requestId: 'revised' })
  expect(finished.status).toBe('compiled')
  expect(f.compile).toHaveBeenCalledTimes(1)
  expect(parsePresentationDeck(f.compile.mock.calls[0]![0]).slides[0]!.id).toBe(
    f.deck.slides[2]!.id,
  )
  const page = await f.call('production_page', {
    requestId: 'revised',
    pageId: f.deck.slides[0]!.id,
  })
  expect(page.planRevision).toBe(2)
  expect((await openPptx(Buffer.from(page.pptxBase64, 'base64'))).deck.slides).toHaveLength(1)
  const reopened = createPresentationService({ userDataPath: f.userDataPath, compile: f.compile })
  const response = decode(
    await reopened(
      {
        operation: 'production_status',
        documentId: 'doc',
        projectId: f.deck.id,
        requestId: 'revised',
      },
      new AbortController().signal,
    ),
  )
  expect(response).toEqual(finished)
  expect((await f.call('production_status', { requestId: 'original' })).planRevision).toBe(1)
})

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
it('compiles opted-in independent pages concurrently and waits for declared dependencies', async () => {
  const entered = [0, 0]
  const gates = [
    (() => {
      let release!: () => void
      const promise = new Promise<void>((resolve) => {
        release = resolve
      })
      return { promise, release }
    })(),
    (() => {
      let release!: () => void
      const promise = new Promise<void>((resolve) => {
        release = resolve
      })
      return { promise, release }
    })(),
  ]
  const compile = vi.fn(async (input: unknown) => {
    const slideId = parsePresentationDeck(input).slides[0]!.id
    const index = slideId === 'slide-1' ? 0 : 1
    entered[index]!++
    await gates[index]!.promise
    return compilePresentationDeck(input)
  })
  const f = await setup(compile)
  f.plan.slides = f.plan.slides.slice(0, 2)
  f.deck.slides = f.deck.slides.slice(0, 2)
  f.plan.parallelism = 2
  f.plan.slides.forEach((slide) => {
    slide.dependsOn = []
  })
  expect((await f.call('save_plan', { expectedRevision: 1, plan: f.plan })).revision).toBe(2)
  await f.call('production_begin', { requestId: 'parallel', planRevision: 2, deck: f.deck })
  const parallel = f.call('production_run', { requestId: 'parallel' })
  await vi.waitFor(() => expect(entered).toEqual([1, 1]))
  gates.forEach((gate) => gate.release())
  expect(await parallel).toMatchObject({ status: 'compiled', compiledCount: 2 })

  const dependency = structuredClone(f.plan)
  dependency.slides[1]!.dependsOn = ['slide-1']
  expect((await f.call('save_plan', { expectedRevision: 2, plan: dependency })).revision).toBe(3)
  const waiting = [
    (() => {
      let release!: () => void
      const promise = new Promise<void>((resolve) => {
        release = resolve
      })
      return { promise, release }
    })(),
    (() => {
      let release!: () => void
      const promise = new Promise<void>((resolve) => {
        release = resolve
      })
      return { promise, release }
    })(),
  ]
  entered.fill(0)
  compile.mockImplementation(async (input: unknown) => {
    const slideId = parsePresentationDeck(input).slides[0]!.id
    const index = slideId === 'slide-1' ? 0 : 1
    entered[index]!++
    await waiting[index]!.promise
    return compilePresentationDeck(input)
  })
  // Both pages have new content so this tests scheduling pending dependencies,
  // rather than the now-supported reuse of an already completed predecessor.
  const dependencyDeck = {
    ...f.deck,
    slides: f.deck.slides.map((slide) => ({ ...slide, notes: '更新后的讲述内容' })),
  }
  await f.call('production_begin', {
    requestId: 'dependent',
    planRevision: 3,
    deck: dependencyDeck,
  })
  const dependent = f.call('production_run', { requestId: 'dependent' })
  await vi.waitFor(() => expect(entered[0]).toBe(1))
  expect(entered[1]).toBe(0)
  waiting[0]!.release()
  await vi.waitFor(() => expect(entered[1]).toBe(1))
  waiting[1]!.release()
  expect(await dependent).toMatchObject({ status: 'compiled', compiledCount: 2 })
})
it('starts a ready dependent page while an unrelated earlier page is still compiling', async () => {
  let releaseSlow!: () => void
  const slow = new Promise<void>((resolve) => {
    releaseSlow = resolve
  })
  const entered: string[] = []
  const compile = vi.fn(async (input: unknown) => {
    const pageId = parsePresentationDeck(input).slides[0]!.id
    entered.push(pageId)
    if (pageId === 'slide-2') await slow
    return compilePresentationDeck(input)
  })
  const f = await setup(compile)
  f.plan.slides = f.plan.slides.slice(0, 3)
  f.deck.slides = f.deck.slides.slice(0, 3)
  f.plan.parallelism = 2
  f.plan.slides[0]!.dependsOn = []
  f.plan.slides[1]!.dependsOn = []
  f.plan.slides[2]!.dependsOn = ['slide-1']
  await f.call('save_plan', { expectedRevision: 1, plan: f.plan })
  await f.call('production_begin', { requestId: 'early-dependent', planRevision: 2, deck: f.deck })
  const running = f.call('production_run', { requestId: 'early-dependent' })
  try {
    await vi.waitFor(() => expect(entered).toContain('slide-2'))
    await vi.waitFor(() => expect(entered).toContain('slide-3'))
    const store = new PresentationStore(f.userDataPath)
    await vi.waitFor(() => {
      const mid = store.production(f.deck.id, 'doc', 'early-dependent')!
      expect(mid.pages[1]).toMatchObject({ state: 'building', attempt: 1 })
      expect(mid.pages[2]).toMatchObject({ state: 'compiled', attempt: 1 })
    })
  } finally {
    releaseSlow()
  }
  expect(await running).toMatchObject({ status: 'compiled', compiledCount: 3 })
  expect(compile).toHaveBeenCalledTimes(3)
})
it('keeps native page structure and explicit style identical between serial and parallel production', async () => {
  const f = await setup()
  await f.call('production_begin', { requestId: 'serial-style', planRevision: 1, deck: f.deck })
  expect((await f.call('production_run', { requestId: 'serial-style' })).status).toBe('compiled')
  f.plan.parallelism = 2
  f.plan.slides.forEach((slide) => {
    slide.dependsOn = []
  })
  expect((await f.call('save_plan', { expectedRevision: 1, plan: f.plan })).revision).toBe(2)
  await f.call('production_begin', { requestId: 'parallel-style', planRevision: 2, deck: f.deck })
  expect((await f.call('production_run', { requestId: 'parallel-style' })).status).toBe('compiled')
  const structure = async (requestId: string, pageId: string) => {
    const page = await f.call('production_page', { requestId, pageId })
    const opened = await openPptx(Buffer.from(page.pptxBase64, 'base64'))
    return opened.deck.slides[0]!.elements.map((element) => ({
      type: element.type,
      name: element.name,
      transform: element.transform,
      ...(element.type === 'shape'
        ? {
            text: element.text,
            fill: element.fill,
            line: element.line,
            presetGeometry: element.presetGeometry,
          }
        : {}),
      ...(element.type === 'table' ? { rows: element.rows } : {}),
      ...(element.type === 'chart' ? { chart: element.chart } : {}),
    }))
  }
  for (const slide of f.deck.slides)
    expect(await structure('parallel-style', slide.id)).toEqual(
      await structure('serial-style', slide.id),
    )
})
it.skipIf(!sofficeAvailable)(
  'renders all eight serial and parallel pages identically in LibreOffice',
  async () => {
    const f = await setup()
    await f.call('production_begin', { requestId: 'serial-visual', planRevision: 1, deck: f.deck })
    expect((await f.call('production_run', { requestId: 'serial-visual' })).status).toBe('compiled')
    f.plan.parallelism = 2
    f.plan.slides.forEach((slide) => {
      slide.dependsOn = []
    })
    expect((await f.call('save_plan', { expectedRevision: 1, plan: f.plan })).revision).toBe(2)
    await f.call('production_begin', {
      requestId: 'parallel-visual',
      planRevision: 2,
      deck: f.deck,
    })
    expect((await f.call('production_run', { requestId: 'parallel-visual' })).status).toBe(
      'compiled',
    )
    for (const slide of f.deck.slides) {
      const [serial, parallel] = await Promise.all(
        ['serial-visual', 'parallel-visual'].map(async (requestId) => {
          const page = await f.call('production_page', { requestId, pageId: slide.id })
          return Buffer.from(
            await convertSinglePagePackageToPng(
              Buffer.from(page.pptxBase64, 'base64'),
              new AbortController().signal,
            ),
          )
        }),
      )
      expect(parallel.equals(serial), `Rendered page ${slide.id} drifted`).toBe(true)
    }
  },
  180_000,
)
it('holds a dependent page pending when its predecessor fails, then resumes both safely', async () => {
  let fail = true
  const compile = vi.fn(async (input: unknown) => {
    const slideId = parsePresentationDeck(input).slides[0]!.id
    if (slideId === 'slide-1' && fail) {
      fail = false
      throw new Error('transient')
    }
    return compilePresentationDeck(input)
  })
  const f = await setup(compile)
  f.plan.slides = f.plan.slides.slice(0, 2)
  f.deck.slides = f.deck.slides.slice(0, 2)
  f.plan.parallelism = 2
  f.plan.slides.forEach((slide) => {
    slide.dependsOn = []
  })
  f.plan.slides[1]!.dependsOn = ['slide-1']
  await f.call('save_plan', { expectedRevision: 1, plan: f.plan })
  await f.call('production_begin', { requestId: 'dependent', planRevision: 2, deck: f.deck })
  const first = await f.call('production_run', { requestId: 'dependent' })
  expect(first).toMatchObject({ status: 'partial', compiledCount: 0 })
  expect(first.pages[0]).toMatchObject({ state: 'failed', attempt: 1 })
  expect(first.pages[1]).toMatchObject({ state: 'pending', attempt: 0 })
  expect(compile).toHaveBeenCalledTimes(1)
  const resumed = await f.call('production_run', { requestId: 'dependent' })
  expect(resumed).toMatchObject({ status: 'compiled', compiledCount: 2 })
  expect(compile).toHaveBeenCalledTimes(3)
})
it('refuses to compile a page whose planned brand logo bytes do not match', async () => {
  const f = await setup()
  f.plan.slides[2]!.layout = 'cover'
  f.plan.brandKit = {
    id: 'research',
    revision: 1,
    name: '研究品牌',
    allowedColors: [f.plan.style.background, f.plan.style.textColor, f.plan.style.accentColor],
    logo: { assetId: 'pixel', assetDigest: '0'.repeat(64), placement: 'cover' },
  }
  expect((await f.call('save_plan', { expectedRevision: 1, plan: f.plan })).revision).toBe(2)
  expect(
    await f.call('production_begin', { requestId: 'branded', planRevision: 2, deck: f.deck }),
  ).toMatchObject({ status: 'pending' })
  const run = await f.call('production_run', { requestId: 'branded' })
  expect(run).toMatchObject({ status: 'partial', compiledCount: 7 })
  expect(run.pages[2]).toMatchObject({ state: 'failed', error: 'invalid_deck' })
  expect(f.compile).toHaveBeenCalledTimes(7)
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
  let beforeCancellation: unknown
  const compile = vi.fn(async (input: unknown) => {
    const result = await compilePresentationDeck(input)
    if (++calls === 2) {
      beforeCancellation = new PresentationStore(f.userDataPath).production(f.deck.id, 'doc', 'run')
      controller.abort()
    }
    return result
  })
  const f = await setup(compile)
  await f.call('production_begin', { requestId: 'run', planRevision: 1, deck: f.deck })
  expect(await f.call('production_run', { requestId: 'run' }, controller.signal)).toEqual({
    error: 'aborted',
  })
  const status = await f.call('production_status', { requestId: 'run' })
  expect(status).toMatchObject({ compiledCount: 1, status: 'building' })
  expect(status.pages[1]).toMatchObject({ state: 'building', attempt: 1 })
  expect(status.pages[1].error).toBeUndefined()
  expect(new PresentationStore(f.userDataPath).production(f.deck.id, 'doc', 'run')).toEqual(
    beforeCancellation,
  )
  expect(status.pages[2]).toMatchObject({ state: 'pending', attempt: 0 })
  expect(compile).toHaveBeenCalledTimes(2)
  expect(await f.call('production_run', { requestId: 'run' })).toMatchObject({ status: 'compiled' })
  expect(compile).toHaveBeenCalledTimes(9)
})
it('resumes P0-20 source-backed production after four persisted pages without recompiling them', async () => {
  const material = new URL(
    '../../../docs/product/ppt-benchmark-materials/PPT-P0-20/',
    import.meta.url,
  )
  const plan = JSON.parse(readFileSync(new URL('reference-plan.json', material), 'utf8'))
  const deck = parsePresentationDeck(
    JSON.parse(readFileSync(new URL('reference-deck.json', material), 'utf8')),
  )
  const controller = new AbortController()
  const compiledPageIds: string[] = []
  const compile = vi.fn(async (input: unknown) => {
    const page = parsePresentationDeck(input)
    compiledPageIds.push(page.slides[0]!.id)
    const result = await compilePresentationDeck(page)
    if (compiledPageIds.length === 5) controller.abort()
    return result
  })
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-p0-20-production-'))
  roots.push(userDataPath)
  const service = createPresentationService({ userDataPath, compile })
  const call = async (operation: string, extra = {}, signal = new AbortController().signal) =>
    decode(
      await service(
        {
          operation,
          documentId: 'doc',
          ...(!operation.startsWith('attachment_') ? { projectId: deck.id } : {}),
          ...extra,
        },
        signal,
      ),
    )
  for (const name of ['deardorff-2020-article.pdf', 'deardorff-2020-checklist.pdf']) {
    const bytes = readFileSync(new URL(name, material))
    const attachmentId = createHash('sha256').update(bytes).digest('hex')
    expect(
      await call('attachment_begin', {
        attachmentId,
        sha256: attachmentId,
        name,
        sizeBytes: bytes.length,
      }),
    ).toMatchObject({ status: 'uploading' })
    for (let offset = 0; offset < bytes.length; offset += 64 * 1024)
      await call('attachment_chunk', {
        attachmentId,
        offset,
        base64: bytes.subarray(offset, offset + 64 * 1024).toString('base64'),
      })
    expect(await call('attachment_finish', { attachmentId })).toMatchObject({ status: 'ready' })
  }
  expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
  await call('production_begin', { requestId: 'p0-20-recovery', planRevision: 1, deck })
  expect(await call('production_run', { requestId: 'p0-20-recovery' }, controller.signal)).toEqual({
    error: 'aborted',
  })
  const interrupted = await call('production_status', { requestId: 'p0-20-recovery' })
  expect(interrupted).toMatchObject({ status: 'building', compiledCount: 4, total: 8 })
  expect(interrupted.pages[4]).toMatchObject({ state: 'building', attempt: 1 })
  expect(
    interrupted.pages.slice(5).every((page: { state: string }) => page.state === 'pending'),
  ).toBe(true)
  const firstFour = await Promise.all(
    deck.slides
      .slice(0, 4)
      .map((slide) => call('production_page', { requestId: 'p0-20-recovery', pageId: slide.id })),
  )
  const reopened = createPresentationService({ userDataPath, compile })
  const resumed = decode(
    await reopened(
      {
        operation: 'production_run',
        documentId: 'doc',
        projectId: deck.id,
        requestId: 'p0-20-recovery',
      },
      new AbortController().signal,
    ),
  )
  expect(resumed).toMatchObject({ status: 'compiled', compiledCount: 8, total: 8 })
  expect(compiledPageIds).toEqual([
    ...deck.slides.slice(0, 5).map((slide) => slide.id),
    ...deck.slides.slice(4).map((slide) => slide.id),
  ])
  for (const [index, slide] of deck.slides.slice(0, 4).entries())
    expect(
      await call('production_page', { requestId: 'p0-20-recovery', pageId: slide.id }),
    ).toEqual(firstFour[index])
  const chartPage = await call('production_page', {
    requestId: 'p0-20-recovery',
    pageId: deck.slides[5]!.id,
  })
  const native = (await openPptx(Buffer.from(chartPage.pptxBase64, 'base64'))).deck.slides[0]!
  expect(native.elements.some((element) => element.type === 'chart')).toBe(true)
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
  const history = (await f.call('status')).assetHistory
  expect(history.scope).toBe('production_asset_resolution')
  expect(
    history.events.filter((event: { type: string }) => event.type === 'asset.rejected'),
  ).toHaveLength(3)
  expect(history.events.at(-1)).toMatchObject({
    type: 'asset.rejected',
    pageId: affected[0]!.id,
    error: 'asset_unavailable',
    attempt: 3,
  })
  expect(new PresentationStore(f.userDataPath).productionAssets(deck.id, 'doc', 'run')).toEqual(
    history,
  )
})
it.each(['legacy', 'url_snapshot'] as const)(
  'isolates a missing %s source to its citing page and retries after upload',
  async (kind) => {
    const f = await setup()
    const plan = structuredClone(f.plan)
    const deck = structuredClone(f.deck)
    const raw = Buffer.from('示例数据仅用于测试')
    const attachmentId = createHash('sha256').update(raw).digest('hex')
    plan.sources[0]!.uri =
      kind === 'legacy' ? `attachment:${attachmentId}` : 'https://example.com/research'
    if (kind === 'url_snapshot') plan.sources[0]!.snapshotAttachmentId = attachmentId
    for (let index = 1; index < plan.slides.length; index++) {
      plan.slides[index]!.claimIds = []
      deck.slides[index]!.claimIds = []
    }
    const { presentationPlanClaims } = await import('@wiswork/pptx-engine/presentation-plan')
    deck.claims = presentationPlanClaims(plan)
    expect(await f.call('save_plan', { expectedRevision: 1, plan })).toMatchObject({ revision: 2 })
    await f.call('production_begin', { requestId: 'source-run', planRevision: 2, deck })
    const first = await f.call('production_run', { requestId: 'source-run' })
    expect(first).toMatchObject({ status: 'partial', compiledCount: 7 })
    expect(first.pages[0]).toMatchObject({ state: 'failed', error: 'source_unavailable' })
    expect(first.pages.slice(1).every((page: { state: string }) => page.state === 'compiled')).toBe(
      true,
    )
    expect(f.compile).toHaveBeenCalledTimes(7)
    const attachment = async (operation: string, extra: Record<string, unknown>) =>
      decode(
        await f.service({ operation, documentId: 'doc', ...extra }, new AbortController().signal),
      )
    await attachment('attachment_begin', {
      attachmentId,
      sha256: attachmentId,
      name: 'source.txt',
      sizeBytes: raw.length,
    })
    await attachment('attachment_chunk', {
      attachmentId,
      offset: 0,
      base64: raw.toString('base64'),
    })
    expect(await attachment('attachment_finish', { attachmentId })).toMatchObject({
      status: 'ready',
    })
    const textPath = join(
      f.userDataPath,
      'presentation-attachments',
      createHash('sha256').update('doc').digest('hex'),
      attachmentId,
      'text.txt',
    )
    writeFileSync(textPath, 'corrupted')
    const corrupted = await f.call('production_run', { requestId: 'source-run' })
    expect(corrupted.pages[0]).toMatchObject({
      state: 'failed',
      attempt: 2,
      error: 'source_unavailable',
    })
    expect(f.compile).toHaveBeenCalledTimes(7)
    writeFileSync(textPath, raw)
    const resumed = await f.call('production_run', { requestId: 'source-run' })
    expect(resumed).toMatchObject({ status: 'compiled', compiledCount: 8 })
    expect(resumed.pages[0]).toMatchObject({ state: 'compiled', attempt: 3 })
    expect(f.compile).toHaveBeenCalledTimes(8)
  },
)
it('isolates a cited excerpt absent from a readable attachment to its page', async () => {
  const f = await setup()
  const plan = structuredClone(f.plan)
  const deck = structuredClone(f.deck)
  const raw = Buffer.from('Actual source text')
  const attachmentId = createHash('sha256').update(raw).digest('hex')
  plan.sources[0]!.uri = `attachment:${attachmentId}`
  plan.sources[0]!.excerpt = 'Invented source text'
  for (let index = 1; index < plan.slides.length; index++) {
    plan.slides[index]!.claimIds = []
    deck.slides[index]!.claimIds = []
  }
  const { presentationPlanClaims } = await import('@wiswork/pptx-engine/presentation-plan')
  deck.claims = presentationPlanClaims(plan)
  expect(await f.call('save_plan', { expectedRevision: 1, plan })).toMatchObject({ revision: 2 })
  const attachment = async (operation: string, extra: Record<string, unknown>) =>
    decode(
      await f.service({ operation, documentId: 'doc', ...extra }, new AbortController().signal),
    )
  await attachment('attachment_begin', {
    attachmentId,
    sha256: attachmentId,
    name: 'source.txt',
    sizeBytes: raw.length,
  })
  await attachment('attachment_chunk', { attachmentId, offset: 0, base64: raw.toString('base64') })
  expect(await attachment('attachment_finish', { attachmentId })).toMatchObject({ status: 'ready' })
  await f.call('production_begin', { requestId: 'excerpt-run', planRevision: 2, deck })
  const result = await f.call('production_run', { requestId: 'excerpt-run' })
  expect(result).toMatchObject({ status: 'partial', compiledCount: 7 })
  expect(result.pages[0]).toMatchObject({ state: 'failed', error: 'source_unavailable' })
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
  const f = await setup(undefined, await binding.documentId())
  await f.call('production_begin', { requestId: 'run', planRevision: 1, deck: f.deck })
  await f.call('production_run', { requestId: 'run' })
  await f.call('compile', { requestId: 'whole', planRevision: 1, deck: f.deck })
  const hostIds = ['original']
  const slides = {
    items: [] as { id: string }[],
    load: () => {},
    getItem: (id: string) => ({
      exportAsBase64: () => ({ value: bytes[hostIds.indexOf(id) - 1] }),
    }),
  }
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
        isSetSupported: (_name: string, version: string) =>
          ['1.2', '1.8', '1.10'].includes(version),
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
    const importProposal = await runtime.skill.executeTool(importCall)
    expect(importProposal.isError, importProposal.output).not.toBe(true)
    await expect(
      confirmReviewed(runtime.proposals, runtime.proposals.pending()!.id),
    ).rejects.toThrow('cancelled')
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
    await confirmReviewed(runtime.proposals, runtime.proposals.pending()!.id)
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
    await confirmReviewed(runtime.proposals, runtime.proposals.pending()!.id)
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
    await expect(confirmReviewed(runtime.proposals, pendingId)).rejects.toThrow('proposal_stale')
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

it.each([false, true])(
  'fences asset success and failure receipts after lifecycle deletion (%s)',
  async (reject) => {
    const { PresentationLifecycleStore } = await import('@wiswork/project-store')
    const { handlePresentationProduction } = await import('../src/main/presentation-production')
    const f = await setup(),
      store = new PresentationStore(f.userDataPath)
    const deck = structuredClone(f.deck)
    const originalAsset = deck.assets[0]!
    deck.assets = deck.assets.map((asset) => ({ id: asset.id, attachmentId: 'f'.repeat(64) }))
    await f.call('production_begin', { requestId: 'asset-fence', planRevision: 1, deck })
    const lifecycle = new PresentationLifecycleStore(f.userDataPath),
      scope = { projectId: deck.id, documentId: 'doc' }
    lifecycle.initialize(scope)
    let release!: () => void, enter!: () => void
    const gate = new Promise<void>((r) => {
        release = r
      }),
      entered = new Promise<void>((r) => {
        enter = r
      })
    const run = handlePresentationProduction(
      { ...scope, operation: 'production_run', requestId: 'asset-fence' },
      {
        store,
        compile: f.compile,
        assertWritable: () => lifecycle.assertActive(scope, 0),
        attachments: async () => {
          enter()
          await gate
          if (reject) throw Error('fetch failed')
          return originalAsset
        },
      },
      new AbortController().signal,
    )
    const settled = run.then(
      () => null,
      (error) => error,
    )
    await entered
    const before = store.production(deck.id, 'doc', 'asset-fence'),
      assets = store.productionAssets(deck.id, 'doc', 'asset-fence')
    lifecycle.beginDeletion(scope, 0, {
      deletionId: 'delete',
      reason: 'user',
      resources: [{ resourceId: 'project', kind: 'project', ownership: 'project_exclusive' }],
    })
    release()
    expect(await settled).toBeInstanceOf(Error)
    expect(store.production(deck.id, 'doc', 'asset-fence')).toEqual(before)
    expect(store.productionAssets(deck.id, 'doc', 'asset-fence')).toEqual(assets)
  },
)
