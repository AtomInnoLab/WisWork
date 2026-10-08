import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PDFDocument } from 'pdf-lib'
import { PresentationLifecycleStore } from '@wiswork/project-store'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { createPresentationService } from '../src/main/presentation-service'
const roots: string[] = []
afterEach(() => {
  vi.restoreAllMocks()
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const decode = (bytes: Uint8Array) => JSON.parse(Buffer.from(bytes).toString('utf8'))
function gate() {
  let resolve!: () => void
  const promise = new Promise<void>((r) => {
    resolve = r
  })
  return { resolve, promise }
}
it.each([false, true])(
  'PDF export cannot return private bytes after lifecycle freeze (renderer fails=%s)',
  async (fails) => {
    const root = mkdtempSync(join(tmpdir(), 'pdf-write-fences-'))
    roots.push(root)
    const entered = gate(),
      release = gate(),
      deck = benchmarkDeck()
    const pdf = await PDFDocument.create()
    for (const _ of deck.slides) pdf.addPage([960, 540])
    const bytes = await pdf.save()
    const service = createPresentationService({
      userDataPath: root,
      renderPdf: async () => {
        entered.resolve()
        await release.promise
        if (fails) throw Error('renderer_unavailable')
        return bytes
      },
    })
    const signal = new AbortController().signal
    const scope = { projectId: deck.id, documentId: 'doc' }
    expect(
      decode(await service({ ...scope, operation: 'compile', requestId: 'r', deck }, signal))
        .status,
    ).toBe('compiled')
    const control = new PresentationLifecycleStore(root)
    const state = control.read(scope) ?? control.initialize(scope)
    const pending = service({ ...scope, operation: 'export_pdf', requestId: 'r' }, signal)
    await entered.promise
    control.beginDeletion(scope, state.revision, {
      deletionId: 'delete',
      reason: 'user',
      resources: [{ resourceId: 'own', kind: 'project', ownership: 'project_exclusive' }],
    })
    release.resolve()
    expect(decode(await pending)).toEqual({ error: 'revision_conflict' })
  },
)

it('owns export identity aliases while rendering without changing the frozen source', async () => {
  const root = mkdtempSync(join(tmpdir(), 'pdf-alias-fence-'))
  roots.push(root)
  const entered = gate(),
    release = gate(),
    deck = benchmarkDeck(),
    pdf = await PDFDocument.create()
  for (const _ of deck.slides) pdf.addPage([960, 540])
  const bytes = await pdf.save(),
    service = createPresentationService({
      userDataPath: root,
      renderPdf: async () => {
        entered.resolve()
        await release.promise
        return bytes
      },
    }),
    signal = new AbortController().signal
  const scope = { projectId: deck.id, documentId: 'doc' }
  expect(
    decode(await service({ ...scope, operation: 'compile', requestId: 'r', deck }, signal)).status,
  ).toBe('compiled')
  const request = { ...scope, operation: 'export_pdf', requestId: 'r', source: 'compiled' }
  const pending = service(request, signal)
  await entered.promise
  request.projectId = 'foreign'
  request.documentId = 'foreign'
  request.requestId = 'foreign'
  request.source = 'production'
  release.resolve()
  expect(decode(await pending)).toMatchObject({
    status: 'exported',
    projectId: deck.id,
    requestId: 'r',
    source: 'compiled',
    slideCount: deck.slides.length,
  })
})
it('checks the original lease after actual production PDF copy before rendering the next page', async () => {
  const root = mkdtempSync(join(tmpdir(), 'production-pdf-copy-fence-'))
  roots.push(root)
  const entered = gate(),
    release = gate(),
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  plan.slides = plan.slides.slice(0, 2)
  deck.slides = deck.slides.slice(0, 2)
  const pdf = await PDFDocument.create()
  pdf.addPage([960, 540])
  const bytes = await pdf.save()
  const renderPdf = vi.fn(async () => bytes),
    service = createPresentationService({
      userDataPath: root,
      renderPdf,
      compile: (input, settings) =>
        compilePresentationDeck(input, { ...settings, fontAvailable: () => true }),
    }),
    signal = new AbortController().signal,
    scope = { projectId: deck.id, documentId: 'doc' }
  const call = async (operation: string, extra = {}) =>
    decode(await service({ ...scope, operation, ...extra }, signal))
  expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
  expect(await call('production_begin', { requestId: 'r', planRevision: 1, deck })).toMatchObject({
    total: 2,
  })
  expect(await call('production_run', { requestId: 'r' })).toMatchObject({ compiledCount: 2 })
  const original = PDFDocument.prototype.copyPages
  vi.spyOn(PDFDocument.prototype, 'copyPages').mockImplementation(
    async function (document, indices) {
      const result = await original.call(this, document, indices)
      entered.resolve()
      await release.promise
      return result
    },
  )
  const pending = call('export_pdf', { source: 'production', requestId: 'r' })
  await entered.promise
  const control = new PresentationLifecycleStore(root),
    state = control.read(scope)!
  control.beginDeletion(scope, state.revision, {
    deletionId: 'delete',
    reason: 'user',
    resources: [{ resourceId: 'own', kind: 'project', ownership: 'project_exclusive' }],
  })
  release.resolve()
  expect(await pending).toEqual({ error: 'revision_conflict' })
  expect(renderPdf).toHaveBeenCalledTimes(1)
})
