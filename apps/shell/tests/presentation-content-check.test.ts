import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, expect, it, vi } from 'vitest'
import { PresentationStore } from '@wiswork/project-store'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
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
function files(root: string): unknown {
  return readdirSync(root, { recursive: true, withFileTypes: true })
    .filter((entry) => entry.isFile())
    .map((entry) => [
      join(entry.parentPath, entry.name),
      readFileSync(join(entry.parentPath, entry.name)).toString('base64'),
    ])
}
async function setup() {
  const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-content-check-'))
  roots.push(userDataPath)
  const compile = vi.fn(compilePresentationDeck)
  const service = createPresentationService({ userDataPath, compile })
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  const call = async (operation: string, extra = {}, signal = new AbortController().signal) =>
    decode(await service({ operation, documentId: 'doc', projectId: deck.id, ...extra }, signal))
  await call('save_plan', { expectedRevision: 0, plan })
  await call('production_begin', { requestId: 'run', planRevision: 1, deck })
  const request = {
    operation: 'production_content_check',
    documentId: 'doc',
    projectId: deck.id,
    requestId: 'run',
    pageId: deck.slides[0]!.id,
  }
  return { call, service, userDataPath, compile, plan, deck, request }
}
it('checks the explicitly selected frozen production before compilation and across restart without writes', async () => {
  const f = await setup()
  const record = new PresentationStore(f.userDataPath).production(f.deck.id, 'doc', 'run')!
  const before = files(f.userDataPath)
  const result = decode(await f.service(f.request, new AbortController().signal))
  expect(result).toMatchObject({
    projectId: f.deck.id,
    requestId: 'run',
    planRevision: 1,
    inputDigest: record.inputDigest,
    planDigest: record.planDigest,
    report: { version: 1, pageId: f.request.pageId, claimIds: ['source-1'] },
  })
  expect(files(f.userDataPath)).toEqual(before)
  const plan = structuredClone(f.plan)
  plan.sources[0]!.excerpt = ''
  await f.call('save_plan', { expectedRevision: 1, plan })
  await f.call('production_begin', { requestId: 'new', planRevision: 2, deck: f.deck })
  const restarted = createPresentationService({ userDataPath: f.userDataPath, compile: f.compile })
  const updatedFiles = files(f.userDataPath)
  expect(decode(await restarted(f.request, new AbortController().signal))).toEqual(result)
  const latest = decode(
    await restarted({ ...f.request, requestId: 'new' }, new AbortController().signal),
  )
  expect(latest.report.findings).toContainEqual({
    code: 'source_excerpt_missing',
    claimId: 'source-1',
    sourceId: 'source',
  })
  expect(latest.planRevision).toBe(2)
  expect(files(f.userDataPath)).toEqual(updatedFiles)
  expect(f.compile).not.toHaveBeenCalled()
})
it('reports configured arithmetic consistently across frozen precheck and delivery report', async () => {
  const f = await setup()
  const plan = structuredClone(f.plan)
  plan.claims[0]!.type = 'calculation'
  plan.claims[0]!.calculation = {
    formula: 'revenue / base - 1',
    inputs: ['revenue', 'base'],
    reproduction: {
      bindings: [
        { name: 'revenue', inputIndex: 0, value: 120, sourceId: 'source' },
        { name: 'base', inputIndex: 1, value: 100, sourceId: 'source' },
      ],
      expected: 0.2,
    },
  }
  const deck = structuredClone(f.deck)
  deck.claims = presentationPlanClaims(plan)
  await f.call('save_plan', { expectedRevision: 1, plan })
  await f.call('production_begin', { requestId: 'calculation-run', planRevision: 2, deck })
  const precheck = await f.call('production_content_check', {
    requestId: 'calculation-run',
    pageId: deck.slides[0]!.id,
  })
  expect(precheck.report.findings).not.toContainEqual({
    code: 'calculation_not_reproduced',
    claimId: 'source-1',
  })
  expect(precheck.report.checks.calculations).toBe('not_verified')
  const delivery = await f.call('production_delivery_report', { requestId: 'calculation-run' })
  expect(delivery.pages[0].calculations[0]).toMatchObject({
    claimId: 'source-1',
    status: 'reproduced',
    scope: 'arithmetic_only',
  })
})
it('requires all five fields, rejects unknown fields and enforces page, request and document binding', async () => {
  const f = await setup()
  const before = files(f.userDataPath)
  for (const key of Object.keys(f.request)) {
    const request: Record<string, unknown> = { ...f.request }
    delete request[key]
    expect(decode(await f.service(request, new AbortController().signal))).toEqual({
      error: 'invalid_request',
    })
  }
  for (const [extra, error] of [
    [{ extra: true }, 'invalid_request'],
    [{ pageId: '../bad' }, 'invalid_request'],
    [{ requestId: '../bad' }, 'invalid_request'],
    [{ pageId: 'missing' }, 'not_found'],
    [{ requestId: 'missing' }, 'not_found'],
    [{ documentId: 'foreign' }, 'document_mismatch'],
  ] as const)
    expect(
      decode(await f.service({ ...f.request, ...extra }, new AbortController().signal)),
    ).toEqual({ error })
  const controller = new AbortController()
  controller.abort()
  expect(decode(await f.service(f.request, controller.signal))).toEqual({ error: 'aborted' })
  expect(files(f.userDataPath)).toEqual(before)
  expect(f.compile).not.toHaveBeenCalled()
})
it('checks changed content only on the derived target and preserves parent reports in any page state', async () => {
  const f = await setup()
  await f.call('production_run', { requestId: 'run' })
  const read = (requestId: string, pageId: string) =>
    f.call('production_content_check', { requestId, pageId })
  const original = await read('run', f.request.pageId)
  const slide = structuredClone(f.deck.slides[0]!)
  slide.elements = [
    { kind: 'text', id: 'claim', x: 1, y: 1, w: 8, h: 1, text: f.plan.claims[0]!.statement },
  ]
  expect(
    await f.call('production_rebuild_page', {
      parentRequestId: 'run',
      requestId: 'child',
      pageId: slide.id,
      slide,
    }),
  ).toMatchObject({ requestId: 'child' })
  const before = files(f.userDataPath)
  const child = await read('child', slide.id)
  expect(child.inputDigest).not.toBe(original.inputDigest)
  expect(child.planDigest).toBe(original.planDigest)
  expect(child.report.findings).not.toContainEqual({
    code: 'claim_text_not_found',
    claimId: 'source-1',
  })
  expect(original.report.findings).toContainEqual({
    code: 'claim_text_not_found',
    claimId: 'source-1',
  })
  expect(await read('run', slide.id)).toEqual(original)
  expect((await read('child', f.deck.slides[1]!.id)).report).toEqual(
    (await read('run', f.deck.slides[1]!.id)).report,
  )
  expect(files(f.userDataPath)).toEqual(before)
  expect(f.compile).toHaveBeenCalledTimes(8)
})
it.each(['building', 'failed'] as const)(
  'checks a %s page without changing its durable state',
  async (state) => {
    const f = await setup()
    const store = new PresentationStore(f.userDataPath)
    const building = store.updateProductionPage(
      store.production(f.deck.id, 'doc', 'run')!,
      f.request.pageId,
      { state: 'building', attempt: 1 },
    )
    if (state === 'failed')
      store.updateProductionPage(building, f.request.pageId, {
        state,
        attempt: 1,
        error: 'compile_failed',
      })
    const before = files(f.userDataPath)
    expect(decode(await f.service(f.request, new AbortController().signal))).toMatchObject({
      requestId: 'run',
      report: { pageId: f.request.pageId },
    })
    expect(files(f.userDataPath)).toEqual(before)
    expect(f.compile).not.toHaveBeenCalled()
  },
)
it('honors cancellation while waiting for the shared project lock', async () => {
  const f = await setup()
  let release!: () => void, entered!: () => void
  const waiting = new Promise<void>((resolve) => {
    release = resolve
  })
  const started = new Promise<void>((resolve) => {
    entered = resolve
  })
  f.compile.mockImplementation(async (input) => {
    entered()
    await waiting
    return compilePresentationDeck(input)
  })
  const running = f.call('production_run', { requestId: 'run' })
  await started
  const controller = new AbortController()
  const checking = f.service(f.request, controller.signal)
  controller.abort()
  release()
  expect(await running).toMatchObject({ status: 'compiled' })
  const before = files(f.userDataPath)
  expect(decode(await checking)).toEqual({ error: 'aborted' })
  expect(files(f.userDataPath)).toEqual(before)
  expect(f.compile).toHaveBeenCalledTimes(8)
})

it('keeps as-of findings bound to the frozen plan across updates and restart without writes', async () => {
  const f = await setup()
  const plan = structuredClone(f.plan)
  plan.claims[0]!.asOf = '2026-09-24'
  delete plan.sources[0]!.asOf
  await f.call('save_plan', { expectedRevision: 1, plan })
  await f.call('production_begin', { requestId: 'missing-date', planRevision: 2, deck: f.deck })
  plan.sources[0]!.asOf = '2025-12-31'
  await f.call('save_plan', { expectedRevision: 2, plan })
  await f.call('production_begin', { requestId: 'different-date', planRevision: 3, deck: f.deck })
  plan.sources[0]!.asOf = '2026-09-24'
  await f.call('save_plan', { expectedRevision: 3, plan })
  const before = files(f.userDataPath)
  const restarted = createPresentationService({ userDataPath: f.userDataPath, compile: f.compile })
  for (const [requestId, planRevision, code] of [
    ['missing-date', 2, 'source_as_of_missing'],
    ['different-date', 3, 'source_as_of_earlier'],
  ] as const) {
    const request = { ...f.request, requestId }
    const report = decode(await f.service(request, new AbortController().signal))
    expect(report.planRevision).toBe(planRevision)
    expect(report.report.findings).toContainEqual({ code, claimId: 'source-1', sourceId: 'source' })
    expect(report.report.checks.timeliness).toBe('not_verified')
    expect(decode(await restarted(request, new AbortController().signal))).toEqual(report)
  }
  expect(files(f.userDataPath)).toEqual(before)
  expect(f.compile).not.toHaveBeenCalled()
})
