import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import type { PresentationCompileReport } from '@wiswork/pptx-engine/presentation'
import { PresentationStore } from '@wiswork/project-store'
import { createPresentationService } from '../src/main/presentation-service.js'

vi.mock('@wiswork/pptx-engine/presentation-compiler', () => ({ compilePresentationDeck: vi.fn() }))

const report: PresentationCompileReport = {
  deckId: 'deck',
  slideCount: 1,
  elementCount: 1,
  geometry: [],
  checks: {
    structure: 'passed',
    geometry: 'passed',
    render: 'not_run',
    sources: 'not_verified',
    roundTrip: 'not_run',
  },
}
const input = {
  operation: 'compile',
  documentId: 'office:/opaque?document',
  requestId: 'first',
  deck: {
    version: 1,
    id: 'deck',
    title: 'One',
    style: { fontFace: 'Arial', background: 'FFFFFF', textColor: '111111', accentColor: '3366FF' },
    assets: [],
    claims: [],
    slides: [
      {
        id: 'slide',
        title: 'Title',
        elements: [{ id: 'text', kind: 'text', text: 'Hello', x: 1, y: 1, w: 4, h: 1 }],
      },
    ],
  },
}
const decode = (bytes: Uint8Array) => JSON.parse(Buffer.from(bytes).toString('utf8'))
const signal = () => new AbortController().signal
const root = () => mkdtempSync(join(tmpdir(), 'presentation-service-'))
const result = () => ({ bytes: new Uint8Array([1, 2, 3]), report })

describe('presentation service', () => {
  it('deduplicates concurrent calls and persists full result across recreation', async () => {
    const userDataPath = root()
    const compile = vi.fn(async () => result())
    const service = createPresentationService({ userDataPath, compile })
    const [a, b] = await Promise.all([service(input, signal()), service(input, signal())])
    expect(decode(a)).toEqual({
      projectId: 'deck',
      requestId: 'first',
      status: 'compiled',
      pptxBase64: 'AQID',
      report,
    })
    expect(decode(b)).toEqual(decode(a))
    const reload = createPresentationService({ userDataPath, compile })
    expect(decode(await reload(input, signal()))).toEqual(decode(a))
    expect(compile).toHaveBeenCalledTimes(1)
    expect(
      decode(
        await reload(
          { operation: 'get', projectId: 'deck', documentId: input.documentId },
          signal(),
        ),
      ),
    ).toEqual(decode(a))
    expect(
      decode(await reload({ ...input, deck: { ...input.deck, title: 'Changed' } }, signal())),
    ).toEqual({ error: 'request_conflict' })
  })
  it('preserves the last ready output on failed or cancelled compilation and can retry', async () => {
    const userDataPath = root()
    const compile = vi.fn(async () => result())
    const service = createPresentationService({ userDataPath, compile })
    const ready = decode(await service(input, signal()))
    compile.mockRejectedValueOnce(new Error('/secret/arbitrary/path'))
    const second = { ...input, requestId: 'second' }
    expect(decode(await service(second, signal()))).toEqual({ error: 'compile_failed' })
    const controller = new AbortController()
    compile.mockImplementationOnce(async () => {
      controller.abort()
      return result()
    })
    expect(decode(await service(second, controller.signal))).toEqual({ error: 'aborted' })
    expect(
      decode(
        await service(
          { operation: 'get', projectId: 'deck', documentId: input.documentId },
          signal(),
        ),
      ),
    ).toEqual(ready)
    expect(
      decode(await createPresentationService({ userDataPath, compile })(second, signal()))
        .requestId,
    ).toBe('second')
  })
  it('denies document mismatch, traversal and oversized output with bounded errors', async () => {
    const compile = vi.fn(async () => result())
    const service = createPresentationService({ userDataPath: root(), compile })
    expect(decode(await service({ ...input, outputPath: '/tmp/arbitrary' }, signal()))).toEqual({
      error: 'invalid_request',
    })
    expect(
      decode(
        await service(
          { operation: 'get', projectId: 'deck', documentId: input.documentId, requestId: 'first' },
          signal(),
        ),
      ),
    ).toEqual({ error: 'invalid_request' })
    expect(
      decode(await service({ ...input, deck: { ...input.deck, arbitraryCode: 'evil' } }, signal())),
    ).toEqual({ error: 'invalid_deck' })
    await service(input, signal())
    expect(decode(await service({ ...input, documentId: 'other' }, signal()))).toEqual({
      error: 'document_mismatch',
    })
    for (const projectId of ['../x', '/tmp/x', '..', 'a/b']) {
      expect(
        decode(await service({ operation: 'get', documentId: 'doc', projectId }, signal())),
      ).toEqual({ error: 'invalid_request' })
    }
    compile.mockResolvedValueOnce({ bytes: new Uint8Array(11 * 1024 * 1024), report })
    expect(decode(await service({ ...input, requestId: 'big' }, signal()))).toEqual({
      error: 'output_too_large',
    })
    expect(
      decode(
        await service({ operation: 'get', projectId: 'missing', documentId: 'doc' }, signal()),
      ),
    ).toEqual({ error: 'not_found' })
  })
})

describe('durable project recovery', () => {
  const query = { operation: 'status', projectId: 'deck', documentId: input.documentId }
  it('reports latest pending slides and resumes the explicit saved request after recreation', async () => {
    const userDataPath = root()
    const compile = vi.fn(async (_deck: unknown) => result())
    const service = createPresentationService({ userDataPath, compile })
    const first = decode(await service(input, signal()))
    const controller = new AbortController()
    compile.mockImplementationOnce(async () => {
      controller.abort()
      return result()
    })
    const second = {
      ...input,
      requestId: 'second',
      deck: {
        ...input.deck,
        title: 'Second',
        slides: [{ ...input.deck.slides[0], id: 'next', title: 'Next' }],
      },
    }
    expect(decode(await service(second, controller.signal))).toEqual({ error: 'aborted' })
    const reload = createPresentationService({ userDataPath, compile })
    expect(decode(await reload(query, signal()))).toEqual({
      projectId: 'deck',
      title: 'Second',
      status: 'pending',
      latestRequestId: 'second',
      latestCompiledRequestId: 'first',
      slideCount: 1,
      slides: [{ id: 'next', title: 'Next' }],
      history: [
        { requestId: 'second', sequence: 2, status: 'pending', slideCount: 1 },
        { requestId: 'first', sequence: 1, status: 'compiled', slideCount: 1 },
      ],
    })
    const resume = { ...query, operation: 'resume', requestId: 'second' }
    const [a, b] = await Promise.all([reload(resume, signal()), reload(resume, signal())])
    expect(decode(a).requestId).toBe('second')
    expect(decode(b)).toEqual(decode(a))
    expect(compile).toHaveBeenCalledTimes(3)
    expect(compile.mock.calls.at(-1)?.[0]).toEqual(second.deck)
    expect(decode(await reload(query, signal())).checks).toEqual(report.checks)
    expect(decode(await reload({ ...resume, requestId: 'first' }, signal()))).toEqual(first)
    expect(decode(await reload({ ...query, operation: 'get' }, signal()))).toEqual(decode(a))
  })
  it('bounds history and refuses unknown, extra, inherited and cross-document requests', async () => {
    const userDataPath = root()
    const store = new PresentationStore(userDataPath)
    for (let i = 0; i < 25; i++) store.begin('deck', input.documentId, `r${i}`, input.deck)
    const service = createPresentationService({ userDataPath })
    const status = decode(await service(query, signal()))
    expect(status.history).toHaveLength(20)
    expect(status.history.map((r: { sequence: number }) => r.sequence)).toEqual(
      Array.from({ length: 20 }, (_, i) => 25 - i),
    )
    expect(status).not.toHaveProperty('checks')
    expect(status).not.toHaveProperty('latestCompiledRequestId')
    for (const operation of ['status', 'resume']) {
      const request = {
        ...query,
        operation,
        ...(operation === 'resume' ? { requestId: 'r0' } : {}),
      }
      expect(decode(await service({ ...request, documentId: 'other' }, signal()))).toEqual({
        error: 'document_mismatch',
      })
      expect(decode(await service({ ...request, deck: input.deck }, signal()))).toEqual({
        error: 'invalid_request',
      })
      expect(decode(await service(Object.create(request), signal()))).toEqual({
        error: 'invalid_request',
      })
    }
    expect(
      decode(await service({ ...query, operation: 'resume', requestId: 'unknown' }, signal())),
    ).toEqual({ error: 'not_found' })
    store.begin('deck', input.documentId, 'invalid', { ...input.deck, extra: true })
    expect(
      decode(await service({ ...query, operation: 'resume', requestId: 'invalid' }, signal())),
    ).toEqual({ error: 'invalid_deck' })
  })
})

it('resume retains cancellation and output bounds without superseding newer completed work', async () => {
  const userDataPath = root()
  const store = new PresentationStore(userDataPath)
  store.begin('deck', input.documentId, 'old', input.deck)
  const compile = vi.fn(async () => result())
  const service = createPresentationService({ userDataPath, compile })
  const latest = decode(await service(input, signal()))
  const resume = {
    operation: 'resume',
    projectId: 'deck',
    documentId: input.documentId,
    requestId: 'old',
  }
  const controller = new AbortController()
  compile.mockImplementationOnce(async () => {
    controller.abort()
    return result()
  })
  expect(decode(await service(resume, controller.signal))).toEqual({ error: 'aborted' })
  expect(store.request('deck', input.documentId, 'old')?.status).toBe('pending')
  compile.mockResolvedValueOnce({ bytes: new Uint8Array(11 * 1024 * 1024), report })
  expect(decode(await service(resume, signal()))).toEqual({ error: 'output_too_large' })
  expect(store.request('deck', input.documentId, 'old')?.status).toBe('pending')
  expect(decode(await service(resume, signal())).requestId).toBe('old')
  expect(
    decode(
      await service(
        { operation: 'get', projectId: 'deck', documentId: input.documentId },
        signal(),
      ),
    ),
  ).toEqual(latest)
  expect(
    decode(
      await service(
        { operation: 'status', projectId: 'deck', documentId: input.documentId },
        signal(),
      ),
    ).latestRequestId,
  ).toBe('first')
})

const plan = () => ({
  version: 1,
  projectId: 'deck',
  title: 'One',
  brief: {
    objective: '说明研究结果',
    audience: '研究团队',
    language: 'zh-CN',
    minutes: 10,
    requiredContent: [],
    constraints: [],
  },
  sources: [],
  claims: [],
  style: input.deck.style,
  slides: [
    {
      id: 'slide',
      title: 'Title',
      purpose: '说明结论',
      claimIds: [],
      layout: 'content',
      requiredAssets: [],
      acceptanceCriteria: ['文字可编辑'],
    },
  ],
})
const planRequest = (value = plan(), expectedRevision = 0) => ({
  operation: 'save_plan',
  documentId: input.documentId,
  projectId: 'deck',
  expectedRevision,
  plan: value,
})
describe('durable presentation planning', () => {
  it('persists a brand palette and rejects off-brand SlideIR before compilation', async () => {
    const compile = vi.fn(async () => result())
    const service = createPresentationService({ userDataPath: root(), compile })
    const branded = { ...plan(), brandKit: {
      id: 'company', revision: 1, name: '公司品牌',
      allowedColors: ['FFFFFF', '111111', '3366FF'],
    } }
    expect(decode(await service(planRequest(branded), signal()))).toMatchObject({ revision: 1, plan: branded })
    const offBrand = { ...input, planRevision: 1, deck: { ...input.deck, slides: [{
      ...input.deck.slides[0]!, elements: [{ ...input.deck.slides[0]!.elements[0]!, color: 'FF0000' }],
    }] } }
    expect(decode(await service(offBrand, signal()))).toEqual({ error: 'plan_mismatch' })
    expect(compile).not.toHaveBeenCalled()
    expect(decode(await service({ ...input, planRevision: 1 }, signal()))).toMatchObject({ status: 'compiled' })
    expect(compile).toHaveBeenCalledTimes(1)
  })
  it('saves and reloads a plan before any compile and exposes a planned project', async () => {
    const userDataPath = root()
    const service = createPresentationService({ userDataPath, compile: vi.fn() })
    expect(decode(await service(planRequest(), signal()))).toMatchObject({
      projectId: 'deck',
      revision: 1,
      plan: plan(),
    })
    const reload = createPresentationService({ userDataPath, compile: vi.fn() })
    expect(
      decode(
        await reload(
          { operation: 'get_plan', projectId: 'deck', documentId: input.documentId },
          signal(),
        ),
      ),
    ).toMatchObject({ revision: 1, plan: plan() })
    expect(
      decode(
        await reload(
          { operation: 'status', projectId: 'deck', documentId: input.documentId },
          signal(),
        ),
      ),
    ).toMatchObject({
      status: 'planned',
      slideCount: 1,
      history: [],
      plan: { revision: 1, value: plan(), revisions: [
        { revision: 1, createdAt: expect.any(String), inputDigest: expect.any(String),
          snapshot: { sourceCount: 0, claimCount: 0, slideCount: 1,
            sourcesDigest: expect.any(String), claimsDigest: expect.any(String),
            slidesDigest: expect.any(String), styleDigest: expect.any(String) } },
      ] },
    })
    expect(decode(await service(planRequest({ ...plan(), title: 'Changed' }), signal()))).toEqual({
      error: 'revision_conflict',
    })
    expect(decode(await service({ ...planRequest(), documentId: 'another' }, signal()))).toEqual({
      error: 'document_mismatch',
    })
  })
  it('requires a matching plan revision/contract and retains the original snapshot for retries', async () => {
    const userDataPath = root()
    const compile = vi.fn(async () => result())
    const service = createPresentationService({ userDataPath, compile })
    await service(planRequest(), signal())
    expect(decode(await service(input, signal()))).toEqual({ error: 'revision_conflict' })
    expect(
      decode(
        await service(
          { ...input, planRevision: 1, deck: { ...input.deck, title: 'Not planned' } },
          signal(),
        ),
      ),
    ).toEqual({ error: 'plan_mismatch' })
    const first = decode(await service({ ...input, planRevision: 1 }, signal()))
    expect(first.status).toBe('compiled')
    const nextPlan = { ...plan(), brief: { ...plan().brief, objective: '修订计划' } }
    expect(decode(await service(planRequest(nextPlan, 1), signal()))).toMatchObject({ revision: 2 })
    expect(decode(await service({ ...input, planRevision: 1 }, signal()))).toEqual(first)
    expect(decode(await service({ ...input, planRevision: 2 }, signal()))).toEqual({
      error: 'request_conflict',
    })
    const stored = new PresentationStore(userDataPath).request('deck', input.documentId, 'first')!
    expect(stored.plan).toMatchObject({ revision: 1, plan: plan() })
    expect(compile).toHaveBeenCalledTimes(1)
    expect(
      decode(await service({ ...input, requestId: 'second', planRevision: 1 }, signal())),
    ).toEqual({ error: 'revision_conflict' })
  })
  it('validates save/get fields and rejects a mismatched plan before creating a project', async () => {
    const service = createPresentationService({ userDataPath: root(), compile: vi.fn() })
    expect(
      decode(
        await service({ ...planRequest(), plan: { ...plan(), projectId: 'other' } }, signal()),
      ),
    ).toEqual({ error: 'invalid_plan' })
    expect(
      decode(await service({ ...planRequest(), plan: { ...plan(), injected: true } }, signal())),
    ).toEqual({ error: 'invalid_plan' })
    expect(decode(await service({ ...planRequest(), unsafe: true }, signal()))).toEqual({
      error: 'invalid_request',
    })
    expect(decode(await service({ ...planRequest(), expectedRevision: -1 }, signal()))).toEqual({
      error: 'invalid_request',
    })
    expect(
      decode(
        await service(
          { operation: 'get_plan', projectId: 'deck', documentId: input.documentId },
          signal(),
        ),
      ),
    ).toEqual({ error: 'not_found' })
  })
})

it('persists ordered source page metadata and restores it after service restart', async () => {
  const userDataPath = root()
  const deck = {
    ...input.deck,
    slides: [
      input.deck.slides[0]!,
      { ...input.deck.slides[0]!, id: 'second-slide', title: 'Second' },
    ],
  }
  const compile = vi.fn(async () => ({ ...result(), sourceSlideIds: ['300#', '257#'] }))
  const service = createPresentationService({ userDataPath, compile })
  const response = decode(await service({ ...input, deck }, signal()))
  const pages = [
    { id: 'slide', title: 'Title', sourceSlideId: '300#' },
    { id: 'second-slide', title: 'Second', sourceSlideId: '257#' },
  ]
  expect(response.pages).toEqual(pages)
  const reload = createPresentationService({ userDataPath, compile })
  expect(
    decode(
      await reload({ operation: 'get', projectId: 'deck', documentId: input.documentId }, signal()),
    ).pages,
  ).toEqual(pages)
  expect(
    decode(
      await reload(
        {
          operation: 'resume',
          projectId: 'deck',
          documentId: input.documentId,
          requestId: 'first',
        },
        signal(),
      ),
    ).pages,
  ).toEqual(pages)
  expect(compile).toHaveBeenCalledTimes(1)
})
it.each(
  [[], ['255#'], ['4294967296#'], ['256'], ['0256#'], ['NaN#'], ['256#', '256#']].map(
    (sourceSlideIds) => ({ sourceSlideIds }),
  ),
)(
  'rejects malformed compiler source IDs $sourceSlideIds before persisting a completed result',
  async ({ sourceSlideIds }) => {
    const userDataPath = root()
    const compile = vi.fn(async () => ({ ...result(), sourceSlideIds }))
    const service = createPresentationService({ userDataPath, compile })
    const request =
      sourceSlideIds.length === 2
        ? {
            ...input,
            deck: {
              ...input.deck,
              slides: [input.deck.slides[0]!, { ...input.deck.slides[0]!, id: 'second-slide' }],
            },
          }
        : input
    expect(decode(await service(request, signal()))).toEqual({ error: 'compile_failed' })
    expect(
      new PresentationStore(userDataPath).request('deck', input.documentId, 'first')?.status,
    ).toBe('pending')
  },
)
