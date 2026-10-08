import { expect, it, vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
import type { PresentationImportRecord } from '../src/skills/powerpoint/presentation-delivery.js'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'

function fixture() {
  const plan = benchmarkPlan()
  let documentId = 'doc'
  const record: PresentationImportRecord = {
    state: 'complete',
    documentId: 'doc',
    slideIds: ['host'],
    checkpoint: {
      version: 2,
      artifactDigest: 'a'.repeat(64),
      pageIds: [plan.slides[0]!.id],
      sourceSlideIds: ['256#'],
      baselineSlideIds: [],
      completed: [{ sourceSlideId: '256#', slideId: 'host' }],
    },
  }
  const receipts = [{ key: `production/${plan.projectId}/old`, record }]
  const hostSlideIds = vi.fn(async () => ['host'])
  const options = {
    request: async (body: unknown) =>
      new Response(
        JSON.stringify(
          (body as { operation: string }).operation === 'read_import_source'
            ? {
                version: 1,
                documentId: 'doc',
                projectId: plan.projectId,
                requestId: 'old',
                source: 'production',
                planRevision: 1,
                artifactDigest: 'a'.repeat(64),
                pages: [
                  { id: plan.slides[0]!.id, title: plan.slides[0]!.title, sourceSlideId: '256#' },
                ],
              }
            : {
                projectId: plan.projectId,
                title: plan.title,
                status: 'planned',
                slideCount: 8,
                slides: plan.slides.map(({ id, title }) => ({ id, title })),
                history: [],
                plan: { revision: 2, value: plan },
                productionTasks: [
                  {
                    requestId: 'old',
                    sequence: 1,
                    planRevision: 1,
                    status: 'compiled',
                    compiledCount: 1,
                    total: 1,
                  },
                ],
              },
        ),
      ),
    available: () => true,
    documentId: async () => documentId,
    lastProject: () => plan.projectId,
    executeTool: vi.fn(),
    listReceipts: () => structuredClone(receipts),
    hostSlideIds,
    rememberProject: async () => undefined,
  }
  const controller = createPresentationProjectController(options)
  return {
    options,
    controller,
    hostSlideIds,
    receipts,
    changeDocument: () => {
      documentId = 'other'
    },
  }
}
it('connects the real host runtime to the Office.js slide snapshot without using supplied cached host IDs', async () => {
  const f = fixture()
  const load = vi.fn()
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
  })
  vi.stubGlobal('PowerPoint', {
    run: async (callback: (context: unknown) => Promise<unknown>) =>
      callback({
        presentation: { slides: { load, items: [{ id: 'manual-only' }] } },
        sync: async () => undefined,
      }),
  })
  const runtime = createOfficeHostRuntime('powerpoint', { presentation: f.options })
  try {
    await runtime.presentation!.refresh()
    expect(load).toHaveBeenCalledWith('items/id')
    expect(f.hostSlideIds).not.toHaveBeenCalled()
    expect(
      runtime.presentation!.snapshot().project?.hostAssociations?.pages[0]!.hostPages[0]!.presence,
    ).toBe('missing')
  } finally {
    runtime.presentation!.clear()
    vi.unstubAllGlobals()
  }
})
it('refreshes durable host association and keeps the project usable when host inspection fails', async () => {
  const f = fixture()
  await f.controller.refresh()
  expect(f.controller.snapshot().project?.hostAssociations?.pages[0]!.hostPages[0]).toMatchObject({
    slideId: 'host',
    sourceProof: 'digest',
    revisionRelation: 'historical',
    presence: 'present',
  })
  f.hostSlideIds.mockRejectedValue(new Error('host unavailable'))
  await f.controller.refresh()
  expect(f.controller.snapshot().project?.plan?.revision).toBe(2)
  expect(f.controller.snapshot().project?.hostAssociations).toBeUndefined()
  expect(f.controller.snapshot().project?.hostAssociationsUnavailable).toBe(true)
})
it('keeps unavailable sources unverified and refuses a descriptor from another document', async () => {
  for (const unavailable of [true, false]) {
    const f = fixture()
    const request = f.options.request
    f.options.request = async (body) => {
      if ((body as { operation: string }).operation !== 'read_import_source') return request(body)
      if (unavailable) return new Response(JSON.stringify({ error: 'invalid_request' }))
      const value = await (await request(body)).json()
      return new Response(JSON.stringify({ ...value, documentId: 'foreign' }))
    }
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.plan?.revision).toBe(2)
    if (unavailable) {
      expect(f.controller.snapshot().project?.hostAssociations?.unverifiedSources).toBe(1)
      expect(
        f.controller.snapshot().project?.hostAssociations?.pages[0]!.hostPages[0]!.sourceProof,
      ).toBe('unverified')
    } else {
      expect(f.controller.snapshot().project?.hostAssociations).toBeUndefined()
      expect(f.controller.snapshot().project?.hostAssociationsUnavailable).toBe(true)
    }
  }
})
it('reads current host presence after slow source metadata instead of publishing old page positions', async () => {
  const f = fixture()
  const request = f.options.request
  let hostIds = ['host']
  f.hostSlideIds.mockImplementation(async () => [...hostIds])
  let release!: (response: Response) => void
  let sourceRequest: unknown
  f.options.request = async (body) => {
    if ((body as { operation: string }).operation !== 'read_import_source') return request(body)
    sourceRequest = body
    return new Promise((resolve) => {
      release = resolve
    })
  }
  const pending = f.controller.refresh()
  await vi.waitFor(() => expect(sourceRequest).toBeDefined())
  hostIds = []
  release(await request(sourceRequest))
  await pending
  expect(f.controller.snapshot().project?.hostAssociations?.pages[0]!.hostPages[0]!.presence).toBe(
    'missing',
  )
})
it('rechecks identity and receipts after the source request, and suppresses cancelled source results', async () => {
  for (const action of ['document', 'cancel', 'receipt']) {
    const f = fixture()
    const request = f.options.request
    let release!: (response: Response) => void
    let sourceRequest: unknown
    f.options.request = async (body) => {
      if ((body as { operation: string }).operation !== 'read_import_source') return request(body)
      sourceRequest = body
      return new Promise((resolve) => {
        release = resolve
      })
    }
    const pending = f.controller.refresh()
    await vi.waitFor(() => expect(sourceRequest).toBeDefined())
    if (action === 'document') f.changeDocument()
    if (action === 'cancel') f.controller.cancel()
    if (action === 'receipt') f.receipts[0]!.record.checkpoint!.artifactDigest = 'b'.repeat(64)
    release(await request(sourceRequest))
    await pending
    expect(f.controller.snapshot().project?.hostAssociations).toBeUndefined()
    if (action === 'receipt')
      expect(f.controller.snapshot().project?.hostAssociationsUnavailable).toBe(true)
    else expect(f.controller.snapshot().project).toBeUndefined()
  }
})
it('does not publish associations after document switching, cancellation or a receipt change during inspection', async () => {
  for (const action of ['document', 'cancel', 'receipt']) {
    const f = fixture()
    let release!: (ids: string[]) => void
    f.hostSlideIds.mockImplementation(
      () =>
        new Promise((resolve) => {
          release = resolve
        }),
    )
    const pending = f.controller.refresh()
    await vi.waitFor(() => expect(f.hostSlideIds).toHaveBeenCalled())
    if (action === 'document') f.changeDocument()
    if (action === 'cancel') f.controller.cancel()
    if (action === 'receipt') f.receipts[0]!.record.checkpoint!.artifactDigest = 'b'.repeat(64)
    release(['host'])
    await pending
    expect(f.controller.snapshot().project?.hostAssociations).toBeUndefined()
    if (action === 'receipt')
      expect(f.controller.snapshot().project?.hostAssociationsUnavailable).toBe(true)
    else expect(f.controller.snapshot().project).toBeUndefined()
  }
})
