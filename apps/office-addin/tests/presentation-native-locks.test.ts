import { expect, it, vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { readPresentationNativeLocks } from '../src/skills/powerpoint/presentation-native-locks.js'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
import type { StructuredProposalController } from '../src/agent/proposal-controller.js'
import { createPresentationDocumentBinding } from '../src/skills/powerpoint/presentation-document.js'
import type { PresentationPageReplacement } from '../src/skills/powerpoint/presentation-page-replacement-record.js'

function fixture() {
  const plan = benchmarkPlan()
  plan.slides[0]!.locked = true
  let documentId = 'doc'
  let revision = 2
  const record = {
    state: 'complete' as const,
    documentId: 'doc',
    slideIds: ['host'],
    checkpoint: {
      version: 2 as const,
      artifactDigest: 'a'.repeat(64),
      pageIds: [plan.slides[0]!.id],
      sourceSlideIds: ['256#'],
      baselineSlideIds: [],
      completed: [{ sourceSlideId: '256#', slideId: 'host' }],
    },
  }
  const receipts = [{ key: `production/${plan.projectId}/old`, record }]
  const source = {
    version: 1,
    documentId: 'doc',
    projectId: plan.projectId,
    requestId: 'old',
    source: 'production',
    planRevision: 1,
    artifactDigest: 'a'.repeat(64),
    pages: [{ id: plan.slides[0]!.id, title: plan.slides[0]!.title, sourceSlideId: '256#' }],
  }
  const options = {
    available: () => true,
    documentId: async () => documentId,
    lastProject: () => plan.projectId,
    listReceipts: () => structuredClone(receipts),
    hostSlideIds: vi.fn(async () => ['host', 'manual']),
    request: vi.fn(
      async (body: unknown) =>
        new Response(
          JSON.stringify(
            (body as { operation: string }).operation === 'get_plan'
              ? { projectId: plan.projectId, revision, plan }
              : source,
          ),
        ),
    ),
    rememberProject: async () => {},
  }
  const input = {
    id: 'proposal',
    operation: 'edit_existing_presentation_text',
    toolName: 'edit_existing_presentation_text',
    title: 'Edit',
    preview: {},
    impact: { host: 'powerpoint', count: 1, targets: ['host'] },
    fingerprint: 'v1',
  }
  return {
    options,
    input,
    receipts,
    plan,
    source,
    changeDocument: () => {
      documentId = 'other'
    },
    nextRevision: () => {
      revision++
    },
  }
}

it('follows effective durable mappings through page replacement and undo after reopening', async () => {
  const f = fixture()
  const values = new Map<string, string>()
  const settings = {
    get: (key: string) => values.get(key),
    set: (key: string, value: string) => {
      values.set(key, value)
    },
    save: async () => {},
    location: () => 'file',
  }
  const create = () => createPresentationDocumentBinding(settings, () => 'doc')
  const binding = create()
  const documentId = await binding.documentId()
  f.options.documentId = async () => documentId
  f.source.documentId = documentId
  f.receipts[0]!.record.documentId = documentId
  const parentReceipt = f.receipts[0]!.record
  const childReceipt = structuredClone(parentReceipt)
  childReceipt.slideIds = ['new']
  childReceipt.checkpoint.completed[0]!.slideId = 'new'
  childReceipt.checkpoint.artifactDigest = 'b'.repeat(64)
  const pending: PresentationPageReplacement = {
    version: 1,
    changeId: 'change',
    documentId,
    projectId: f.plan.projectId,
    parentRequestId: 'old',
    requestId: 'new',
    pageId: f.plan.slides[0]!.id,
    backupId: 'backup',
    parentArtifactDigest: 'a'.repeat(64),
    backupDigest: 'c'.repeat(64),
    originalPackageDigest: 'd'.repeat(64),
    replacementPackageDigest: 'e'.repeat(64),
    sourceSlideId: '256#',
    oldSlideId: 'host',
    beforeSlideIds: ['host', 'manual'],
    state: 'pending',
  }
  await binding.writeReceipt(f.receipts[0]!.key, parentReceipt)
  const inserted = { ...pending, newSlideId: 'new', state: 'inserted' as const }
  const staged = { ...inserted, state: 'staged' as const }
  const commit = { ...staged, state: 'commit_pending' as const, parentReceipt, childReceipt }
  const applied = { ...commit, state: 'applied' as const }
  await binding.writePageReplacement(pending, undefined)
  await binding.writePageReplacement(inserted, pending)
  await binding.writePageReplacement(staged, inserted)
  await binding.writePageReplacement(commit, staged)
  await binding.writePageReplacement(applied, commit)
  const original = f.options.request
  f.options.request = vi.fn(async (body) =>
    (body as { requestId?: string }).requestId === 'new'
      ? new Response(
          JSON.stringify({ ...f.source, requestId: 'new', artifactDigest: 'b'.repeat(64) }),
        )
      : original(body),
  )
  f.options.listReceipts = () => create().listReceipts() as typeof f.receipts
  f.options.hostSlideIds.mockResolvedValue(['new', 'manual'])
  f.input.impact.targets = ['new']
  const read = () => readPresentationNativeLocks(f.options, f.input, new AbortController().signal)
  const replacement = await read()
  expect(replacement?.pages[0]!.slideIds).toEqual(['new'])
  const undo = { ...applied, state: 'undo_pending' as const }
  const restored = { ...undo, restoredSlideId: 'restored', state: 'restore_inserted' as const }
  await binding.writePageReplacement(undo, applied)
  await binding.writePageReplacement(restored, undo)
  await binding.writePageReplacement({ ...restored, state: 'undone' }, restored)
  f.options.hostSlideIds.mockResolvedValue(['restored', 'manual'])
  f.input.impact.targets = ['restored']
  const after = await read()
  expect(after?.pages[0]!.slideIds).toEqual(['restored'])
  expect(after?.token).not.toBe(replacement?.token)
})
it('reads current locks against exact historical sources, limits stable native scope and changes the approval token when locks change', async () => {
  const f = fixture()
  const read = () => readPresentationNativeLocks(f.options, f.input, new AbortController().signal)
  const first = await read()
  expect(first?.pages).toEqual([
    {
      projectId: f.plan.projectId,
      pageId: f.plan.slides[0]!.id,
      title: f.plan.slides[0]!.title,
      slideIds: ['host'],
    },
  ])
  f.input.impact.targets = ['manual']
  expect((await read())?.pages).toEqual([])
  f.input.impact.targets = ['host']
  delete f.plan.slides[0]!.locked
  f.nextRevision()
  const next = await read()
  expect(next?.pages).toEqual([])
  expect(next?.token).not.toBe(first?.token)
})
it('unknown scripts include all locked copies; append, backup release and local records require no lock lookup', async () => {
  const f = fixture()
  f.input.operation = f.input.toolName = 'execute_office_js'
  f.input.impact.targets = ['selection']
  expect(
    (await readPresentationNativeLocks(f.options, f.input, new AbortController().signal))?.pages,
  ).toHaveLength(1)
  f.options.request.mockClear()
  for (const name of ['import_presentation_production', 'release_slide_chart_values_change']) {
    f.input.operation = f.input.toolName = name
    expect(
      await readPresentationNativeLocks(f.options, f.input, new AbortController().signal),
    ).toBeUndefined()
  }
  f.input.impact.host = 'local_preference'
  expect(
    await readPresentationNativeLocks(f.options, f.input, new AbortController().signal),
  ).toBeUndefined()
  expect(f.options.request).not.toHaveBeenCalled()
})
it.each(['digest', 'document', 'receipt', 'plan', 'cancel'] as const)(
  'rejects %s changes during lock evidence reads',
  async (kind) => {
    const f = fixture()
    const controller = new AbortController()
    const original = f.options.request
    f.options.request = vi.fn(async (body) => {
      const response = await original(body)
      if ((body as { operation: string }).operation === 'read_import_source') {
        if (kind === 'digest')
          return new Response(JSON.stringify({ ...f.source, artifactDigest: 'b'.repeat(64) }))
        if (kind === 'document') f.changeDocument()
        if (kind === 'receipt') f.receipts[0]!.record.slideIds = ['changed']
        if (kind === 'plan') f.nextRevision()
        if (kind === 'cancel') controller.abort()
      }
      return response
    })
    await expect(
      readPresentationNativeLocks(f.options, f.input, controller.signal),
    ).rejects.toThrow()
  },
)
it.each(['unchanged', 'changed', 'during_qa'] as const)(
  'the actual runtime rechecks locked pages after QA persistence: %s',
  async (kind) => {
    const f = fixture()
    vi.stubGlobal('Office', {
      context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
    })
    vi.stubGlobal('PowerPoint', {
      run: async (callback: (context: unknown) => Promise<unknown>) =>
        callback({
          presentation: { slides: { load: () => {}, items: [{ id: 'host' }, { id: 'manual' }] } },
          sync: async () => {},
        }),
    })
    const runtime = createOfficeHostRuntime('powerpoint', {
      presentation: {
        ...f.options,
        invalidateQa: async () => {
          if (kind === 'during_qa') f.nextRevision()
        },
      },
    })
    const proposals = runtime.proposals as StructuredProposalController
    const execute = vi.fn(async () => {})
    try {
      const proposal = proposals.propose({ ...f.input, validate: async () => true, execute })
      await vi.waitFor(() =>
        expect(proposals.pending()?.lockReview).toMatchObject({
          state: 'ready',
          pages: [{ pageId: f.plan.slides[0]!.id }],
        }),
      )
      if (kind === 'changed') f.nextRevision()
      const decision = proposals.waitForDecision(proposal.id)
      if (kind === 'unchanged') {
        await proposals.confirm(proposal.id)
        expect(execute).toHaveBeenCalledOnce()
        expect(f.plan.slides[0]!.locked).toBe(true)
        await expect(decision).resolves.toEqual({ status: 'confirmed' })
      } else {
        await expect(proposals.confirm(proposal.id)).rejects.toThrow(
          'presentation_lock_review_stale',
        )
        expect(execute).not.toHaveBeenCalled()
        await expect(decision).resolves.toEqual({
          status: 'failed',
          error: 'presentation_lock_review_stale',
        })
      }
    } finally {
      runtime.dispose()
      vi.unstubAllGlobals()
    }
  },
)
