import { createHash } from 'node:crypto'
import { expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createPresentationPageReplacementSkill } from '../src/skills/powerpoint/presentation-page-replacement.js'
import { presentationArtifactContent } from '../src/skills/powerpoint/presentation-page-delivery.js'
import type { PresentationPageReplacement } from '../src/skills/powerpoint/presentation-page-replacement-record.js'
vi.mock('../src/skills/powerpoint/powerpoint-package.js', () => ({
  presentationPackageDigest: async (base64: string) =>
    createHash('sha256').update(base64).digest('hex'),
}))
function setup() {
  const base64 = Buffer.from('PK\x03\x04revision').toString('base64')
  const backupBase64 = Buffer.from('original').toString('base64')
  const artifact = {
    documentId: 'doc',
    projectId: 'project',
    requestId: 'parent',
    pptxBase64: '',
    slideCount: 1,
    planRevision: 1,
    pages: [{ id: 'page', title: 'Page', sourceSlideId: '256#' }],
    pagePptxBase64: [base64],
  }
  const receipt = {
    state: 'complete' as const,
    documentId: 'doc',
    slideIds: ['host'],
    checkpoint: {
      version: 2 as const,
      artifactDigest: createHash('sha256')
        .update(presentationArtifactContent(artifact))
        .digest('hex'),
      sourceSlideIds: ['256#'],
      pageIds: ['page'],
      baselineSlideIds: ['original'],
      completed: [{ sourceSlideId: '256#', slideId: 'host' }],
    },
  }
  const status = {
    projectId: 'project',
    requestId: 'child',
    planRevision: 1,
    status: 'compiled',
    compiledCount: 1,
    total: 1,
    pages: [{ id: 'page', title: 'Page', state: 'compiled', attempt: 1 }],
    revision: { parentRequestId: 'parent', pageId: 'page', parentInputDigest: 'a'.repeat(64) },
  }
  const metadata = {
    backupId: 'backup',
    projectId: 'project',
    documentId: 'doc',
    requestId: 'child',
    parentRequestId: 'parent',
    pageId: 'page',
    hostSlideId: 'host',
    slideIds: ['original', 'host'],
    sha256: createHash('sha256').update(Buffer.from(backupBase64, 'base64')).digest('hex'),
    sizeBytes: 8,
    parentInputDigest: 'a'.repeat(64),
    inputDigest: 'b'.repeat(64),
    status: 'ready' as const,
    receivedBytes: 8,
  }
  let journal: PresentationPageReplacement | undefined
  let hostStatus: 'baseline' | 'staged' | 'conflict' = 'baseline'
  const write = vi.fn(
    async (
      next: PresentationPageReplacement,
      expected: PresentationPageReplacement | undefined,
    ) => {
      expect(journal).toEqual(expected)
      journal = structuredClone(next)
    },
  )
  const adapter = {
    inspect: vi.fn(async () => ({
      status: hostStatus,
      slideIds: hostStatus === 'staged' ? ['original', 'host', 'new'] : ['original', 'host'],
    })),
    stage: vi.fn(
      async (
        _r: PresentationPageReplacement,
        _b: string,
        inserted: (id: string) => Promise<void>,
        current: () => Promise<void>,
      ) => {
        await current()
        expect(journal?.state).toBe('pending')
        hostStatus = 'staged'
        await inserted('new')
      },
    ),
    discard: vi.fn(async (_r: PresentationPageReplacement, current: () => Promise<void>) => {
      await current()
      expect(journal?.state).toBe('discard_pending')
      hostStatus = 'baseline'
    }),
  }
  const proposals = createStructuredProposalController()
  const options = {
    available: () => true,
    documentId: vi.fn(async () => 'doc'),
    artifact: () => artifact,
    readReceipt: () => receipt,
    readPageReplacement: () => journal,
    writePageReplacement: write,
    loadBackup: vi.fn(async () => ({ metadata, base64: backupBase64 })),
    request: vi.fn(async (body: unknown) =>
      Response.json(
        (body as { operation: string }).operation === 'production_status'
          ? status
          : {
              projectId: 'project',
              requestId: 'child',
              pageId: 'page',
              planRevision: 1,
              status: 'compiled',
              pptxBase64: base64,
              sourceSlideId: '256#',
              report: { deckId: 'project', slideCount: 1 },
            },
      ),
    ),
    adapter,
    proposals,
  }
  const skill = createPresentationPageReplacementSkill(options)
  const stage = {
    id: 'stage',
    name: 'stage_presentation_page_replacement',
    input: {
      project_id: 'project',
      request_id: 'child',
      page_id: 'page',
      backup_id: 'backup',
      change_id: 'change',
    },
  }
  const call = (action: string) =>
    skill.executeTool({
      id: action,
      name: `${action}_presentation_page_replacement`,
      input: { project_id: 'project', change_id: 'change' },
    })
  const confirm = async () => {
    const id = proposals.pending()!.id
    const result = proposals.waitForDecision(id)
    await proposals.confirm(id).catch(() => {})
    return result
  }
  return {
    options,
    skill,
    stage,
    call,
    confirm,
    adapter,
    write,
    receipt,
    status,
    metadata,
    proposals,
    journal: () => journal,
    setJournal: (value: PresentationPageReplacement) => {
      journal = value
    },
    setHost: (value: typeof hostStatus) => {
      hostStatus = value
    },
  }
}
it('confirms staging, saves pending before insertion and inserted before staged; preserves parent mapping', async () => {
  const f = setup()
  expect(await f.skill.executeTool(f.stage)).toMatchObject({
    output: expect.stringContaining('awaiting_confirmation'),
    mutated: false,
  })
  expect(f.adapter.stage).not.toHaveBeenCalled()
  expect(await f.confirm()).toEqual({ status: 'confirmed' })
  expect(f.write.mock.calls.map((c) => c[0].state)).toEqual(['pending', 'inserted', 'staged'])
  expect(f.journal()).toMatchObject({ oldSlideId: 'host', newSlideId: 'new' })
  expect(f.receipt.slideIds).toEqual(['host'])
  expect(f.proposals.pending()).toBeUndefined()
})
it('requires confirmation to discard only the owned staged page', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  await f.call('discard')
  expect(f.adapter.discard).not.toHaveBeenCalled()
  expect(await f.confirm()).toEqual({ status: 'confirmed' })
  expect(f.journal()?.state).toBe('discarded')
})
it('keeps inserted after proof failure and resumes by receipt only', async () => {
  const f = setup()
  f.adapter.stage.mockImplementationOnce(async (_r, _b, inserted) => {
    f.setHost('staged')
    await inserted('new')
    throw new Error('office_state_uncertain')
  })
  await f.skill.executeTool(f.stage)
  expect(await f.confirm()).toMatchObject({ status: 'failed' })
  expect(f.journal()?.state).toBe('inserted')
  await f.call('resume')
  expect(await f.confirm()).toEqual({ status: 'confirmed' })
  expect(f.journal()?.state).toBe('staged')
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
})
it('does not repeat an insertion when its inserted receipt failed', async () => {
  const f = setup()
  f.write
    .mockImplementationOnce(async (next) => f.setJournal(structuredClone(next)))
    .mockRejectedValueOnce(new Error('office_write_failed'))
  await f.skill.executeTool(f.stage)
  await f.confirm()
  expect(f.journal()?.state).toBe('pending')
  expect(await f.call('resume')).toMatchObject({ isError: true })
  expect(await f.skill.executeTool(f.stage)).toMatchObject({ isError: true })
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
})
it('finishes discard_pending after host deletion without another deletion', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  f.setJournal({ ...f.journal()!, state: 'discard_pending' })
  f.setHost('baseline')
  await f.call('discard')
  expect(await f.confirm()).toEqual({ status: 'confirmed' })
  expect(f.adapter.discard).not.toHaveBeenCalled()
  expect(f.journal()?.state).toBe('discarded')
})
it.each(['clear', 'document', 'order', 'receipt'])(
  'rejects confirmation after %s drift',
  async (kind) => {
    const f = setup()
    await f.skill.executeTool(f.stage)
    if (kind === 'clear') f.skill.clear()
    if (kind === 'document') f.options.documentId.mockResolvedValue('other')
    if (kind === 'order') f.setHost('conflict')
    if (kind === 'receipt') f.receipt.slideIds[0] = 'other'
    expect(await f.confirm()).toMatchObject({ status: 'failed' })
    expect(f.adapter.stage).not.toHaveBeenCalled()
    expect(f.journal()).toBeUndefined()
  },
)
it.each(['backup', 'revision', 'plan'])('rejects mismatched %s before proposing', async (kind) => {
  const f = setup()
  if (kind === 'backup') f.metadata.requestId = 'other'
  if (kind === 'revision') f.status.revision.parentRequestId = 'other'
  if (kind === 'plan') f.status.planRevision = 2
  expect(await f.skill.executeTool(f.stage)).toMatchObject({ isError: true })
  expect(f.proposals.pending()).toBeUndefined()
})
it('inspects an uncertain pending transaction without inserting or deleting', async () => {
  const f = setup()
  f.adapter.stage.mockRejectedValueOnce(new Error('office_state_uncertain'))
  await f.skill.executeTool(f.stage)
  await f.confirm()
  expect(await f.call('inspect')).toMatchObject({
    mutated: false,
    output: expect.stringContaining('pending'),
  })
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
  expect(f.adapter.discard).not.toHaveBeenCalled()
})
it('keeps inserted when post-insertion content proof conflicts', async () => {
  const f = setup()
  f.adapter.stage.mockImplementationOnce(async (_r, _b, inserted) => {
    await inserted('new')
    f.setHost('conflict')
  })
  await f.skill.executeTool(f.stage)
  expect(await f.confirm()).toMatchObject({ status: 'failed' })
  expect(f.journal()?.state).toBe('inserted')
  expect(await f.call('resume')).toMatchObject({ isError: true })
  expect(await f.call('discard')).toMatchObject({ isError: true })
  expect(f.adapter.discard).not.toHaveBeenCalled()
})
it('rejects backup bytes that do not match the durable digest', async () => {
  const f = setup()
  f.metadata.sha256 = 'c'.repeat(64)
  expect(await f.skill.executeTool(f.stage)).toMatchObject({ isError: true })
  expect(f.proposals.pending()).toBeUndefined()
})
it('blocks another project while an unresolved journal exists', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  f.setJournal({ ...f.journal()!, projectId: 'other' })
  expect(await f.skill.executeTool(f.stage)).toMatchObject({ isError: true })
  expect(f.adapter.stage).toHaveBeenCalledTimes(1)
})
it('does not delete a staged page whose content has changed', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  await f.call('discard')
  f.setHost('conflict')
  expect(await f.confirm()).toMatchObject({ status: 'failed' })
  expect(f.journal()?.state).toBe('staged')
  expect(f.adapter.discard).not.toHaveBeenCalled()
})
it('retains discard_pending when the host write fails and can retry after inspection', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  f.adapter.discard.mockRejectedValueOnce(new Error('office_state_uncertain'))
  await f.call('discard')
  expect(await f.confirm()).toMatchObject({ status: 'failed' })
  expect(f.journal()?.state).toBe('discard_pending')
  await f.call('discard')
  expect(await f.confirm()).toEqual({ status: 'confirmed' })
  expect(f.journal()?.state).toBe('discarded')
})
