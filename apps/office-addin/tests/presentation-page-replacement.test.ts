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
  let hostStatus: 'baseline' | 'staged' | 'conflict' | 'applied' | 'restore_staged' | 'undone' =
    'baseline'
  let active = artifact
  const receipts = new Map([['production/project/parent', receipt]])
  const write = vi.fn(
    async (
      next: PresentationPageReplacement,
      expected: PresentationPageReplacement | undefined,
    ) => {
      expect(journal).toEqual(expected)
      journal = structuredClone(next)
      if (next.state === 'applied') {
        receipts.delete('production/project/parent')
        receipts.set('production/project/child', next.childReceipt! as typeof receipt)
      }
      if (next.state === 'undone') {
        receipts.delete('production/project/child')
        const restored = structuredClone(next.parentReceipt!) as typeof receipt
        restored.slideIds[0] = next.restoredSlideId!
        restored.checkpoint.completed[0].slideId = next.restoredSlideId!
        receipts.set('production/project/parent', restored)
      }
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
    commit: vi.fn(async (_r: PresentationPageReplacement, current: () => Promise<void>) => {
      await current()
      expect(journal?.state).toBe('commit_pending')
      hostStatus = 'applied'
    }),
    undo: vi.fn(
      async (
        _r: PresentationPageReplacement,
        _b: string,
        restored: (id: string) => Promise<void>,
        current: () => Promise<void>,
      ) => {
        await current()
        if (_r.state === 'undo_pending') {
          hostStatus = 'restore_staged'
          await restored('restored')
        }
        expect(journal?.state).toBe('restore_inserted')
        await current()
        hostStatus = 'undone'
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
    artifact: () => active,
    readReceipt: (key: string) => {
      const value = receipts.get(key)
      if (!value) throw new Error('presentation_import_superseded')
      return value
    },
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
    receipts,
    activateChild: () => {
      active = { ...artifact, requestId: 'child' }
    },
    activateParent: () => {
      active = artifact
    },
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
it.each(['pending', 'inserted', 'staged'] as const)(
  'reports a missing original separately from the %s checkpoint',
  async (state) => {
    const f = setup()
    await f.skill.executeTool(f.stage)
    await f.confirm()
    const record = { ...f.journal()!, state }
    if (state === 'pending') delete record.newSlideId
    f.setJournal(record)
    f.adapter.inspect.mockResolvedValueOnce({ status: 'conflict', slideIds: ['original', 'new'] })
    const result = await f.call('inspect')
    expect(result.isError).not.toBe(true)
    expect(JSON.parse(result.output)).toMatchObject({
      state,
      stateSource: 'persisted_checkpoint',
      inspection: { status: 'conflict', slideIds: ['original', 'new'] },
      originalRetained: false,
      originalContentVerified: false,
    })
    expect(f.adapter.discard).not.toHaveBeenCalled()
    expect(f.journal()).toEqual(record)
  },
)

it('commits with a frozen complete child mapping then undoes from the original backup', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  expect(await f.call('commit')).toMatchObject({
    output: expect.stringContaining('awaiting_confirmation'),
  })
  expect(f.adapter.commit).not.toHaveBeenCalled()
  expect(await f.confirm()).toEqual({ status: 'confirmed' })
  expect(f.journal()).toMatchObject({ state: 'applied', childReceipt: { slideIds: ['new'] } })
  expect(f.receipts.has('production/project/parent')).toBe(false)
  expect(await f.call('undo')).toMatchObject({ isError: true })
  f.activateChild()
  expect(await f.call('undo')).toMatchObject({
    output: expect.stringContaining('awaiting_confirmation'),
  })
  expect(await f.confirm()).toEqual({ status: 'confirmed' })
  expect(f.journal()).toMatchObject({ state: 'undone', restoredSlideId: 'restored' })
  expect(f.receipts.get('production/project/parent')?.slideIds).toEqual(['restored'])
  expect(f.receipts.has('production/project/child')).toBe(false)
  expect(f.write.mock.calls.map((c) => c[0].state)).toEqual([
    'pending',
    'inserted',
    'staged',
    'commit_pending',
    'applied',
    'undo_pending',
    'restore_inserted',
    'undone',
  ])
})
it('refuses commit when downloaded child bytes no longer match the staged page', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  f.options.request.mockImplementation(async (body) =>
    (body as { operation: string }).operation === 'production_status'
      ? Response.json(f.status)
      : Response.json({
          projectId: 'project',
          requestId: 'child',
          pageId: 'page',
          planRevision: 1,
          status: 'compiled',
          pptxBase64: Buffer.from('PK\x03\x04changed').toString('base64'),
          sourceSlideId: '256#',
          report: { deckId: 'project', slideCount: 1 },
        }),
  )
  expect(await f.call('commit')).toMatchObject({ isError: true })
  expect(f.adapter.commit).not.toHaveBeenCalled()
})
it('reloads and rejects changed backup before commit or undo writes', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  f.metadata.sha256 = 'c'.repeat(64)
  expect(await f.call('commit')).toMatchObject({ isError: true })
  expect(f.adapter.commit).not.toHaveBeenCalled()
})

it('recovers commit_pending after host deletion with the same frozen receipts', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  f.adapter.commit.mockImplementationOnce(async () => {
    f.setHost('applied')
    throw new Error('office_state_uncertain')
  })
  await f.call('commit')
  expect(await f.confirm()).toMatchObject({ status: 'failed' })
  expect(f.journal()?.state).toBe('commit_pending')
  expect(await f.call('commit')).toMatchObject({
    output: expect.stringContaining('awaiting_confirmation'),
  })
  expect(await f.confirm()).toEqual({ status: 'confirmed' })
  expect(f.journal()?.state).toBe('applied')
})
it('recovers a known restored original without requesting another insertion', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  await f.call('commit')
  await f.confirm()
  f.activateChild()
  f.adapter.undo.mockImplementationOnce(async (_r, _b, restored) => {
    f.setHost('restore_staged')
    await restored('restored')
    throw new Error('office_state_uncertain')
  })
  await f.call('undo')
  expect(await f.confirm()).toMatchObject({ status: 'failed' })
  expect(f.journal()?.state).toBe('restore_inserted')
  expect(await f.call('undo')).toMatchObject({
    output: expect.stringContaining('awaiting_confirmation'),
  })
  expect(await f.confirm()).toEqual({ status: 'confirmed' })
  expect(f.adapter.undo.mock.calls[1][0]).toMatchObject({
    state: 'restore_inserted',
    restoredSlideId: 'restored',
  })
  expect(f.journal()?.state).toBe('undone')
})
it.each(['backup', 'receipt', 'journal', 'clear', 'cancel'] as const)(
  'rejects commit confirmation after %s drift',
  async (kind) => {
    const f = setup()
    await f.skill.executeTool(f.stage)
    await f.confirm()
    await f.call('commit')
    if (kind === 'backup') f.metadata.sha256 = 'c'.repeat(64)
    if (kind === 'receipt') f.receipt.slideIds[0] = 'other'
    if (kind === 'journal') f.setJournal({ ...f.journal()!, changeId: 'other' })
    if (kind === 'clear') f.skill.clear()
    if (kind === 'cancel') {
      f.proposals.reject()
      expect(f.adapter.commit).not.toHaveBeenCalled()
      return
    }
    expect(await f.confirm()).toMatchObject({ status: 'failed' })
    expect(f.adapter.commit).not.toHaveBeenCalled()
  },
)
it('requires the exact expected receipt overlay after a terminal save', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  await f.call('commit')
  const write = f.write.getMockImplementation()!
  f.write.mockImplementation(async (next, expected) => {
    await write(next, expected)
    if (next.state === 'applied')
      f.receipts.get('production/project/child')!.slideIds[0] = 'foreign'
  })
  expect(await f.confirm()).toMatchObject({ status: 'failed' })
  expect(f.journal()?.state).toBe('applied')
  expect(f.adapter.commit).toHaveBeenCalledTimes(1)
})
it('never deletes the original if saving commit_pending fails', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  await f.call('commit')
  f.write.mockRejectedValueOnce(new Error('office_write_failed'))
  expect(await f.confirm()).toMatchObject({ status: 'failed' })
  expect(f.adapter.commit).not.toHaveBeenCalled()
  expect(f.journal()?.state).toBe('staged')
})
it('rechecks original backup bytes after undo confirmation before restoring', async () => {
  const f = setup()
  await f.skill.executeTool(f.stage)
  await f.confirm()
  await f.call('commit')
  await f.confirm()
  f.activateChild()
  await f.call('undo')
  f.metadata.sha256 = 'c'.repeat(64)
  expect(await f.confirm()).toMatchObject({ status: 'failed' })
  expect(f.adapter.undo).not.toHaveBeenCalled()
  expect(f.journal()?.state).toBe('applied')
})
