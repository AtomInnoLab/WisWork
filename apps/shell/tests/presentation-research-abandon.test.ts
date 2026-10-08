import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { createPresentationAttachmentService } from '../src/main/presentation-attachments'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationResearchService } from '../src/main/presentation-research'
import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import { researchDraft } from '../../../packages/project-store/tests/fixtures/presentation-research'
const roots: string[] = []
const held = vi.hoisted(() => ({
  gate: undefined as { wait: Promise<void>; entered: () => void } | undefined,
}))
vi.mock('../src/main/presentation-attachments', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/main/presentation-attachments')>()
  return {
    ...actual,
    createPresentationAttachmentService: (
      options: Parameters<typeof actual.createPresentationAttachmentService>[0],
    ) => {
      const service = actual.createPresentationAttachmentService(options)
      return async (request: Record<string, unknown>, signal: AbortSignal) => {
        if (request.operation === 'attachment_match_excerpt' && held.gate) {
          held.gate.entered()
          await held.gate.wait
        }
        return service(request, signal)
      }
    },
  }
})
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'research-abandon-pc-'))
  roots.push(root)
  const store = new PresentationResearchStore(root)
  let service = createPresentationService({ userDataPath: root })
  const call = async (
    operation: string,
    fields: Record<string, unknown> = {},
    signal = new AbortController().signal,
  ) =>
    JSON.parse(
      Buffer.from(
        await service(
          {
            operation,
            documentId: 'doc',
            ...(operation === 'research_capabilities' ? {} : { projectId: 'research' }),
            ...fields,
          },
          signal,
        ),
      ).toString(),
    )
  return {
    root,
    store,
    call,
    restart: () => {
      service = createPresentationService({ userDataPath: root })
    },
  }
}
it('ends actual orphan running after restart, restores terminal without replay, then cleans and builds new ID', async () => {
  const f = fixture(),
    draft = researchDraft()
  const begun = await f.store.begin('doc', 'research', 0, 'orphan', draft)
  f.restart()
  const input = {
    ledgerId: 'orphan',
    expectedDraftDigest: begun.record.draftDigest,
    expectedRevision: 1,
  }
  const ended = await f.call('research_abandon', input)
  expect(ended).toMatchObject({
    id: 'orphan',
    sequence: 1,
    state: 'failed',
    error: 'aborted',
    draft,
    draftDigest: begun.record.draftDigest,
  })
  expect(ended).not.toHaveProperty('sources')
  f.restart()
  expect(await f.call('research_abandon', input)).toEqual(ended)
  expect(await f.call('research_read', { ledgerId: 'orphan' })).toEqual(ended)
  const replay = await f.call('research_build', { ledgerId: 'orphan', expectedRevision: 0, draft })
  expect(replay.record).toEqual(ended)
  expect(replay.history.revision).toBe(2)
  const deleted = await f.call('research_delete', {
    ledgerId: 'orphan',
    deleteId: 'cleanup',
    expectedDraftDigest: ended.draftDigest,
    expectedRevision: 2,
  })
  expect(deleted).toHaveProperty('ledgerId', 'orphan')
  expect(await f.call('research_abandon', input)).toEqual({ error: 'record_deleted' })
  const next = await f.call('research_build', {
    ledgerId: 'new',
    expectedRevision: 3,
    draft,
    historyVersion: 2,
  })
  expect(next.record).toMatchObject({ id: 'new', sequence: 2, state: 'completed' })
})
it('strictly negotiates recovery flags while preserving both existing capability shapes', async () => {
  const f = fixture()
  expect(await f.call('research_capabilities')).toEqual({ version: 1, available: true })
  expect(await f.call('research_capabilities', { includeCleanup: true })).toEqual({
    version: 1,
    available: true,
    cleanupAvailable: true,
    historyVersions: [1, 2],
  })
  expect(
    await f.call('research_capabilities', { includeCleanup: true, includeRecovery: true }),
  ).toEqual({
    version: 1,
    available: true,
    cleanupAvailable: true,
    recoveryAvailable: true,
    historyVersions: [1, 2],
  })
  for (const fields of [
    { includeRecovery: true },
    { includeCleanup: true, includeRecovery: false },
    { includeCleanup: true, includeRecovery: undefined },
    { includeCleanup: false, includeRecovery: true },
    { unknown: true },
  ])
    expect(await f.call('research_capabilities', fields)).toEqual({ error: 'invalid_request' })
})
it('rejects bad identity, digest, stale CAS and cancelled recovery without changing the record', async () => {
  const f = fixture(),
    draft = researchDraft(),
    begun = await f.store.begin('doc', 'research', 0, 'orphan', draft)
  const input = {
    ledgerId: 'orphan',
    expectedDraftDigest: begun.record.draftDigest,
    expectedRevision: 1,
  }
  for (const [patch, error] of [
    [{ expectedRevision: 0 }, 'revision_conflict'],
    [{ expectedDraftDigest: 'f'.repeat(64) }, 'request_conflict'],
    [{ ledgerId: 'missing' }, 'not_found'],
    [{ expectedRevision: 0.5 }, 'invalid_request'],
    [{ extra: true }, 'invalid_request'],
  ] as const)
    expect(await f.call('research_abandon', { ...input, ...patch })).toEqual({ error })
  const aborted = new AbortController()
  aborted.abort()
  expect(await f.call('research_abandon', input, aborted.signal)).toEqual({ error: 'aborted' })
  expect(await f.call('research_read', { ledgerId: 'orphan' })).toEqual(begun.record)
})
it('waits behind actual active source matching and never overwrites a completed build', async () => {
  const f = fixture(),
    bytes = Buffer.from('收入增长仅是管理层预测。'),
    attachmentId = createHash('sha256').update(bytes).digest('hex')
  const attachments = createPresentationAttachmentService({ userDataPath: f.root })
  const attachment = (operation: string, fields: Record<string, unknown> = {}) =>
    attachments(
      { operation, documentId: 'doc', attachmentId, ...fields },
      new AbortController().signal,
    )
  await attachment('attachment_begin', {
    name: 'original.txt',
    sizeBytes: bytes.length,
    sha256: attachmentId,
  })
  await attachment('attachment_chunk', { offset: 0, base64: bytes.toString('base64') })
  await attachment('attachment_finish')
  let release!: () => void, entered!: () => void
  const wait = new Promise<void>((done) => {
      release = done
    }),
    started = new Promise<void>((done) => {
      entered = done
    })
  held.gate = { wait, entered }
  try {
    const draft = researchDraft(attachmentId)
    const build = f.call('research_build', { ledgerId: 'active', expectedRevision: 0, draft })
    await started
    const running = await f.store.read('doc', 'research', 'active')
    let abandonSettled = false
    const abandoning = f
      .call('research_abandon', {
        ledgerId: 'active',
        expectedDraftDigest: running.draftDigest,
        expectedRevision: 1,
      })
      .then((value) => {
        abandonSettled = true
        return value
      })
    await Promise.resolve()
    expect(abandonSettled).toBe(false)
    release()
    const completed = await build
    expect(completed.record).toMatchObject({
      state: 'completed',
      sources: [{ sourceId: 'original', status: 'found' }],
    })
    expect(await abandoning).toEqual({ error: 'record_not_running' })
    expect(await f.call('research_read', { ledgerId: 'active' })).toEqual(completed.record)
  } finally {
    release()
    held.gate = undefined
  }
})
it('refuses standalone abandon without the shared project lock', async () => {
  const f = fixture(),
    begun = await f.store.begin('doc', 'research', 0, 'orphan', researchDraft())
  const service = createPresentationResearchService({ userDataPath: f.root })
  await expect(
    service(
      {
        operation: 'research_abandon',
        documentId: 'doc',
        projectId: 'research',
        ledgerId: 'orphan',
        expectedRevision: 1,
        expectedDraftDigest: begun.record.draftDigest,
      },
      new AbortController().signal,
    ),
  ).rejects.toThrow('invalid_state')
  expect(await f.store.read('doc', 'research', 'orphan')).toEqual(begun.record)
})
