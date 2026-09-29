import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, readdirSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationLifecycleStore } from '@wiswork/project-store'
import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import { researchDraft } from '../../../packages/project-store/tests/fixtures/presentation-research'
import { createPresentationResearchService } from '../src/main/presentation-research'
import {
  capturePresentationProjectCreationLease,
  capturePresentationProjectWriteLease,
  capturePresentationProjectAsyncReadLease,
} from '../src/main/presentation-project-write-lease'
import { stopPresentationProjectWork } from '../src/main/presentation-project-work'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture(
  controlled = true,
  acquireProjectLock?: (projectId: string) => Promise<() => void>,
) {
  const root = mkdtempSync(join(tmpdir(), 'research-service-fence-'))
  roots.push(root)
  const scope = { projectId: 'p', documentId: 'doc' },
    lifecycle = new PresentationLifecycleStore(root),
    store = new PresentationResearchStore(root)
  if (controlled) lifecycle.initialize(scope)
  let entered!: () => void, finish!: (value: unknown) => void, fail!: (error: unknown) => void
  const arrived = new Promise<void>((r) => {
      entered = r
    }),
    evidence = new Promise<unknown>((r, j) => {
      finish = r
      fail = j
    })
  const service = createPresentationResearchService({
    userDataPath: root,
    acquireProjectLock,
    attachments: async () => {
      entered()
      return evidence
    },
    captureProjectLease: ({ scope, operation, signal }) =>
      operation === 'research_build'
        ? capturePresentationProjectCreationLease({ store: lifecycle, scope, signal })
        : operation === 'research_abandon'
          ? capturePresentationProjectWriteLease({
              store: lifecycle,
              scope,
              signal,
              readExistingProject: () => undefined,
            })
          : capturePresentationProjectAsyncReadLease({
              store: lifecycle,
              scope,
              signal,
              readExistingProject: async (current) =>
                (await store.summary(current.documentId, current.projectId)).totalRecords
                  ? current
                  : undefined,
            }),
  })
  const request = {
    operation: 'research_build',
    ...scope,
    ledgerId: 'ledger',
    expectedRevision: 0,
    draft: researchDraft('a'.repeat(64)),
  }
  const freeze = () =>
    lifecycle.beginDeletion(scope, 0, {
      deletionId: 'delete',
      reason: 'user',
      resources: [{ resourceId: 'research', kind: 'research', ownership: 'project_exclusive' }],
    })
  return { root, scope, lifecycle, store, service, request, arrived, finish, fail, freeze }
}
it.each(['success', 'failure'])(
  'does not finish a research body when evidence returns after freeze (%s)',
  async (mode) => {
    const f = fixture(),
      pending = f.service(f.request, new AbortController().signal)
    await f.arrived
    const before = await f.store.read('doc', 'p', 'ledger')
    f.freeze()
    if (mode === 'success')
      f.finish({ attachmentId: 'a'.repeat(64), sha256: 'a'.repeat(64), status: 'uploading' })
    else f.fail(Error('source_unavailable'))
    await expect(pending).rejects.toThrow('revision_conflict')
    expect(await new PresentationResearchStore(f.root).read('doc', 'p', 'ledger')).toEqual(before)
  },
)
it('drains registered research across factory instances only after actual evidence work settles', async () => {
  const f = fixture(),
    pending = f.service(f.request, new AbortController().signal)
  await f.arrived
  f.freeze()
  let settled = false
  const drained = stopPresentationProjectWork({ root: f.root, ...f.scope }).then(() => {
    settled = true
  })
  await Promise.resolve()
  await Promise.resolve()
  expect(settled).toBe(false)
  f.finish({ attachmentId: 'a'.repeat(64), sha256: 'a'.repeat(64), status: 'uploading' })
  await expect(pending).rejects.toThrow('aborted')
  await drained
  expect((await f.store.read('doc', 'p', 'ledger')).state).toBe('running')
})
it('legacy pure reads leave absent research and lifecycle namespaces absent', async () => {
  const f = fixture(false),
    service = createPresentationResearchService({ userDataPath: f.root })
  expect(
    await service({ operation: 'research_list', ...f.scope }, new AbortController().signal),
  ).toMatchObject({ totalRecords: 0 })
  expect(readdirSync(f.root)).toEqual([])
})

it('owns validated request and draft aliases across project lock admission', async () => {
  let arrived!: () => void, unlock!: (release: () => void) => void
  const waiting = new Promise<void>((r) => {
      arrived = r
    }),
    lock = new Promise<() => void>((r) => {
      unlock = r
    })
  const f = fixture(true, async () => {
    arrived()
    return lock
  })
  const pending = f.service(f.request, new AbortController().signal)
  await waiting
  f.request.documentId = 'foreign'
  f.request.projectId = 'foreign'
  f.request.ledgerId = 'foreign'
  f.request.draft.scope = 'mutated'
  unlock(() => {})
  await f.arrived
  f.finish({ attachmentId: 'a'.repeat(64), sha256: 'a'.repeat(64), status: 'uploading' })
  const result = (await pending) as {
    record: { documentId: string; projectId: string; id: string; draft: { scope: string } }
  }
  expect(result.record).toMatchObject({ documentId: 'doc', projectId: 'p', id: 'ledger' })
  expect(result.record.draft.scope).not.toBe('mutated')
  expect((await f.store.summary('foreign', 'foreign')).totalRecords).toBe(0)
})
it('refuses a queued original revision after freeze without creating research content', async () => {
  let arrived!: () => void, unlock!: (release: () => void) => void
  const waiting = new Promise<void>((r) => {
      arrived = r
    }),
    lock = new Promise<() => void>((r) => {
      unlock = r
    })
  const f = fixture(true, async () => {
    arrived()
    return lock
  })
  const pending = f.service(f.request, new AbortController().signal)
  await waiting
  f.freeze()
  unlock(() => {})
  await expect(pending).rejects.toThrow('revision_conflict')
  expect(readdirSync(f.root)).toEqual(['presentation-project-lifecycles'])
})

it('guarded cancellation preserves the running last-known body until explicit abandon', async () => {
  const f = fixture(true, async () => () => {}),
    controller = new AbortController()
  const pending = f.service(f.request, controller.signal)
  await f.arrived
  const before = await f.store.read('doc', 'p', 'ledger')
  controller.abort()
  f.finish({ attachmentId: 'a'.repeat(64), sha256: 'a'.repeat(64), status: 'uploading' })
  await expect(pending).rejects.toThrow('aborted')
  expect(await f.store.read('doc', 'p', 'ledger')).toEqual(before)
  expect(
    await f.service(
      {
        operation: 'research_abandon',
        ...f.scope,
        ledgerId: 'ledger',
        expectedRevision: 1,
        expectedDraftDigest: before.draftDigest,
      },
      new AbortController().signal,
    ),
  ).toMatchObject({ state: 'failed', error: 'aborted' })
})
it('allows a guarded build before any ProjectStore plan and creates preserving control only', async () => {
  const f = fixture(false),
    pending = f.service(f.request, new AbortController().signal)
  await f.arrived
  f.finish({ attachmentId: 'a'.repeat(64), sha256: 'a'.repeat(64), status: 'uploading' })
  expect(await pending).toMatchObject({ record: { state: 'completed' } })
  expect(f.lifecycle.read(f.scope)?.policy).toEqual({
    contentRetentionDays: null,
    auditRetentionDays: null,
  })
  expect(readdirSync(f.root).sort()).toEqual([
    'presentation-project-lifecycles',
    'presentation-research',
  ])
})

it('preserves the full legacy 4096-character document scope through guarded pre-plan creation and reopen', async () => {
  const f = fixture(false),
    documentId = 'd'.repeat(4096)
  f.scope.documentId = documentId
  f.request.documentId = documentId
  const pending = f.service(f.request, new AbortController().signal)
  await f.arrived
  f.finish({ attachmentId: 'a'.repeat(64), sha256: 'a'.repeat(64), status: 'uploading' })
  expect(await pending).toMatchObject({
    record: { documentId, projectId: 'p', state: 'completed' },
  })
  expect(new PresentationLifecycleStore(f.root).read(f.scope)?.revision).toBe(0)
  expect(await new PresentationResearchStore(f.root).read(documentId, 'p', 'ledger')).toMatchObject(
    { documentId, state: 'completed' },
  )
  expect(() =>
    new PresentationLifecycleStore(f.root).read({ ...f.scope, documentId: 'e'.repeat(4096) }),
  ).toThrow('document_mismatch')
  await expect(
    f.service({ ...f.request, documentId: 'd'.repeat(4097) }, new AbortController().signal),
  ).rejects.toThrow('invalid_request')
})

it.each([
  { operation: 'research_read', ledgerId: '../foreign' },
  {
    operation: 'research_delete',
    ledgerId: 'ledger',
    deleteId: '../foreign',
    expectedRevision: 0,
    expectedDraftDigest: 'a'.repeat(64),
  },
  {
    operation: 'research_delete',
    ledgerId: 'ledger',
    deleteId: 'delete',
    expectedRevision: -1,
    expectedDraftDigest: 'a'.repeat(64),
  },
  {
    operation: 'research_abandon',
    ledgerId: 'ledger',
    expectedRevision: 0,
    expectedDraftDigest: 'invalid',
  },
  { operation: 'research_delete_status', deleteId: '../foreign' },
])('validates operation-specific fields before lifecycle admission: $operation', async (body) => {
  const root = mkdtempSync(join(tmpdir(), 'research-invalid-admission-'))
  roots.push(root)
  const service = createPresentationResearchService({
    userDataPath: root,
    captureProjectLease: () => {
      throw Error('admission_should_not_run')
    },
  })
  await expect(
    service({ ...body, documentId: 'doc', projectId: 'p' }, new AbortController().signal),
  ).rejects.toThrow('invalid_request')
  expect(readdirSync(root)).toEqual([])
})
