import { afterEach, expect, it, vi } from 'vitest'
import { createHash } from 'node:crypto'
import { existsSync, readFileSync, readdirSync, lstatSync } from 'node:fs'
import { join } from 'node:path'
import { PresentationLifecycleStore } from '@wiswork/project-store'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { researchDraft } from '../../../packages/project-store/tests/fixtures/presentation-research'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationAttachmentService } from '../src/main/presentation-attachments'
import { createPresentationResearchService } from '../src/main/presentation-research'
import { PresentationCommentLibrary } from '../src/main/presentation-comments'
import { PresentationPreferenceLibrary } from '../src/main/presentation-preferences'
import { PresentationManualObservationLibrary } from '../src/main/presentation-manual-observations'
import { inspectPresentationProjectInventory } from '../src/main/presentation-project-inventory'
import { createPresentationProjectDeletionService } from '../src/main/presentation-project-deletion'
import {
  deliveryBundleFixture,
  cleanupDeliveryBundleFixtures,
} from './helpers/delivery-bundle-fixture'
const fault = vi.hoisted(() => ({
  singleFile: false,
  projectMove: undefined as undefined | ((source: string) => void),
}))
vi.mock('node:fs', async (load) => {
  const actual = await load<typeof import('node:fs')>()
  return {
    ...actual,
    renameSync: (...args: Parameters<typeof actual.renameSync>) => {
      fault.projectMove?.(String(args[0]))
      return actual.renameSync(...args)
    },
    unlinkSync: (...args: Parameters<typeof actual.unlinkSync>) => {
      if (
        fault.singleFile &&
        String(args[0]).includes('presentation-project-deletion-work') &&
        String(args[0]).endsWith('/data')
      ) {
        fault.singleFile = false
        throw Error('synthetic-unlink-failure')
      }
      return actual.unlinkSync(...args)
    },
  }
})
afterEach(() => {
  fault.singleFile = false
  fault.projectMove = undefined
  cleanupDeliveryBundleFixtures()
})
const hash = (v: string | Uint8Array) => createHash('sha256').update(v).digest('hex')
const kinds = [
  'project',
  'research',
  'delivery_bundles',
  'page_backups',
  'preferences',
  'comments',
  'manual_observations',
] as const
function tree(path: string): Map<string, string> {
  const found = new Map<string, string>()
  const walk = (p: string) => {
    if (!existsSync(p)) return
    const stat = lstatSync(p)
    expect(stat.isSymbolicLink()).toBe(false)
    if (stat.isDirectory()) for (const name of readdirSync(p)) walk(join(p, name))
    else found.set(p, hash(readFileSync(p)))
  }
  walk(path)
  return found
}
async function fixture(shared = false) {
  const exported = await compilePresentationDeck(benchmarkPlannedDeck())
  const f = await deliveryBundleFixture((files) => {
      files.set('presentation.pptx', Buffer.from(exported.bytes))
    }),
    scope = { documentId: 'doc', projectId: f.base.projectId },
    plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck(),
    other = { documentId: 'doc', projectId: 'other-project' }
  f.store.savePlan(scope.projectId, scope.documentId, 0, plan)
  f.store.savePlan(other.projectId, other.documentId, 0, { ...plan, projectId: other.projectId })
  const comments = new PresentationCommentLibrary(f.root),
    preferences = new PresentationPreferenceLibrary(f.root),
    manual = new PresentationManualObservationLibrary(f.root)
  for (const s of [scope, other]) {
    const p = { ...plan, projectId: s.projectId }
    comments.add(
      s.documentId,
      s.projectId,
      0,
      1,
      {
        id: 'review',
        targetKind: 'slide',
        targetId: plan.slides[0]!.id,
        authorLabel: 'Reviewer',
        text: 'Keep exact original',
      },
      { revision: 1, plan: p },
    )
    preferences.save(s.documentId, {
      projectId: s.projectId,
      changeId: 'edit',
      text: 'Short titles',
    })
    manual.begin(s.documentId, s.projectId, 'observation', 'native-slide', {
      id: 'native-shape',
      name: 'Title',
      type: 'TextBox',
      left: 1,
      top: 2,
      width: 300,
      height: 40,
      text: 'Original',
    })
  }
  const attachments = createPresentationAttachmentService({ userDataPath: f.root }),
    research = createPresentationResearchService({ userDataPath: f.root, attachments })
  let attachmentId: string | undefined
  if (shared) {
    const raw = Buffer.from('收入增长仅是管理层预测。')
    attachmentId = hash(raw)
    const call = (operation: string, body = {}) =>
      attachments(
        { operation, documentId: scope.documentId, attachmentId, ...body },
        new AbortController().signal,
      )
    await call('attachment_begin', {
      name: 'original.txt',
      sizeBytes: raw.length,
      sha256: attachmentId,
    })
    await call('attachment_chunk', { offset: 0, base64: raw.toString('base64') })
    await call('attachment_finish')
  }
  const researched = (await research(
    {
      operation: 'research_build',
      ...scope,
      ledgerId: 'ledger',
      expectedRevision: 0,
      draft: researchDraft(attachmentId),
    },
    new AbortController().signal,
  )) as { record: { state: string; sources: { status: string }[] } }
  expect(researched.record.state).toBe('completed')
  expect(researched.record.sources[0]!.status).toBe(shared ? 'found' : 'missing')
  await f.upload()
  expect(await f.call('finish')).toMatchObject({ state: 'ready' })
  const pc = createPresentationService({ userDataPath: f.root }),
    call = async (operation: string, body = {}) =>
      JSON.parse(
        Buffer.from(
          await pc({ operation, ...scope, ...body }, new AbortController().signal),
        ).toString(),
      )
  expect(await call('production_run', { requestId: 'req' })).not.toHaveProperty('error')
  const pageId = deck.slides[0]!.id
  expect(
    await call('production_rebuild_page', {
      parentRequestId: 'req',
      requestId: 'child',
      pageId,
      slide: { ...deck.slides[0], notes: 'revision' },
    }),
  ).not.toHaveProperty('error')
  expect(await call('production_run', { requestId: 'child' })).not.toHaveProperty('error')
  const compiled = await compilePresentationDeck({ ...deck, slides: [deck.slides[0]] }),
    raw = Buffer.from(compiled.bytes),
    backupId = 'original-page',
    hostSlideId = 'native-9001',
    slideIds = ['native-9001', 'native-9002']
  expect(compiled.sourceSlideIds).toHaveLength(1)
  expect(compiled.sourceSlideIds![0]).not.toBe(hostSlideId)
  expect(
    await call('page_backup_begin', {
      backupId,
      requestId: 'child',
      pageId,
      hostSlideId,
      slideIds,
      sha256: hash(raw),
      sizeBytes: raw.length,
    }),
  ).toMatchObject({ status: 'uploading', hostSlideId, slideIds })
  for (let offset = 0; offset < raw.length; offset += 128 * 1024)
    expect(
      await call('page_backup_chunk', {
        backupId,
        offset,
        base64: raw.subarray(offset, offset + 128 * 1024).toString('base64'),
      }),
    ).not.toHaveProperty('error')
  expect(await call('page_backup_finish', { backupId })).toMatchObject({
    status: 'ready',
    sha256: hash(raw),
  })
  expect(await call('page_backup_read', { backupId, offset: 0, length: 128 * 1024 })).toMatchObject(
    { sha256: hash(raw), base64: raw.subarray(0, 128 * 1024).toString('base64') },
  )
  const inventory = await inspectPresentationProjectInventory({ userDataPath: f.root, ...scope })
  for (const kind of kinds) {
    const resource = inventory.resources.find((r) => r.kind === kind)!
    expect(resource).toMatchObject({
      ownership: 'project_exclusive',
      disposition: 'candidate',
      fileCount: expect.any(Number),
      bytes: expect.any(Number),
    })
    expect(resource.fileCount).toBeGreaterThan(0)
    expect(resource.bytes).toBeGreaterThan(0)
  }
  const singleNames = ['comments', 'preferences', 'manual-observations'],
    ownSingles = singleNames.map((n) =>
      join(
        f.root,
        'presentation-' + n,
        hash(JSON.stringify([scope.documentId, scope.projectId])) + '.json',
      ),
    )
  for (const p of ownSingles) {
    expect(lstatSync(p).isFile()).toBe(true)
    expect(readFileSync(p).length).toBeGreaterThan(0)
  }
  const otherPaths = [
      join(f.root, 'projects', 'presentations', hash(other.projectId)),
      ...singleNames.map((n) =>
        join(
          f.root,
          'presentation-' + n,
          hash(JSON.stringify([other.documentId, other.projectId])) + '.json',
        ),
      ),
    ],
    beforeOther = new Map(otherPaths.flatMap((p) => [...tree(p)]))
  const sharedPath = join(f.root, 'presentation-attachments', hash(scope.documentId)),
    beforeShared = tree(sharedPath),
    life = new PresentationLifecycleStore(f.root)
  let tail = Promise.resolve()
  const acquireProjectLock = async () => {
    const previous = tail
    let release!: () => void
    tail = new Promise<void>((r) => {
      release = r
    })
    await previous
    return release
  }
  const make = () =>
    createPresentationProjectDeletionService({ userDataPath: f.root, acquireProjectLock })
  return {
    ...f,
    scope,
    other,
    life,
    inventory,
    ownSingles,
    otherPaths,
    beforeOther,
    sharedPath,
    beforeShared,
    make,
  }
}
it('deletes all seven valid production namespaces, with single-file preferences and project content last', async () => {
  const f = await fixture(),
    service = f.make(),
    preview = await service.preview(f.scope)
  let projectMoved = false
  const projectPath = join(f.root, 'projects', 'presentations', hash(f.scope.projectId))
  const exclusivePaths = [
    ...f.ownSingles,
    join(f.root, 'presentation-research', hash(f.scope.documentId), hash(f.scope.projectId)),
    join(
      f.root,
      'presentation-delivery-bundles',
      hash(f.scope.documentId),
      hash(f.scope.projectId),
    ),
    join(f.root, 'presentation-page-backups', hash(f.scope.projectId)),
  ]
  fault.projectMove = (source) => {
    if (source === projectPath) {
      projectMoved = true
      for (const path of exclusivePaths) expect(existsSync(path)).toBe(false)
    }
  }
  const result = await service.confirm({
    scope: f.scope,
    expectedRevision: preview.expectedRevision,
    confirmationToken: preview.confirmationToken,
    deletionId: 'delete-seven',
  })
  expect(result).toMatchObject({
    state: 'deleted',
    projectContentRetained: false,
    counts: { removed: 7, retained: 0 },
  })
  expect(projectMoved).toBe(true)
  const record = f.life.read(f.scope)!,
    receipts = record.deletion!.resources
  expect(receipts.map((r) => r.kind).sort()).toEqual([...kinds].sort())
  expect(receipts.every((r) => r.status === 'removed')).toBe(true)
  for (const path of f.ownSingles) expect(existsSync(path)).toBe(false)
  expect(f.store.projectScope(f.scope.projectId, f.scope.documentId)).toBeUndefined()
  expect(
    record.audit.filter((e) => e.action === 'resource_result').map((e) => e.counts.removed),
  ).toEqual([1, 2, 3, 4, 5, 6, 7])
  expect(record.audit.at(-1)?.action).toBe('deletion_finished')
  expect(new Map(f.otherPaths.flatMap((p) => [...tree(p)]))).toEqual(f.beforeOther)
  expect(
    (await inspectPresentationProjectInventory({ userDataPath: f.root, ...f.scope })).resources.map(
      (r) => r.kind,
    ),
  ).toEqual(['lifecycle_control'])
})
it('retains actual document-shared source bytes and same-document other project while reporting partial', async () => {
  const f = await fixture(true),
    service = f.make(),
    preview = await service.preview(f.scope)
  expect(preview.resources.find((r) => r.kind === 'attachments')).toMatchObject({
    ownership: 'document_shared',
    disposition: 'retained',
  })
  const result = await service.confirm({
    scope: f.scope,
    expectedRevision: preview.expectedRevision,
    confirmationToken: preview.confirmationToken,
    deletionId: 'shared',
  })
  expect(result).toMatchObject({ state: 'partial', projectContentRetained: true })
  expect(result.retained.map((r) => r.kind)).toContain('attachments')
  expect(tree(f.sharedPath)).toEqual(f.beforeShared)
  expect(new Map(f.otherPaths.flatMap((p) => [...tree(p)]))).toEqual(f.beforeOther)
  const records = f.life.read(f.scope)!.deletion!.resources
  for (const kind of kinds.filter((k) => k !== 'project'))
    expect(records.find((r) => r.kind === kind)?.status).toBe('removed')
  expect(records.find((r) => r.kind === 'attachments')).toMatchObject({
    status: 'retained',
    code: 'shared_resource',
  })
  expect(records.find((r) => r.kind === 'project')?.status).toBe('pending')
})
it('recovers a failed single-file namespace unlink after reopening without losing other project bytes', async () => {
  const f = await fixture(),
    service = f.make(),
    preview = await service.preview(f.scope)
  fault.singleFile = true
  const partial = await service.confirm({
    scope: f.scope,
    expectedRevision: preview.expectedRevision,
    confirmationToken: preview.confirmationToken,
    deletionId: 'retry-single',
  })
  expect(partial.state).toBe('partial')
  expect(partial.counts.failed).toBe(1)
  expect(partial.projectContentRetained).toBe(true)
  expect(f.life.read(f.scope)!.deletion!.resources.find((r) => r.status === 'failed')?.kind).toBe(
    'preferences',
  )
  const restored = await f
    .make()
    .resume({ scope: f.scope, expectedRevision: partial.revision, deletionId: 'retry-single' })
  expect(restored).toMatchObject({ state: 'deleted', counts: { removed: 7, retained: 0 } })
  expect(new Map(f.otherPaths.flatMap((p) => [...tree(p)]))).toEqual(f.beforeOther)
})
