import { afterEach, expect, it } from 'vitest'
import JSZip from 'jszip'
import { createHash } from 'node:crypto'
import { presentationDeliveryBundleFiles } from '@wiswork/project-store/presentation-delivery-bundle'
import { buildPresentationDeliveryReport } from '@wiswork/pptx-engine/presentation-delivery-report'
import { presentationPlanClaims } from '@wiswork/pptx-engine/presentation-plan'
import { mkdtempSync, rmSync, readdirSync, writeFileSync, unlinkSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { createPresentationService } from '../src/main/presentation-service'
import { PresentationStore } from '@wiswork/project-store'
import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import { researchFixture } from '../../../packages/pptx-engine/tests/fixtures/presentation-research'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
async function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'research-cleanup-pc-'))
  roots.push(root)
  const service = createPresentationService({ userDataPath: root })
  const { plan, record } = researchFixture()
  const call = async (operation: string, fields: Record<string, unknown> = {}) =>
    JSON.parse(
      Buffer.from(
        await service(
          { operation, documentId: 'doc', projectId: plan.projectId, ...fields },
          new AbortController().signal,
        ),
      ).toString(),
    )
  const built = await call('research_build', {
    ledgerId: 'A',
    expectedRevision: 0,
    draft: record.draft,
  })
  plan.research!.ledgerId = 'A'
  plan.research!.draftDigest = built.record.draftDigest
  const remove = () =>
    call('research_delete', {
      ledgerId: 'A',
      deleteId: 'delete-A',
      expectedRevision: built.history.revision,
      expectedDraftDigest: built.record.draftDigest,
    })
  return { root, plan, call, remove, built, store: new PresentationStore(root) }
}
it('deletes an unregistered research project with negotiated V2 history and idempotent ACK', async () => {
  const f = await fixture()
  expect(
    await f.call('research_capabilities', { includeCleanup: true, projectId: undefined }),
  ).toHaveProperty('error')
  const deleted = await f.remove()
  expect(deleted).toMatchObject({ ledgerId: 'A', deleteId: 'delete-A', version: 1 })
  expect(await f.remove()).toEqual(deleted)
  expect(await f.call('research_delete_status', { deleteId: 'delete-A' })).toEqual(deleted)
  expect(await f.call('research_list')).toEqual({ error: 'upgrade_required' })
  expect(await f.call('research_list', { historyVersion: 2 })).toMatchObject({
    version: 2,
    totalRecords: 0,
    lastSequence: 1,
  })
  expect(await f.call('research_read', { ledgerId: 'A' })).toEqual({ error: 'record_deleted' })
  expect(await f.call('save_plan', { expectedRevision: 0, plan: f.plan })).toEqual({
    error: 'research_unavailable',
  })
})
it('protects current and recoverable historical plans and rejects corrupt history', async () => {
  const f = await fixture()
  expect(await f.call('save_plan', { expectedRevision: 0, plan: f.plan })).not.toHaveProperty(
    'error',
  )
  expect(await f.remove()).toEqual({ error: 'record_protected' })
  const changed = structuredClone(f.plan)
  delete changed.research
  expect(await f.call('save_plan', { expectedRevision: 1, plan: changed })).not.toHaveProperty(
    'error',
  )
  expect(await f.remove()).toEqual({ error: 'record_protected' })
  const project = join(
    f.root,
    'projects',
    'presentations',
    readdirSync(join(f.root, 'projects', 'presentations'))[0]!,
  )
  writeFileSync(join(project, 'plan-revision-1.json'), 'corrupt')
  expect(await f.remove()).toEqual({ error: 'invalid_state' })
  expect(await f.call('research_read', { ledgerId: 'A' })).toHaveProperty('id', 'A')
})
it('protects all frozen compile plans beyond the public 20-item window', async () => {
  const f = await fixture(),
    deck = benchmarkDeck()
  f.store.begin(f.plan.projectId, 'doc', 'old-bound', deck, { revision: 1, plan: f.plan })
  const unbound = structuredClone(f.plan)
  delete unbound.research
  for (let index = 0; index < 21; index++)
    f.store.begin(f.plan.projectId, 'doc', 'new-' + index, deck, { revision: 1, plan: unbound })
  expect(f.store.history(f.plan.projectId, 'doc')).toHaveLength(20)
  expect(f.store.history(f.plan.projectId, 'doc', true)).toHaveLength(22)
  expect(await f.remove()).toEqual({ error: 'record_protected' })
})
it.each(['plan.json', 'project.json'])(
  'refuses cleanup when %s disappears but recoverable snapshots remain',
  async (missing) => {
    const f = await fixture()
    expect(await f.call('save_plan', { expectedRevision: 0, plan: f.plan })).not.toHaveProperty(
      'error',
    )
    const changed = structuredClone(f.plan)
    delete changed.research
    expect(await f.call('save_plan', { expectedRevision: 1, plan: changed })).not.toHaveProperty(
      'error',
    )
    const project = join(
      f.root,
      'projects',
      'presentations',
      readdirSync(join(f.root, 'projects', 'presentations'))[0]!,
    )
    unlinkSync(join(project, missing))
    expect(await f.remove()).toEqual({ error: 'invalid_state' })
    expect(await f.call('research_read', { ledgerId: 'A' })).toHaveProperty('id', 'A')
    expect(existsSync(join(project, missing))).toBe(false)
  },
)
it('protects frozen production plans and refuses running records and cross-document deletion', async () => {
  const f = await fixture()
  f.store.beginProduction(f.plan.projectId, 'doc', 'production', benchmarkDeck(), {
    revision: 1,
    plan: f.plan,
  })
  expect(await f.remove()).toEqual({ error: 'record_protected' })
  await new PresentationResearchStore(f.root).begin(
    'doc',
    f.plan.projectId,
    f.built.history.revision,
    'running',
    f.built.record.draft,
  )
  expect(
    await f.call('research_delete', {
      ledgerId: 'running',
      deleteId: 'delete-running',
      expectedRevision: 3,
      expectedDraftDigest: f.built.record.draftDigest,
    }),
  ).toEqual({ error: 'record_running' })
  expect(
    await f.call('research_delete', {
      documentId: 'other',
      ledgerId: 'A',
      deleteId: 'cross',
      expectedRevision: 2,
      expectedDraftDigest: f.built.record.draftDigest,
    }),
  ).toHaveProperty('error')
})
it('serializes binding and cleanup in either order on the same project lock', async () => {
  const boundFirst = await fixture()
  const [saved, protectedResult] = await Promise.all([
    boundFirst.call('save_plan', { expectedRevision: 0, plan: boundFirst.plan }),
    boundFirst.remove(),
  ])
  expect(saved).not.toHaveProperty('error')
  expect(protectedResult).toEqual({ error: 'record_protected' })
  const deleteFirst = await fixture()
  const [deleted, refused] = await Promise.all([
    deleteFirst.remove(),
    deleteFirst.call('save_plan', { expectedRevision: 0, plan: deleteFirst.plan }),
  ])
  expect(deleted).toHaveProperty('deleteId', 'delete-A')
  expect(refused).toEqual({ error: 'research_unavailable' })
})
it('blocks unfinished research ZIP uploads but preserves ready ZIP independently after cleanup', async () => {
  const f = await fixture(),
    plan = structuredClone(f.plan),
    deck = benchmarkDeck()
  delete plan.research
  deck.claims = presentationPlanClaims(plan)
  const production = f.store.beginProduction(plan.projectId, 'doc', 'bundle', deck, {
    revision: 1,
    plan,
  })
  const report = await buildPresentationDeliveryReport({
    plan,
    deck,
    metadata: {
      projectId: plan.projectId,
      documentId: 'doc',
      requestId: 'bundle',
      planRevision: 1,
      inputDigest: production.inputDigest,
      planDigest: production.planDigest,
    },
    pageStates: production.pages.map((page) => ({ pageId: page.pageId, state: page.state })),
    reviews: [],
    issueLedger: f.store.issueActions(plan.projectId, 'doc', 'bundle'),
  })
  const hash = (bytes: string | Uint8Array) => createHash('sha256').update(bytes).digest('hex')
  const files = new Map(
    presentationDeliveryBundleFiles.map((name) => [
      name,
      Buffer.from(
        name === 'presentation.pptx'
          ? 'PK\u0003\u0004native'
          : name === 'evidence.json'
            ? JSON.stringify(report)
            : name === 'claims.json'
              ? JSON.stringify(plan.claims)
              : name === 'sources.json'
                ? JSON.stringify(plan.sources)
                : 'Historical not verified',
      ),
    ]),
  )
  files.set('research.json', Buffer.from(JSON.stringify(f.built.record)))
  files.set('research.md', Buffer.from('Historical original research'))
  const zip = new JSZip()
  for (const [name, bytes] of files) zip.file(name, bytes)
  const manifest = {
    version: 1,
    scope: 'current_office_document',
    documentId: 'doc',
    projectId: plan.projectId,
    requestId: 'bundle',
    planRevision: 1,
    inputDigest: production.inputDigest,
    planDigest: production.planDigest,
    createdAt: new Date().toISOString(),
    files: [...files].map(([name, data]) => ({ name, sizeBytes: data.length, sha256: hash(data) })),
    checks: {
      completion: 'not_verified',
      sourceAuthority: 'not_verified',
      timeliness: 'not_verified',
      roundTrip: 'not_run',
      hostQa: 'not_checked',
      pdf: 'not_requested',
    },
  }
  zip.file('manifest.json', JSON.stringify(manifest))
  const bytes = await zip.generateAsync({ type: 'nodebuffer' }),
    bundleId = hash(bytes)
  expect(
    await f.call('delivery_bundle_begin', {
      requestId: 'bundle',
      bundleId,
      sha256: bundleId,
      sizeBytes: bytes.length,
      manifest,
    }),
  ).toHaveProperty('state', 'uploading')
  expect(await f.remove()).toEqual({ error: 'busy' })
  expect(
    await f.call('delivery_bundle_chunk', {
      requestId: 'bundle',
      bundleId,
      offset: 0,
      base64: bytes.toString('base64'),
    }),
  ).not.toHaveProperty('error')
  expect(await f.call('delivery_bundle_finish', { requestId: 'bundle', bundleId })).toHaveProperty(
    'state',
    'ready',
  )
  expect(await f.remove()).toHaveProperty('deleteId', 'delete-A')
  expect(
    await f.call('delivery_bundle_metadata', { requestId: 'bundle', bundleId }),
  ).toHaveProperty('state', 'ready')
  expect(await f.call('delivery_bundle_finish', { requestId: 'bundle', bundleId })).toHaveProperty(
    'state',
    'ready',
  )
  const read = await f.call('delivery_bundle_read', {
    requestId: 'bundle',
    bundleId,
    offset: 0,
    length: bytes.length,
  })
  expect(Buffer.from(read.base64, 'base64')).toEqual(bytes)
})
