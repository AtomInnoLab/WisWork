import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync, readdirSync, unlinkSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { createPresentationService } from '../src/main/presentation-service'
import { PresentationStore, PresentationLifecycleStore } from '@wiswork/project-store'
import { researchFixture } from '../../../packages/pptx-engine/tests/fixtures/presentation-research'
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

it.each(['deleting', 'deleted'] as const)(
  'private research read still refuses %s control after production binding disappears',
  async (state) => {
    const f = await fixture(),
      scope = { documentId: 'doc', projectId: f.plan.projectId },
      life = new PresentationLifecycleStore(f.root)
    f.store.savePlan(scope.projectId, scope.documentId, 0, f.plan)
    const r = life.read(scope)!
    const deleting = life.beginDeletion(scope, r.revision, {
      deletionId: 'd',
      reason: 'user',
      resources: [{ resourceId: 'r', kind: 'project', ownership: 'project_exclusive' }],
    })
    if (state === 'deleted') {
      const result = life.recordDeletionResult(scope, deleting.revision, {
        deletionId: 'd',
        resourceId: 'r',
        status: 'removed',
      })
      life.finishDeletion(scope, result.revision, 'd')
    }
    const projects = join(f.root, 'projects', 'presentations')
    unlinkSync(join(projects, readdirSync(projects)[0]!, 'project.json'))
    expect(await f.call('research_read', { ledgerId: 'A' })).toEqual({ error: 'project_' + state })
  },
)
it('replays a sole durable deletion receipt without inventing new research ownership', async () => {
  const f = await fixture()
  const first = await f.remove()
  expect(await f.remove()).toEqual(first)
  expect(
    await f.call('research_abandon', {
      ledgerId: 'A',
      expectedRevision: 1,
      expectedDraftDigest: f.built.record.draftDigest,
    }),
  ).toEqual({ error: 'record_deleted' })
})
