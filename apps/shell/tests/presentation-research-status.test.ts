import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { createHash } from 'node:crypto'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationStore } from '@wiswork/project-store'
import { PresentationResearchStore } from '@wiswork/project-store/presentation-research-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationService } from '../src/main/presentation-service'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
const hash = (value: string) => createHash('sha256').update(value).digest('hex')
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'research-status-'))
  roots.push(root)
  const plan = benchmarkPlan()
  new PresentationStore(root).savePlan(plan.projectId, 'doc', 0, plan)
  const research = new PresentationResearchStore(root)
  let service = createPresentationService({ userDataPath: root })
  return {
    root,
    plan,
    research,
    restart: () => {
      service = createPresentationService({ userDataPath: root })
    },
    status: async (documentId = 'doc') =>
      JSON.parse(
        Buffer.from(
          await service(
            { operation: 'status', documentId, projectId: plan.projectId },
            new AbortController().signal,
          ),
        ).toString(),
      ),
  }
}
const draft = { scope: 'synthetic original draft', sources: [], facts: [] }
it('returns the actual empty summary without inventing records or changing the project', async () => {
  const f = fixture(),
    status = await f.status()
  expect(status.researchSummary).toEqual(await f.research.summary('doc', f.plan.projectId))
  expect(status).not.toHaveProperty('researchHistoryUnavailable')
  expect(status.plan.value).toEqual(f.plan)
  expect(status.status).toBe('planned')
})
it('tracks actual begin, finish, abandon, delete and restart while keeping V2 sequence', async () => {
  const f = fixture(),
    project = f.plan.projectId
  const a = (await f.research.begin('doc', project, 0, 'A', draft)).record
  expect((await f.status()).researchSummary.records[0].state).toBe('running')
  await f.research.finish('doc', project, 'A', { state: 'completed', sources: [] })
  const b = (await f.research.begin('doc', project, 2, 'B', draft)).record
  await f.research.abandon('doc', project, 3, 'B', b.draftDigest)
  await f.research.deleteRecord('doc', project, 4, 'delete-A', 'A', a.draftDigest)
  f.restart()
  const status = await f.status()
  expect(status.researchSummary).toEqual(await f.research.summary('doc', project))
  expect(status.researchSummary).toMatchObject({
    version: 2,
    revision: 5,
    totalRecords: 1,
    lastSequence: 2,
    records: [{ id: 'B', state: 'failed', error: 'aborted' }],
  })
  expect(JSON.stringify(status.researchSummary)).not.toContain(draft.scope)
})
it('bounds the genuine summary to the latest 32 and retains the total', async () => {
  const f = fixture()
  for (let i = 0; i < 33; i++) await f.research.begin('doc', f.plan.projectId, i, `R${i}`, draft)
  const summary = (await f.status()).researchSummary
  expect(summary).toEqual(await f.research.summary('doc', f.plan.projectId))
  expect(summary.totalRecords).toBe(33)
  expect(summary.records).toHaveLength(32)
  expect(summary.records.at(-1).sequence).toBe(33)
})
it('marks corrupt research unavailable while preserving the real project and rejects cross-document status', async () => {
  const f = fixture()
  await f.research.begin('doc', f.plan.projectId, 0, 'A', draft)
  writeFileSync(
    join(f.root, 'presentation-research', hash('doc'), hash(f.plan.projectId), 'state.json'),
    '{broken',
  )
  f.restart()
  const status = await f.status()
  expect(status.researchHistoryUnavailable).toBe(true)
  expect(status).not.toHaveProperty('researchSummary')
  expect(status.plan.value).toEqual(f.plan)
  expect(status.status).toBe('planned')
  expect(await f.status('other-doc')).toEqual({ error: 'document_mismatch' })
})
