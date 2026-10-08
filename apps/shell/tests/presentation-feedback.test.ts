import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createPresentationService } from '../src/main/presentation-service'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
async function setup() {
  const root = mkdtempSync(join(tmpdir(), 'feedback-service-'))
  roots.push(root)
  const service = createPresentationService({
    userDataPath: root,
    compile: compilePresentationDeck,
  })
  const plan = benchmarkPlan(),
    deck = benchmarkPlannedDeck()
  plan.slides = plan.slides.slice(0, 2)
  deck.slides = deck.slides.slice(0, 2)
  const call = async (body: Record<string, unknown>, signal = new AbortController().signal) =>
    JSON.parse(
      Buffer.from(
        await service({ documentId: 'doc', projectId: plan.projectId, ...body }, signal),
      ).toString('utf8'),
    )
  expect(await call({ operation: 'save_plan', expectedRevision: 0, plan })).toMatchObject({
    revision: 1,
  })
  expect(
    await call({ operation: 'production_begin', requestId: 'run', planRevision: 1, deck }),
  ).not.toHaveProperty('error')
  expect(await call({ operation: 'production_run', requestId: 'run' })).toMatchObject({
    status: 'compiled',
  })
  return { call, pageId: plan.slides[0]!.id }
}
it('strictly rejects forged server binding, invalid patches, stale CAS and cancelled writes', async () => {
  const { call, pageId } = await setup()
  const record = {
    operation: 'production_feedback_record',
    requestId: 'run',
    expectedRevision: 0,
    pages: [{ pageId, status: 'needs_correction', note: '原说明' }],
  }
  for (const extra of [
    { source: 'qa' },
    { inputDigest: 'a'.repeat(64) },
    { approvedAt: '2026-09-29T00:00:00.000Z' },
    { pages: [{ pageId, status: 'passed' }] },
    { pages: [{ pageId, status: 'no_correction', extra: true }] },
    { expectedRevision: -1 },
    { requestId: '../run' },
  ])
    expect(await call({ ...record, ...extra })).toEqual({ error: 'invalid_request' })
  expect(await call({ ...record, documentId: 'other' })).toEqual({ error: 'document_mismatch' })
  const abort = new AbortController()
  abort.abort()
  expect(await call(record, abort.signal)).toEqual({ error: 'aborted' })
  expect(await call({ operation: 'production_feedback_read', requestId: 'run' })).toEqual({
    feedback: null,
  })
  const saved = await call(record)
  expect(saved.feedback.revision).toBe(1)
  expect(await call(record)).toEqual(saved)
  expect(await call({ ...record, pages: [{ pageId, status: 'no_correction' }] })).toEqual({
    error: 'revision_conflict',
  })
  expect(await call({ operation: 'production_feedback_read', requestId: 'missing' })).toEqual({
    error: 'not_found',
  })
  expect(await call({ operation: 'production_feedback_read', requestId: 'run' })).toEqual(saved)
})
