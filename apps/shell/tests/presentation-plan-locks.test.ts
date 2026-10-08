import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationService } from '../src/main/presentation-service.js'

it('protects saved locked plans and allows an exact, explicit page unlock after service reopen', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-lock-service-'))
  try {
    let service = createPresentationService({ userDataPath })
    const plan = benchmarkPlan()
    const call = async (operation: string, fields: Record<string, unknown> = {}) =>
      JSON.parse(
        Buffer.from(
          await service(
            { operation, projectId: plan.projectId, documentId: 'doc', ...fields },
            new AbortController().signal,
          ),
        ).toString('utf8'),
      )
    await call('save_plan', { expectedRevision: 0, plan })
    const deck = benchmarkPlannedDeck()
    expect(
      await call('production_begin', { requestId: 'original', planRevision: 1, deck }),
    ).not.toHaveProperty('error')
    const pageId = plan.slides[1]!.id
    const locked = await call('set_plan_page_lock', { expectedRevision: 1, pageId, locked: true })
    expect(locked).toMatchObject({ revision: 2 })
    service = createPresentationService({ userDataPath })
    const changed = structuredClone(deck)
    const text = changed.slides[1]!.elements.find((element) => element.kind === 'text')!
    if (text.kind === 'text') text.text += ' changed'
    expect(
      await call('production_begin', { requestId: 'changed', planRevision: 2, deck: changed }),
    ).toEqual({ error: 'page_locked' })
    expect(
      await call('compile', { requestId: 'compile-changed', planRevision: 2, deck: changed }),
    ).toEqual({ error: 'page_locked' })
    expect(await call('save_plan', { expectedRevision: 2, plan })).toEqual({ error: 'page_locked' })
    expect(
      await call('set_plan_page_lock', { expectedRevision: 1, pageId, locked: false }),
    ).toEqual({ error: 'revision_conflict' })
    expect(
      await call('set_plan_page_lock', { expectedRevision: 2, pageId, locked: 'false' }),
    ).toEqual({ error: 'invalid_request' })
    expect(
      await call('set_plan_page_lock', { expectedRevision: 2, pageId, locked: false }),
    ).toMatchObject({ revision: 3, plan })
    expect(
      await call('save_plan', { expectedRevision: 3, plan: { ...plan, title: 'changed' } }),
    ).toMatchObject({ revision: 4 })
  } finally {
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
