import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationService } from '../src/main/presentation-service.js'

it('reads a requested historical plan exactly and preserves current-state and document validation', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-history-service-'))
  try {
    const service = createPresentationService({ userDataPath })
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
    expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
    const revised = { ...plan, title: '更新后的计划' }
    expect(await call('save_plan', { expectedRevision: 1, plan: revised })).toMatchObject({
      revision: 2,
    })
    expect(await call('get_plan', { revision: 1 })).toMatchObject({ revision: 1, plan })
    expect(await call('get_plan')).toMatchObject({ revision: 2, plan: revised })
    expect(await call('get_plan', { revision: 99 })).toEqual({ error: 'plan_revision_unavailable' })
    expect(await call('get_plan', { revision: 1, documentId: 'foreign' })).toEqual({
      error: 'document_mismatch',
    })
    for (const revision of [0, -1, '1', 1.5, Number.MAX_SAFE_INTEGER + 1])
      expect(await call('get_plan', { revision })).toEqual({ error: 'invalid_request' })
  } finally {
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
