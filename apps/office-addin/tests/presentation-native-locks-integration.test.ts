import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { expect, it } from 'vitest'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { readPresentationNativeLocks } from '../src/skills/powerpoint/presentation-native-locks.js'
import type { PresentationImportRecord } from '../src/skills/powerpoint/presentation-delivery.js'

it('restores current lock coverage from a real compiled PC source after service restart and explicit unlock', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-native-locks-'))
  try {
    let service = createPresentationService({ userDataPath })
    const plan = benchmarkPlan()
    const request = async (body: unknown, signal?: AbortSignal) =>
      new Response(
        Buffer.from(await service(body, signal ?? new AbortController().signal)).toString('utf8'),
      )
    const call = async (operation: string, fields: Record<string, unknown> = {}) =>
      (await request({ operation, documentId: 'doc', projectId: plan.projectId, ...fields })).json()
    expect(await call('save_plan', { expectedRevision: 0, plan })).toHaveProperty('revision', 1)
    const compiled = await call('compile', {
      requestId: 'whole',
      planRevision: 1,
      deck: benchmarkPlannedDeck(),
    })
    expect(compiled).not.toHaveProperty('error')
    expect(
      await call('set_plan_page_lock', {
        expectedRevision: 1,
        pageId: plan.slides[0]!.id,
        locked: true,
      }),
    ).toHaveProperty('revision', 2)
    const record: PresentationImportRecord = {
      state: 'complete',
      documentId: 'doc',
      slideIds: plan.slides.map((_, index) => `host-${index}`),
      checkpoint: {
        version: 1,
        artifactDigest: createHash('sha256').update(compiled.pptxBase64).digest('hex'),
        sourceSlideIds: compiled.pages.map((page: { sourceSlideId: string }) => page.sourceSlideId),
        baselineSlideIds: [],
        completed: compiled.pages.map((page: { sourceSlideId: string }, index: number) => ({
          sourceSlideId: page.sourceSlideId,
          slideId: `host-${index}`,
        })),
      },
    }
    const options = {
      request,
      available: () => true,
      documentId: async () => 'doc',
      lastProject: () => plan.projectId,
      listReceipts: () => [{ key: `${plan.projectId}/whole`, record: structuredClone(record) }],
      hostSlideIds: async () => [...record.slideIds!],
    }
    const proposal = {
      id: 'p',
      operation: 'edit_existing_presentation_text',
      toolName: 'edit_existing_presentation_text',
      title: 'Edit',
      preview: {},
      impact: { host: 'powerpoint', targets: ['host-0'], count: 1 },
      fingerprint: 'fp',
    }
    service = createPresentationService({ userDataPath })
    const before = await readPresentationNativeLocks(
      options,
      proposal,
      new AbortController().signal,
    )
    expect(before?.pages).toEqual([
      {
        projectId: plan.projectId,
        pageId: plan.slides[0]!.id,
        title: plan.slides[0]!.title,
        slideIds: ['host-0'],
      },
    ])
    expect(
      await call('set_plan_page_lock', {
        expectedRevision: 2,
        pageId: plan.slides[0]!.id,
        locked: false,
      }),
    ).toHaveProperty('revision', 3)
    const after = await readPresentationNativeLocks(options, proposal, new AbortController().signal)
    expect(after?.pages).toEqual([])
    expect(after?.token).not.toBe(before?.token)
  } finally {
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
