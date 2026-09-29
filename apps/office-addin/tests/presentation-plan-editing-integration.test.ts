import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { PresentationStore } from '@wiswork/project-store'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'

it('persists workbench reorder/delete through the real PC service and recovers verified revision history', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'ppt-plan-adjustment-'))
  const plan = benchmarkPlan()
  const deck = benchmarkPlannedDeck()
  const compile = vi.fn(compilePresentationDeck)
  const create = () => {
    const service = createPresentationService({ userDataPath, compile })
    const request = vi.fn(
      async (body: unknown, signal?: AbortSignal) =>
        new Response(
          Buffer.from(await service(body, signal ?? new AbortController().signal)).toString('utf8'),
        ),
    )
    const runtime = createOfficeHostRuntime('powerpoint', {
      presentation: {
        request,
        documentId: async () => 'doc',
        available: () => true,
        lastProject: () => plan.projectId,
        rememberProject: async () => undefined,
      },
    })
    return { service, request, runtime }
  }
  const first = create()
  let second: ReturnType<typeof create> | undefined
  try {
    await first.service(
      {
        operation: 'save_plan',
        documentId: 'doc',
        projectId: plan.projectId,
        expectedRevision: 0,
        plan,
      },
      new AbortController().signal,
    )
    await first.service(
      {
        operation: 'production_begin',
        documentId: 'doc',
        projectId: plan.projectId,
        requestId: 'original',
        planRevision: 1,
        deck,
      },
      new AbortController().signal,
    )
    await first.service(
      {
        operation: 'production_run',
        documentId: 'doc',
        projectId: plan.projectId,
        requestId: 'original',
      },
      new AbortController().signal,
    )
    expect(compile).toHaveBeenCalledTimes(8)
    const controller = first.runtime.presentation!
    await controller.refresh()
    expect(controller.snapshot().project?.plan?.revisions).toHaveLength(1)
    await controller.editPlan!(1, { kind: 'move', pageId: plan.slides[1]!.id, direction: 'down' })
    expect(controller.snapshot().error).toBeUndefined()
    expect(controller.snapshot().project?.plan?.revision).toBe(2)
    await controller.editPlan!(2, { kind: 'delete', pageId: plan.slides[1]!.id })
    expect(controller.snapshot().project?.plan?.revisions).toHaveLength(3)
    expect(controller.snapshot().project?.plan?.value.sources).toEqual(plan.sources)
    expect(controller.snapshot().planNotice).toContain('已有 PowerPoint 页面保留')
    const stored = new PresentationStore(userDataPath).plan(plan.projectId, 'doc')!
    expect(stored.revisions?.map((revision) => revision.snapshot?.slideCount)).toEqual([8, 8, 7])
    expect(
      first.request.mock.calls.every(([body]) =>
        ['status', 'save_plan', 'production_job_status'].includes(
          (body as { operation: string }).operation,
        ),
      ),
    ).toBe(true)
    first.runtime.dispose()
    second = create()
    await second.runtime.presentation!.refresh()
    const reopened = second.runtime.presentation!.snapshot()
    expect(reopened.error).toBeUndefined()
    expect(reopened.project?.plan?.revision).toBe(3)
    expect(reopened.project?.plan?.value.slides.map((slide) => slide.id)).toEqual(
      plan.slides.filter((slide) => slide.id !== plan.slides[1]!.id).map((slide) => slide.id),
    )
    expect(reopened.project?.plan?.value.claims).toEqual(plan.claims)
    expect(reopened.project?.production?.planRevision).toBe(1)
    const revisedDeck = {
      ...deck,
      slides: deck.slides.filter((slide) => slide.id !== plan.slides[1]!.id),
    }
    const fresh = JSON.parse(
      Buffer.from(
        await second.service(
          {
            operation: 'production_begin',
            documentId: 'doc',
            projectId: plan.projectId,
            requestId: 'revised',
            planRevision: 3,
            deck: revisedDeck,
          },
          new AbortController().signal,
        ),
      ).toString('utf8'),
    )
    expect(fresh).toMatchObject({ status: 'compiled', compiledCount: 7, planRevision: 3 })
    await second.service(
      {
        operation: 'production_run',
        documentId: 'doc',
        projectId: plan.projectId,
        requestId: 'revised',
      },
      new AbortController().signal,
    )
    expect(compile).toHaveBeenCalledTimes(8)
  } finally {
    first.runtime.dispose()
    second?.runtime.dispose()
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
