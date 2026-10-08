import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import {
  benchmarkPlan,
  benchmarkPlannedDeck,
} from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createPresentationProjectController } from '../../office-addin/src/skills/powerpoint/presentation-project'
import { createPresentationService } from '../src/main/presentation-service'

async function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'wiswork-feedback-integration-'))
  let service = createPresentationService({ userDataPath: root })
  const plan = benchmarkPlan()
  const call = async (operation: string, extra: Record<string, unknown> = {}) =>
    JSON.parse(
      Buffer.from(
        await service(
          {
            operation,
            documentId: 'doc',
            projectId: plan.projectId,
            ...(operation === 'save_plan' ? {} : { requestId: 'frozen' }),
            ...extra,
          },
          new AbortController().signal,
        ),
      ).toString('utf8'),
    )
  expect(await call('save_plan', { expectedRevision: 0, plan })).toMatchObject({ revision: 1 })
  expect(
    await call('production_begin', { planRevision: 1, deck: benchmarkPlannedDeck() }),
  ).not.toHaveProperty('error')
  return {
    plan,
    call,
    request: async (body: unknown, signal?: AbortSignal) =>
      new Response(await service(body, signal ?? new AbortController().signal)),
    restart: () => {
      service = createPresentationService({ userDataPath: root })
    },
    dispose: () => rmSync(root, { recursive: true, force: true }),
  }
}

it('retains user-reported feedback on the frozen task across plan changes and actual PC restart without certifying quality', async () => {
  const f = await fixture()
  try {
    expect(await f.call('production_feedback_read')).toEqual({ feedback: null })
    expect(await f.call('production_run')).toMatchObject({ status: 'compiled' })
    const before = await f.call('production_delivery_report')
    const pageId = f.plan.slides[0]!.id
    const saved = await f.call('production_feedback_record', {
      expectedRevision: 0,
      pages: [{ pageId, status: 'needs_correction', note: '标题过长，需要人工缩短。' }],
    })
    expect(saved).not.toHaveProperty('error')
    expect(saved.feedback).toMatchObject({
      source: 'user_reported',
      requestId: 'frozen',
      planRevision: 1,
      revision: 1,
    })
    expect(saved.feedback.snapshots[0].pages).toEqual(
      f.plan.slides.map((page, i) =>
        i === 0
          ? { pageId: page.id, status: 'needs_correction', note: '标题过长，需要人工缩短。' }
          : { pageId: page.id, status: 'not_evaluated' },
      ),
    )
    const changed = structuredClone(f.plan)
    changed.title = '修改后的当前计划'
    expect(await f.call('save_plan', { expectedRevision: 1, plan: changed })).toMatchObject({
      revision: 2,
    })
    f.restart()
    expect(await f.call('production_feedback_read')).toEqual(saved)
    const after = await f.call('production_delivery_report')
    expect(after.checks).toEqual(before.checks)
    expect(after.issueLedger).toEqual(before.issueLedger)
    expect(after.pages).toEqual(before.pages)
    expect(after.planRevision).toBe(1)
    expect(after.plan.title).toBe(f.plan.title)
    const second = await f.call('production_feedback_record', {
      expectedRevision: 1,
      pages: [{ pageId, status: 'no_correction' }],
    })
    expect(second.feedback.snapshots[0]).toEqual(saved.feedback.snapshots[0])
    expect(second.feedback.revision).toBe(2)
    expect(second.feedback.snapshots[1].pages[0]).toEqual({ pageId, status: 'no_correction' })
    expect(await f.call('production_feedback_read', { requestId: 'missing' })).toHaveProperty(
      'error',
    )
    expect(await f.call('production_feedback_read', { documentId: 'other' })).toHaveProperty(
      'error',
    )
  } finally {
    f.dispose()
  }
})

it('serializes concurrent feedback and acknowledges identical lost-response retries without erasing unknown pages', async () => {
  const f = await fixture()
  try {
    expect(await f.call('production_run')).toMatchObject({ status: 'compiled' })
    const pageId = f.plan.slides[0]!.id
    const patch = { expectedRevision: 0, pages: [{ pageId, status: 'needs_correction' }] }
    const [first, replay] = await Promise.all([
      f.call('production_feedback_record', patch),
      f.call('production_feedback_record', patch),
    ])
    expect(first).not.toHaveProperty('error')
    expect(replay).toEqual(first)
    expect(first.feedback.revision).toBe(1)
    expect(first.feedback.snapshots).toHaveLength(1)
    const [a, b] = await Promise.all([
      f.call('production_feedback_record', {
        expectedRevision: 1,
        pages: [{ pageId, status: 'no_correction' }],
      }),
      f.call('production_feedback_record', {
        expectedRevision: 1,
        pages: [{ pageId, status: 'needs_correction', note: '不同的人工判断' }],
      }),
    ])
    expect([a, b].filter((result) => !result.error)).toHaveLength(1)
    expect([a, b].filter((result) => result.error)).toHaveLength(1)
    f.restart()
    const latest = await f.call('production_feedback_read')
    expect(latest.feedback.revision).toBe(2)
    expect(latest.feedback.snapshots[0]).toEqual(first.feedback.snapshots[0])
    expect(
      latest.feedback.snapshots[1].pages
        .slice(1)
        .every((page: { status: string }) => page.status === 'not_evaluated'),
    ).toBe(true)
    expect(
      await f.call('production_feedback_record', {
        expectedRevision: 2,
        pages: [{ pageId, status: 'verified' }],
      }),
    ).toHaveProperty('error')
    expect(await f.call('production_feedback_read')).toEqual(latest)
  } finally {
    f.dispose()
  }
})

it('collects and reopens feedback through the actual Office controller and PC service without invoking an Agent tool', async () => {
  const f = await fixture()
  try {
    expect(await f.call('production_run')).toMatchObject({ status: 'compiled' })
    let toolCalls = 0
    const create = () =>
      createPresentationProjectController({
        request: f.request,
        available: () => true,
        documentId: async () => 'doc',
        lastProject: () => f.plan.projectId,
        executeTool: async () => {
          toolCalls++
          throw new Error('feedback must not invoke an Agent tool')
        },
      })
    const controller = create()
    await controller.refresh()
    expect(controller.snapshot().project?.production?.requestId).toBe('frozen')
    await controller.readProductionFeedback!()
    expect(controller.snapshot().productionFeedback).toBeNull()
    await controller.recordProductionFeedback!([
      { pageId: f.plan.slides[0]!.id, status: 'needs_correction', note: '需要人工修改标题' },
    ])
    expect(controller.snapshot().productionFeedback).toMatchObject({
      requestId: 'frozen',
      revision: 1,
      source: 'user_reported',
    })
    const saved = controller.snapshot().productionFeedback
    controller.clear()
    f.restart()
    const reopened = create()
    await reopened.refresh()
    await reopened.readProductionFeedback!()
    expect(reopened.snapshot().productionFeedback).toEqual(saved)
    expect(toolCalls).toBe(0)
    expect(reopened.snapshot().project?.production?.status).toBe('compiled')
  } finally {
    f.dispose()
  }
})
