import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationService } from '../src/main/presentation-service.js'
import { PresentationStore } from '@wiswork/project-store'
import { createPresentationProjectController } from '../../office-addin/src/skills/powerpoint/presentation-project.js'
it('keeps ordinary saves unapproved and persists only exact explicit decisions through PC restart', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'pc-plan-acceptance-'))
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
        ).toString(),
      )
    await call('save_plan', { expectedRevision: 0, plan })
    expect((await call('status')).planAcceptance.records).toEqual([])
    const planDigest = new PresentationStore(userDataPath).plan(plan.projectId, 'doc')!.inputDigest
    const accepted = await call('accept_plan', {
      decisionId: 'decision',
      expectedRevision: 1,
      planDigest,
    })
    expect(accepted).toMatchObject({
      projectId: plan.projectId,
      documentId: 'doc',
      acceptance: { decisionId: 'decision', planRevision: 1, planDigest },
    })
    service = createPresentationService({ userDataPath })
    expect((await call('status')).planAcceptance.records).toEqual([accepted.acceptance])
    await call('save_plan', { expectedRevision: 1, plan: { ...plan, title: 'new plan' } })
    expect(
      await call('accept_plan', { decisionId: 'stale', expectedRevision: 1, planDigest }),
    ).toEqual({ error: 'revision_conflict' })
    expect(
      await call('accept_plan', { decisionId: 'decision', expectedRevision: 1, planDigest }),
    ).toEqual(accepted)
    expect(
      await call('accept_plan', { decisionId: 'bad', expectedRevision: 2, planDigest: 'secret' }),
    ).toEqual({ error: 'invalid_request' })
  } finally {
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
it('accepts from the actual plugin controller, restores the exact decision and keeps a changed plan unapproved', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'pc-plan-acceptance-workbench-'))
  try {
    let service = createPresentationService({ userDataPath })
    const plan = benchmarkPlan()
    await service(
      {
        operation: 'save_plan',
        projectId: plan.projectId,
        documentId: 'doc',
        expectedRevision: 0,
        plan,
      },
      new AbortController().signal,
    )
    const open = () =>
      createPresentationProjectController({
        available: () => true,
        documentId: async () => 'doc',
        lastProject: () => plan.projectId,
        executeTool: async () => ({ output: '', summary: '' }),
        request: async (body, signal) =>
          new Response(Buffer.from(await service(body, signal ?? new AbortController().signal))),
      })
    const controller = open()
    await controller.refresh()
    expect(controller.snapshot().project?.planAcceptanceCurrent).toBeUndefined()
    expect(typeof controller.acceptPlan).toBe('function')
    await controller.acceptPlan!(1)
    expect(controller.snapshot().error).toBeUndefined()
    const decision = controller.snapshot().project?.planAcceptanceCurrent
    expect(decision).toMatchObject({ planRevision: 1 })
    controller.clear()
    service = createPresentationService({ userDataPath })
    const reopened = open()
    await reopened.refresh()
    expect(reopened.snapshot().project?.planAcceptanceCurrent).toEqual(decision)
    await service(
      {
        operation: 'save_plan',
        projectId: plan.projectId,
        documentId: 'doc',
        expectedRevision: 1,
        plan: { ...plan, title: 'revised' },
      },
      new AbortController().signal,
    )
    await reopened.refresh()
    expect(reopened.snapshot().project?.planAcceptanceCurrent).toBeUndefined()
    expect(reopened.snapshot().project?.planAcceptance?.records).toEqual([decision])
    await reopened.acceptPlan!(1)
    expect(reopened.snapshot().error).toBeTruthy()
    expect(
      new PresentationStore(userDataPath).planAcceptances(plan.projectId, 'doc').records,
    ).toEqual([decision])
    reopened.clear()
  } finally {
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
it('recovers a decision saved by PC when its response is lost without replaying acceptance', async () => {
  const userDataPath = mkdtempSync(join(tmpdir(), 'pc-plan-acceptance-lost-'))
  try {
    const service = createPresentationService({ userDataPath }),
      plan = benchmarkPlan()
    await service(
      {
        operation: 'save_plan',
        projectId: plan.projectId,
        documentId: 'doc',
        expectedRevision: 0,
        plan,
      },
      new AbortController().signal,
    )
    let accepts = 0
    const controller = createPresentationProjectController({
      available: () => true,
      documentId: async () => 'doc',
      lastProject: () => plan.projectId,
      executeTool: async () => ({ output: '', summary: '' }),
      request: async (body, signal) => {
        const result = await service(body, signal ?? new AbortController().signal)
        if ((body as { operation: string }).operation === 'accept_plan') {
          accepts++
          throw new Error('lost response')
        }
        return new Response(Buffer.from(result))
      },
    })
    await controller.refresh()
    await controller.acceptPlan!(1)
    expect(controller.snapshot().error).toContain('不自动重复提交')
    const records = new PresentationStore(userDataPath).planAcceptances(
      plan.projectId,
      'doc',
    ).records
    expect(records).toHaveLength(1)
    await controller.refresh()
    expect(controller.snapshot().project?.planAcceptanceCurrent).toEqual(records[0])
    expect(accepts).toBe(1)
    controller.clear()
  } finally {
    rmSync(userDataPath, { recursive: true, force: true })
  }
})
