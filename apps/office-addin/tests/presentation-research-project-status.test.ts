import { expect, it, vi } from 'vitest'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project'
import { researchSummary } from './presentation-research-fixture'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
const project = {
  projectId: 'research',
  title: 'Research project',
  status: 'pending',
  latestRequestId: 'request',
  slideCount: 1,
  slides: [{ id: 'slide', title: 'Page' }],
  history: [{ requestId: 'request', sequence: 1, status: 'pending', slideCount: 1 }],
}
function fixture(summary: unknown) {
  let status: Record<string, unknown> = {
    ...project,
    ...(summary !== undefined ? { researchSummary: summary } : {}),
  }
  let document = 'doc'
  const request = vi.fn(async () => Response.json(status))
  const create = () =>
    createPresentationProjectController({
      available: () => true,
      request,
      documentId: async () => document,
      lastProject: () => project.projectId,
      executeTool: vi.fn(async () => ({ output: '{}', mutated: false, summary: 'read' })),
    })
  return {
    request,
    create,
    set: (value: Record<string, unknown>) => {
      status = value
    },
    changeDocument: () => {
      document = 'other'
    },
  }
}
it('receives actual V2 summary on refresh and reopen without additional requests', async () => {
  const summary = {
    ...researchSummary(),
    version: 2,
    lastSequence: 2,
    revision: 4,
    totalRecords: 1,
    records: researchSummary().records.map((record) => ({
      ...record,
      sequence: 2,
      state: 'running',
      finishedAt: undefined,
    })),
  }
  // Omitted optional terminal fields are required for a running record.
  delete summary.records[0]!.finishedAt
  const f = fixture(summary),
    controller = f.create()
  await controller.refresh()
  expect(controller.snapshot().project?.researchSummary).toEqual(summary)
  const reopened = f.create()
  await reopened.refresh()
  expect(reopened.snapshot().project?.researchSummary?.projectId).toBe('research')
  expect(f.request).toHaveBeenCalledTimes(2)
})
it('fails only research projection for malformed, wrong-project or wrong-document summaries', async () => {
  const f = fixture(researchSummary()),
    controller = f.create()
  await controller.refresh()
  for (const summary of [
    { ...researchSummary(), projectId: 'wrong' },
    { ...researchSummary(), documentId: 'other' },
    { ...researchSummary(), version: 2, lastSequence: 0 },
    { ...researchSummary(), extra: true },
  ]) {
    f.set({ ...project, researchSummary: summary })
    await controller.refresh()
    expect(controller.snapshot().project).toMatchObject({
      ...project,
      researchHistoryUnavailable: true,
    })
    expect(controller.snapshot().project).not.toHaveProperty('researchSummary')
  }
})
it('keeps existing project and saved plan when research projection is corrupt', async () => {
  const f = fixture(researchSummary()),
    plan = benchmarkPlan()
  plan.projectId = 'research'
  const saved = { revision: 1, value: plan }
  f.set({ ...project, plan: saved, researchSummary: { ...researchSummary(), projectId: 'wrong' } })
  const controller = f.create()
  await controller.refresh()
  expect(controller.snapshot().project?.plan).toEqual(saved)
  expect(controller.snapshot().project?.researchHistoryUnavailable).toBe(true)
  expect(controller.snapshot().error).toBeUndefined()
})
it('honors PC unavailable over legal summary and retains exact legacy absence on refresh', async () => {
  const f = fixture(researchSummary()),
    controller = f.create()
  await controller.refresh()
  f.set({ ...project, researchSummary: researchSummary(), researchHistoryUnavailable: true })
  await controller.refresh()
  expect(controller.snapshot().project).toMatchObject({ researchHistoryUnavailable: true })
  expect(controller.snapshot().project).not.toHaveProperty('researchSummary')
  f.set({ ...project })
  await controller.refresh()
  expect(controller.snapshot().project).toEqual(project)
})
it('refuses publication when current document changes while awaiting status', async () => {
  const f = fixture(researchSummary()),
    controller = f.create()
  await controller.refresh()
  f.request.mockImplementationOnce(async () => {
    f.changeDocument()
    return Response.json({ ...project, researchSummary: researchSummary() })
  })
  await controller.refresh()
  expect(controller.snapshot().project).toBeUndefined()
  expect(controller.snapshot().error).toBeTruthy()
})
