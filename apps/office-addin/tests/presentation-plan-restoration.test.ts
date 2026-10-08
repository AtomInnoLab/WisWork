import { createHash } from 'node:crypto'
import { afterEach, expect, it, vi } from 'vitest'
import { canonicalPresentationValue } from '@wiswork/project-store/presentation-canonical'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'

afterEach(() => vi.restoreAllMocks())
function fixture() {
  const original = benchmarkPlan()
  const plans = [original, { ...original, title: '第二版' }, structuredClone(original)]
  const revisions = plans.map((plan, index) => ({
    revision: index + 1,
    inputDigest: createHash('sha256').update(canonicalPresentationValue(plan)).digest('hex'),
    createdAt: '2026-09-29T00:00:00.000Z',
  }))
  const documentId = vi.fn(async () => 'doc')
  const project = {
    projectId: original.projectId,
    title: original.title,
    status: 'planned',
    slideCount: original.slides.length,
    slides: original.slides.map(({ id, title }) => ({ id, title })),
    history: [],
    plan: { revision: 3, value: plans[2], revisions },
  }
  const request = vi.fn(async () => new Response(JSON.stringify(project)))
  const executeTool = vi.fn(async (call: { name: string; input: Record<string, unknown> }) => ({
    output: JSON.stringify({
      projectId: original.projectId,
      revision: call.input.revision,
      plan: plans[Number(call.input.revision) - 1],
    }),
    mutated: false,
    isError: false,
    summary: '历史计划',
  }))
  const controller = createPresentationProjectController({
    request,
    executeTool,
    documentId,
    available: () => true,
    lastProject: () => original.projectId,
  })
  return { controller, executeTool, documentId, plans, original, project }
}

it('refreshes current state before saying a historical restoration needs no new revision', async () => {
  const f = fixture()
  await f.controller.refresh()
  const concurrent = { ...f.original, title: '并发更新' }
  f.project.title = concurrent.title
  f.project.plan = {
    revision: 4,
    value: concurrent,
    revisions: [
      ...f.project.plan.revisions,
      {
        revision: 4,
        createdAt: '2026-09-29T00:00:00.000Z',
        inputDigest: createHash('sha256')
          .update(canonicalPresentationValue(concurrent))
          .digest('hex'),
      },
    ],
  }
  await f.controller.editPlan!(3, { kind: 'restore', revision: 1 })
  expect(f.controller.snapshot().project?.plan?.revision).toBe(4)
  expect(f.controller.snapshot().project?.plan?.value.title).toBe('并发更新')
  expect(f.controller.snapshot().planNotice).toBeUndefined()
  expect(f.controller.snapshot().error).toContain('重新读取')
  expect(f.executeTool).toHaveBeenCalledOnce()
})

it('rechecks document identity after hashing even when historical content equals the current plan', async () => {
  const f = fixture()
  await f.controller.refresh()
  let finish!: (value: ArrayBuffer) => void
  vi.spyOn(crypto.subtle, 'digest').mockImplementationOnce(
    () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const pending = f.controller.editPlan!(3, { kind: 'restore', revision: 1 })
  await vi.waitFor(() => expect(finish).toBeTypeOf('function'))
  f.documentId.mockResolvedValue('other')
  finish(
    Uint8Array.from(createHash('sha256').update(canonicalPresentationValue(f.original)).digest())
      .buffer,
  )
  await pending
  expect(f.controller.snapshot().project).toBeUndefined()
  expect(f.controller.snapshot().planNotice).toBeUndefined()
  expect(f.executeTool.mock.calls.some(([call]) => call.name === 'save_presentation_plan')).toBe(
    false,
  )
})

it('rejects historical content that differs from its saved digest before saving', async () => {
  const f = fixture()
  await f.controller.refresh()
  f.executeTool.mockResolvedValue({
    output: JSON.stringify({
      projectId: f.original.projectId,
      revision: 1,
      plan: { ...f.original, title: 'forged history' },
    }),
    mutated: false,
    isError: false,
    summary: 'wrong',
  })
  await f.controller.editPlan!(3, { kind: 'restore', revision: 1 })
  expect(f.executeTool).toHaveBeenCalledOnce()
  expect(f.controller.snapshot().project?.plan?.revision).toBe(3)
  expect(f.controller.snapshot().planNotice).toBeUndefined()
  expect(f.controller.snapshot().error).toBeDefined()
})

it('does not resume a cancelled historical read or dispatch its subsequent save', async () => {
  const f = fixture()
  await f.controller.refresh()
  let finish!: (value: Awaited<ReturnType<typeof f.executeTool>>) => void
  f.executeTool.mockImplementationOnce(
    () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const pending = f.controller.editPlan!(3, { kind: 'restore', revision: 2 })
  await vi.waitFor(() => expect(finish).toBeTypeOf('function'))
  f.controller.cancel()
  finish({
    output: JSON.stringify({ projectId: f.original.projectId, revision: 2, plan: f.plans[1] }),
    mutated: false,
    isError: false,
    summary: 'late',
  })
  await pending
  expect(f.executeTool).toHaveBeenCalledOnce()
  expect(f.controller.snapshot().project).toBeUndefined()
  expect(f.controller.snapshot().planNotice).toBeUndefined()
})
