import { describe, expect, it, vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'

function fixture(runtime = false) {
  let plan = benchmarkPlan()
  let revision = 1
  const documentId = vi.fn(async () => 'doc')
  const project = () => ({
    projectId: plan.projectId,
    title: plan.title,
    status: 'planned',
    slideCount: plan.slides.length,
    slides: plan.slides.map(({ id, title }) => ({ id, title })),
    history: [],
    plan: { revision, value: plan },
  })
  const request = vi.fn(async (body: unknown) => {
    const input = body as { operation: string; expectedRevision: number; plan: typeof plan }
    if (input.operation === 'save_plan') {
      if (input.expectedRevision !== revision)
        return new Response(JSON.stringify({ error: 'revision_conflict' }))
      plan = structuredClone(input.plan)
      revision++
      return new Response(JSON.stringify({ projectId: plan.projectId, revision, plan }))
    }
    return new Response(JSON.stringify(project()))
  })
  const executeTool = vi.fn(async (call: { name: string; input: Record<string, unknown> }) => {
    const response = await request({
      operation: 'save_plan',
      expectedRevision: call.input.expected_revision,
      plan: call.input.plan,
    })
    const value = await response.json()
    return {
      output: value.error ? `presentation_${value.error}` : JSON.stringify(value),
      isError: !!value.error,
      mutated: false,
      summary: '计划保存',
    }
  })
  const options = {
    request,
    documentId,
    available: () => true,
    lastProject: () => plan.projectId,
    rememberProject: async () => undefined,
  }
  const host = runtime
    ? createOfficeHostRuntime('powerpoint', { presentation: options })
    : undefined
  const controller =
    host?.presentation ?? createPresentationProjectController({ ...options, executeTool })
  return {
    controller,
    request,
    executeTool,
    documentId,
    host,
    current: () => ({ plan, revision }),
    setPlan: (value: typeof plan) => {
      plan = value
    },
    concurrent: () => {
      revision++
    },
  }
}

describe('saved plan adjustments', () => {
  it('keeps required domain sections and the last page without persisting invalid plans', async () => {
    const f = fixture()
    const plan = f.current().plan
    plan.domain = 'report'
    const sections = [
      'executive_summary',
      'findings',
      'supporting_evidence',
      'risks',
      'actions',
    ] as const
    plan.slides.forEach((page, index) => {
      page.domainSection = sections[Math.min(index, 4)]
    })
    await f.controller.refresh()
    await f.controller.editPlan!(1, { kind: 'delete', pageId: plan.slides[0]!.id })
    expect(f.executeTool).not.toHaveBeenCalled()
    expect(f.controller.snapshot().error).toContain('章节')
    const single = benchmarkPlan()
    single.slides = single.slides.slice(0, 1)
    f.setPlan(single)
    await f.controller.refresh()
    await f.controller.editPlan!(1, { kind: 'delete', pageId: single.slides[0]!.id })
    expect(f.executeTool).not.toHaveBeenCalled()
    expect(f.controller.snapshot().error).toContain('最后一页')
    await f.controller.editPlan!(1, { kind: 'move', pageId: single.slides[0]!.id, direction: 'up' })
    expect(f.controller.snapshot().phase).toBe('idle')
    expect(f.executeTool).not.toHaveBeenCalled()
  })

  it('rejects a save acknowledgement for different content without showing success', async () => {
    const f = fixture()
    await f.controller.refresh()
    f.executeTool.mockResolvedValue({
      output: JSON.stringify({
        projectId: f.current().plan.projectId,
        revision: 2,
        plan: f.current().plan,
      }),
      isError: false,
      mutated: false,
      summary: 'wrong plan',
    })
    await f.controller.editPlan!(1, { kind: 'delete', pageId: f.current().plan.slides[1]!.id })
    expect(f.controller.snapshot().project?.plan?.revision).toBe(1)
    expect(f.controller.snapshot().planNotice).toBeUndefined()
    expect(f.controller.snapshot().error).toBeDefined()
  })
  it('saves reordered and deleted pages through the actual host runtime while preserving research', async () => {
    const f = fixture(true)
    await f.controller.refresh()
    expect(f.controller.snapshot().project?.plan?.revision).toBe(1)
    const initial = structuredClone(f.current().plan)
    await f.controller.editPlan!(1, {
      kind: 'move',
      pageId: initial.slides[1]!.id,
      direction: 'down',
    })
    expect(f.current().revision).toBe(2)
    expect(f.current().plan.slides[2]!.id).toBe(initial.slides[1]!.id)
    expect(f.controller.snapshot().project?.plan?.revision).toBe(2)
    await f.controller.editPlan!(2, { kind: 'delete', pageId: initial.slides[1]!.id })
    expect(f.current().revision).toBe(3)
    expect(f.current().plan.slides).toHaveLength(initial.slides.length - 1)
    expect(f.current().plan.sources).toEqual(initial.sources)
    expect(f.current().plan.claims).toEqual(initial.claims)
    expect(f.current().plan.style).toEqual(initial.style)
    expect(f.controller.snapshot().planNotice).toContain('已有 PowerPoint 页面保留')
    expect(
      f.request.mock.calls
        .filter(([body]) => (body as { operation: string }).operation !== 'status')
        .map(([body]) => (body as { operation: string }).operation),
    ).toEqual(['save_plan', 'save_plan'])
    f.host!.dispose()
  })

  it('rejects removal or reordering of a required predecessor without a save', async () => {
    const f = fixture()
    const plan = f.current().plan
    plan.slides[1]!.dependsOn = [plan.slides[0]!.id]
    await f.controller.refresh()
    await f.controller.editPlan!(1, { kind: 'delete', pageId: plan.slides[0]!.id })
    expect(f.executeTool).not.toHaveBeenCalled()
    expect(f.controller.snapshot().error).toContain('依赖')
    await f.controller.editPlan!(1, { kind: 'move', pageId: plan.slides[1]!.id, direction: 'up' })
    expect(f.executeTool).not.toHaveBeenCalled()
    expect(f.current().revision).toBe(1)
  })

  it('rejects stale UI revisions and concurrent PC changes without overwriting them', async () => {
    const f = fixture()
    await f.controller.refresh()
    const pageId = f.current().plan.slides[1]!.id
    await f.controller.editPlan!(0, { kind: 'delete', pageId })
    expect(f.executeTool).not.toHaveBeenCalled()
    expect(f.controller.snapshot().error).toContain('重新读取')
    f.concurrent()
    await f.controller.editPlan!(1, { kind: 'delete', pageId })
    expect(f.current().revision).toBe(2)
    expect(f.current().plan.slides.some((page) => page.id === pageId)).toBe(true)
    expect(f.controller.snapshot().error).toContain('重新读取')
  })

  it('does not save into a changed document or revive a cancelled operation', async () => {
    const f = fixture()
    await f.controller.refresh()
    const pageId = f.current().plan.slides[1]!.id
    f.documentId.mockResolvedValue('other')
    await f.controller.editPlan!(1, { kind: 'delete', pageId })
    expect(f.executeTool).not.toHaveBeenCalled()
    expect(f.controller.snapshot().project).toBeUndefined()
    f.documentId.mockResolvedValue('doc')
    await f.controller.refresh()
    let finish!: () => void
    f.executeTool.mockImplementation(
      () =>
        new Promise((resolve) => {
          finish = () =>
            resolve({
              output: JSON.stringify({
                projectId: f.current().plan.projectId,
                revision: 2,
                plan: f.current().plan,
              }),
              isError: false,
              mutated: false,
              summary: 'late',
            })
        }),
    )
    const pending = f.controller.editPlan!(1, { kind: 'delete', pageId })
    await vi.waitFor(() => expect(finish).toBeTypeOf('function'))
    f.controller.cancel()
    finish()
    await pending
    expect(f.controller.snapshot().project).toBeUndefined()
    expect(f.controller.snapshot().planNotice).toBeUndefined()
  })
})
