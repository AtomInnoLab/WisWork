import { describe, expect, it, vi } from 'vitest'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'

const project = {
  projectId: 'project-1',
  title: '研究汇报',
  status: 'pending',
  latestRequestId: 'request-2',
  latestCompiledRequestId: 'request-1',
  slideCount: 1,
  slides: [{ id: 'slide-1', title: '研究结论' }],
  history: [
    { requestId: 'request-2', sequence: 2, status: 'pending', slideCount: 1 },
    { requestId: 'request-1', sequence: 1, status: 'compiled', slideCount: 1 },
  ],
}
function fixture() {
  const request = vi.fn(
    async (_body: unknown, _signal?: AbortSignal) => new Response(JSON.stringify(project)),
  )
  const executeTool = vi.fn(async () => ({ output: '{}', mutated: false, summary: '已恢复' }))
  const documentId = vi.fn(async () => 'document-1')
  const available = vi.fn(() => true)
  const lastProject = vi.fn((): string | undefined => 'project-1')
  const controller = createPresentationProjectController({
    request,
    executeTool,
    documentId,
    available,
    lastProject,
  })
  return { controller, request, executeTool, documentId, available, lastProject }
}
describe('presentation project controls', () => {
  it('loads persisted page inventory/history and restores or resumes through the generation skill', async () => {
    const f = fixture()
    await f.controller.refresh()
    expect(f.controller.snapshot()).toMatchObject({ phase: 'idle', project })
    await f.controller.restore()
    expect(f.executeTool).toHaveBeenCalledWith(
      expect.objectContaining({
        name: 'restore_presentation_project',
        input: { project_id: 'project-1' },
      }),
      expect.any(AbortSignal),
    )
    await f.controller.resume('request-2')
    expect(f.executeTool).toHaveBeenCalledWith(
      expect.objectContaining({
        name: 'resume_presentation_project',
        input: { project_id: 'project-1', request_id: 'request-2' },
      }),
      expect.any(AbortSignal),
    )
  })
  it('does not create a document identity or call the PC when no project is known', async () => {
    const f = fixture()
    f.lastProject.mockReturnValue(undefined)
    await f.controller.refresh()
    expect(f.controller.snapshot()).toEqual({ phase: 'idle' })
    expect(f.request).not.toHaveBeenCalled()
    expect(f.documentId).not.toHaveBeenCalled()
  })
  it.each(['clear', 'cancel'] as const)(
    'invalidates a late status response after %s',
    async (action) => {
      const f = fixture()
      let done!: (response: Response) => void
      f.request.mockImplementation(
        () =>
          new Promise((resolve) => {
            done = resolve
          }),
      )
      const pending = f.controller.refresh()
      await vi.waitFor(() => expect(f.request).toHaveBeenCalled())
      const signal = f.request.mock.calls[0]![1]!
      f.controller[action]()
      expect(signal.aborted).toBe(true)
      done(new Response(JSON.stringify(project)))
      await pending
      expect(f.controller.snapshot().project).toBeUndefined()
      expect(f.controller.snapshot().phase).toBe('idle')
    },
  )
  it('rejects foreign, malformed, and falsely reviewed project status', async () => {
    for (const value of [
      { ...project, projectId: 'other' },
      { ...project, slideCount: 2 },
      { ...project, checks: { render: 'passed' } },
    ]) {
      const f = fixture()
      f.request.mockResolvedValue(new Response(JSON.stringify(value)))
      await f.controller.refresh()
      expect(f.controller.snapshot().project).toBeUndefined()
      expect(f.controller.snapshot().error).toBeTruthy()
    }
  })
  it('does not publish a status after switching documents or losing capability', async () => {
    const f = fixture()
    f.documentId.mockResolvedValueOnce('document-1').mockResolvedValue('document-2')
    await f.controller.refresh()
    expect(f.controller.snapshot().project).toBeUndefined()
    expect(f.controller.snapshot().error).toContain('文档')
    const g = fixture()
    g.request.mockImplementation(async () => {
      g.available.mockReturnValue(false)
      return new Response(JSON.stringify(project))
    })
    await g.controller.refresh()
    expect(g.controller.snapshot().project).toBeUndefined()
  })
  it('explains old-PC compatibility without exposing a raw code', async () => {
    const f = fixture()
    f.request.mockResolvedValue(new Response(JSON.stringify({ error: 'invalid_request' })))
    await f.controller.refresh()
    expect(f.controller.snapshot().error).toContain('升级')
    expect(f.controller.snapshot().error).not.toContain('invalid_request')
  })
  it('does not duplicate an active action and aborts it on clear', async () => {
    const f = fixture()
    await f.controller.refresh()
    let done!: () => void
    f.executeTool.mockImplementation(async () => {
      await new Promise<void>((resolve) => {
        done = resolve
      })
      return { output: '{}', mutated: false, summary: '已恢复' }
    })
    const pending = f.controller.resume('request-2')
    await vi.waitFor(() => expect(f.executeTool).toHaveBeenCalledOnce())
    await f.controller.resume('request-2')
    expect(f.executeTool).toHaveBeenCalledOnce()
    f.controller.clear()
    done()
    await pending
    expect(f.controller.snapshot()).toEqual({ phase: 'idle' })
  })
  it('rejects an arbitrary stale request instead of resuming a different revision', async () => {
    const f = fixture()
    await f.controller.refresh()
    await f.controller.resume('other-request')
    expect(f.executeTool).not.toHaveBeenCalled()
  })
})

describe('plan-only project status', () => {
  it('accepts a saved plan before compilation without inventing request history or QA', async () => {
    const { benchmarkPlan } =
      await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan.js')
    const plan = benchmarkPlan()
    const value = {
      projectId: plan.projectId,
      title: plan.title,
      status: 'planned',
      slideCount: plan.slides.length,
      slides: plan.slides.map(({ id, title }) => ({ id, title })),
      history: [],
      plan: { revision: 1, value: plan },
    }
    const f = fixture()
    f.lastProject.mockReturnValue(plan.projectId)
    f.request.mockResolvedValue(new Response(JSON.stringify(value)))
    await f.controller.refresh()
    expect(f.controller.snapshot().error).toBeUndefined()
    expect(f.controller.snapshot().project).toEqual(value)
    await f.controller.resume('made-up')
    expect(f.executeTool).not.toHaveBeenCalled()
  })
})
it('projects page production and resumes its explicit request through the page tool', async () => {
  const f = fixture(),
    production = {
      projectId: 'project-1',
      requestId: 'pages-1',
      planRevision: 1,
      status: 'partial',
      compiledCount: 1,
      total: 2,
      pages: [
        { id: 'a', title: 'A', state: 'compiled', attempt: 1 },
        { id: 'b', title: 'B', state: 'failed', attempt: 1, error: 'compile_failed' },
      ],
    }
  f.request.mockImplementation(async () => new Response(JSON.stringify({ ...project, production })))
  await f.controller.refresh()
  expect(f.controller.snapshot().project?.production).toEqual(production)
  await f.controller.runProduction('wrong')
  expect(f.executeTool).not.toHaveBeenCalled()
  await f.controller.runProduction('pages-1')
  expect(f.executeTool).toHaveBeenCalledWith(
    expect.objectContaining({
      name: 'run_presentation_production',
      input: { project_id: 'project-1', request_id: 'pages-1' },
    }),
    expect.any(AbortSignal),
  )
})
