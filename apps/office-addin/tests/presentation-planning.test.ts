import { describe, expect, it, vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationPlanningSkill } from '../src/skills/powerpoint/presentation-planning.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
const plan = benchmarkPlan()
function setup() {
  const vfs = new InMemoryVfs()
  const request = vi.fn(
    async (_body: unknown, _signal?: AbortSignal) =>
      new Response(JSON.stringify({ projectId: plan.projectId, revision: 1, plan })),
  )
  const rememberProject = vi.fn(async () => {})
  const documentId = vi.fn(async () => 'doc-1')
  const available = vi.fn(() => true)
  const skill = createPresentationPlanningSkill({
    vfs,
    request,
    rememberProject,
    documentId,
    available,
    lastProject: () => plan.projectId,
  })
  return { skill, vfs, request, rememberProject, documentId, available }
}
describe('saved presentation planning tools', () => {
  it('saves a versioned plan and returns exact compile claim mapping', async () => {
    const f = setup()
    const result = await f.skill.executeTool({
      id: 'save',
      name: 'save_presentation_plan',
      input: { expected_revision: 0, plan },
    })
    expect(result.isError).not.toBe(true)
    expect(f.request).toHaveBeenCalledWith(
      {
        operation: 'save_plan',
        documentId: 'doc-1',
        projectId: plan.projectId,
        expectedRevision: 0,
        plan,
      },
      undefined,
    )
    expect(JSON.parse(result.output)).toMatchObject({
      revision: 1,
      plan,
      compileClaims: expect.any(Array),
    })
    expect(f.vfs.list('/home/user')).toContain(`/home/user/generated/${plan.projectId}.plan.json`)
    expect(f.rememberProject).toHaveBeenCalledWith(plan.projectId)
  })
  it('reloads the current saved plan without compiling or importing', async () => {
    const f = setup()
    const result = await f.skill.executeTool({
      id: 'read',
      name: 'read_presentation_plan',
      input: {},
    })
    expect(result.isError).not.toBe(true)
    expect(f.request).toHaveBeenCalledWith(
      { operation: 'get_plan', documentId: 'doc-1', projectId: plan.projectId },
      undefined,
    )
    expect(result.mutated).toBe(false)
  })
  it('does not replace local evidence on revision conflicts or unknown fields', async () => {
    const f = setup()
    f.request.mockResolvedValue(new Response(JSON.stringify({ error: 'revision_conflict' })))
    expect(
      await f.skill.executeTool({
        id: 'save',
        name: 'save_presentation_plan',
        input: { expected_revision: 0, plan },
      }),
    ).toMatchObject({ isError: true, output: 'presentation_revision_conflict' })
    expect(f.vfs.list('/home/user')).toEqual([])
    expect(
      await f.skill.executeTool({
        id: 'bad',
        name: 'read_presentation_plan',
        input: { injected: true },
      }),
    ).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  })
  it('rejects a foreign plan response and failed read preserves project selection', async () => {
    const f = setup()
    f.request.mockResolvedValue(
      new Response(JSON.stringify({ projectId: 'other', revision: 1, plan })),
    )
    expect(
      await f.skill.executeTool({ id: 'read', name: 'read_presentation_plan', input: {} }),
    ).toMatchObject({ isError: true })
    expect(f.rememberProject).not.toHaveBeenCalled()
    expect(f.vfs.list('/home/user')).toEqual([])
  })
  it('does not publish a late plan after clear or a document change', async () => {
    const f = setup()
    f.request.mockImplementation(async () => {
      f.skill.clear()
      return new Response(JSON.stringify({ projectId: plan.projectId, revision: 1, plan }))
    })
    expect(
      await f.skill.executeTool({ id: 'read', name: 'read_presentation_plan', input: {} }),
    ).toMatchObject({ isError: true, output: 'cancelled' })
    expect(f.vfs.list('/home/user')).toEqual([])
    const g = setup()
    g.documentId.mockResolvedValueOnce('doc-1').mockResolvedValue('doc-2')
    expect(
      await g.skill.executeTool({ id: 'read', name: 'read_presentation_plan', input: {} }),
    ).toMatchObject({ isError: true, output: 'presentation_document_changed' })
    expect(g.vfs.list('/home/user')).toEqual([])
  })
})

it('does not lose an existing project selection when saving a foreign new project is rejected', async () => {
  const f = setup()
  f.request.mockResolvedValue(new Response(JSON.stringify({ error: 'document_mismatch' })))
  expect(
    await f.skill.executeTool({
      id: 'foreign',
      name: 'save_presentation_plan',
      input: { expected_revision: 0, plan: { ...plan, projectId: 'foreign-project' } },
    }),
  ).toMatchObject({ isError: true })
  expect(f.rememberProject).not.toHaveBeenCalled()
})
