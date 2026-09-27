import { describe, expect, it, vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationPlanningSkill } from '../src/skills/powerpoint/presentation-planning.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import type { PresentationHistoryEntry } from '../src/skills/powerpoint/presentation-change-history.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
const plan = benchmarkPlan()
function setup(history?: PresentationHistoryEntry[]) {
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
    listChangeHistory: history ? () => history : undefined,
  })
  return { skill, vfs, request, rememberProject, documentId, available }
}
describe('saved presentation planning tools', () => {
  it('exposes applied edit candidates without changing the plan or brand kit', async () => {
    const record = { version: 1 as const, changeId: 'edit1', documentId: 'doc-1', projectId: plan.projectId, requestId: 'request1', artifactDigest: 'a'.repeat(64), pageId: 'page1', hostSlideId: 'slide1', shapeId: 'shape1', before: '长标题', after: '短标题', state: 'applied' as const }
    const f = setup([{ id: 'text:edit1', sequence: 1, legacy: false, kind: 'text', record }])
    const result = await f.skill.executeTool({ id: 'preferences', name: 'read_presentation_preference_candidates', input: { project_id: plan.projectId } })
    expect(result.mutated).toBe(false)
    expect(JSON.parse(result.output).candidates).toMatchObject([{ changeId: 'edit1', status: 'candidate', after: '短标题' }])
    expect(f.request).not.toHaveBeenCalled()
    expect(f.rememberProject).not.toHaveBeenCalled()
  })
  it('requires visible confirmation before saving a preference', async () => {
    const record = { version: 1 as const, changeId: 'edit1', documentId: 'doc-1', projectId: plan.projectId, requestId: 'request1', artifactDigest: 'a'.repeat(64), pageId: 'page1', hostSlideId: 'slide1', shapeId: 'shape1', before: '长标题', after: '短标题', state: 'applied' as const }
    const history: PresentationHistoryEntry[] = [{ id: 'text:edit1', sequence: 1, legacy: false, kind: 'text', record }]
    const proposals = createStructuredProposalController()
    const request = vi.fn(async (body: unknown) => new Response(JSON.stringify({ preference: (body as { preference: unknown }).preference })))
    const skill = createPresentationPlanningSkill({ vfs: new InMemoryVfs(), available: () => true, documentId: async () => 'doc-1', lastProject: () => plan.projectId, rememberProject: async () => {}, request, listChangeHistory: () => history, proposals })
    const result = await skill.executeTool({ id: 'save', name: 'save_presentation_preference', input: { project_id: plan.projectId, change_id: 'edit1', preference: '标题尽量简短' } })
    expect(JSON.parse(result.output).status).toBe('awaiting_confirmation')
    expect(request).not.toHaveBeenCalled()
    await proposals.confirm(JSON.parse(result.output).proposalId)
    expect(request).toHaveBeenCalledWith({ operation: 'preference_save', documentId: 'doc-1', preference: { projectId: plan.projectId, changeId: 'edit1', text: '标题尽量简短' } }, expect.any(AbortSignal))
    const list = await skill.executeTool({ id: 'list', name: 'list_presentation_preferences', input: { project_id: plan.projectId } })
    expect(list.isError).toBe(true) // Reject malformed PC response rather than trusting it as a catalog.
  })
  it('requires visible confirmation before deleting a stored preference', async () => {
    const proposals = createStructuredProposalController()
    let preferences = [{ projectId: plan.projectId, changeId: 'edit1', text: '标题尽量简短' }]
    const request = vi.fn(async (body: unknown) => {
      const operation = (body as { operation: string }).operation
      if (operation === 'preference_list') return new Response(JSON.stringify({ preferences }))
      preferences = []
      return new Response(JSON.stringify({ deleted: true }))
    })
    const skill = createPresentationPlanningSkill({ vfs: new InMemoryVfs(), available: () => true, documentId: async () => 'doc-1', lastProject: () => plan.projectId, rememberProject: async () => {}, request, proposals })
    const proposed = await skill.executeTool({ id: 'delete', name: 'delete_presentation_preference', input: { project_id: plan.projectId, change_id: 'edit1' } })
    expect(preferences).toHaveLength(1)
    await proposals.confirm(JSON.parse(proposed.output).proposalId)
    expect(preferences).toEqual([])
    expect(JSON.parse((await skill.executeTool({ id: 'list', name: 'list_presentation_preferences', input: { project_id: plan.projectId } })).output)).toEqual({ preferences: [] })
  })
  it('returns five local domain planning skills without claiming source verification', async () => {
    const f = setup()
    for (const domain of ['pitch', 'report', 'training', 'research', 'sales']) {
      const result = await f.skill.executeTool({ id: domain, name: 'read_presentation_domain_skill', input: { domain } })
      expect(result.isError).not.toBe(true)
      expect(JSON.parse(result.output)).toMatchObject({ domain, sections: expect.any(Array), questions: expect.any(Array) })
    }
    expect(f.request).not.toHaveBeenCalled()
    expect(await f.skill.executeTool({ id: 'bad', name: 'read_presentation_domain_skill', input: { domain: 'finance' } })).toMatchObject({ isError: true, output: 'invalid_tool_input' })
  })
  it('saves, lists and reads a pinned PC brand kit revision', async () => {
    const f = setup()
    const brandKit = { id: 'research', revision: 1, name: 'Research', allowedColors: ['FFFFFF', '172033'] }
    f.request.mockResolvedValueOnce(new Response(JSON.stringify({ brandKit })))
      .mockResolvedValueOnce(new Response(JSON.stringify({ brandKits: [brandKit] })))
      .mockResolvedValueOnce(new Response(JSON.stringify({ brandKit })))
    expect((await f.skill.executeTool({ id: 'save-kit', name: 'save_presentation_brand_kit', input: { expected_revision: 0, brand_kit: brandKit } })).isError).not.toBe(true)
    expect(f.request).toHaveBeenNthCalledWith(1, { operation: 'brand_kit_save', documentId: 'doc-1', expectedRevision: 0, brandKit }, undefined)
    expect(JSON.parse((await f.skill.executeTool({ id: 'list-kit', name: 'list_presentation_brand_kits', input: {} })).output)).toEqual({ brandKits: [brandKit] })
    expect(JSON.parse((await f.skill.executeTool({ id: 'read-kit', name: 'read_presentation_brand_kit', input: { brand_kit_id: 'research', revision: 1 } })).output)).toEqual({ brandKit })
    expect(f.rememberProject).not.toHaveBeenCalled()
  })
  it('reports brand revision conflicts and rejects invalid catalog responses', async () => {
    const f = setup()
    const brandKit = { id: 'research', revision: 1, name: 'Research', allowedColors: ['FFFFFF'] }
    f.request.mockResolvedValueOnce(new Response(JSON.stringify({ error: 'revision_conflict' })))
      .mockResolvedValueOnce(new Response(JSON.stringify({ brandKits: [{ ...brandKit, revision: 0 }] })))
    expect(await f.skill.executeTool({ id: 'save-kit', name: 'save_presentation_brand_kit', input: { expected_revision: 0, brand_kit: brandKit } })).toMatchObject({ isError: true, output: 'presentation_revision_conflict' })
    expect(await f.skill.executeTool({ id: 'list-kit', name: 'list_presentation_brand_kits', input: {} })).toMatchObject({ isError: true, output: 'presentation_response_invalid' })
  })
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
