import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createPresentationPlanningSkill } from '../src/skills/powerpoint/presentation-planning.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
import {
  createStructuredProposalController,
  type StructuredProposalController,
} from '../src/agent/proposal-controller.js'

async function confirm(proposals: StructuredProposalController, id: string) {
  await vi.waitFor(() => expect(proposals.pending()?.lockReview?.state).not.toBe('checking'))
  const decision = proposals.waitForDecision(id)
  await proposals.confirm(id)
  return decision
}

it('copies a visibly approved source preference through the actual runtime, preserves provenance after restart and lets a target plan consume the suggestion', async () => {
  const root = mkdtempSync(join(tmpdir(), 'preference-cross-project-'))
  let runtime: ReturnType<typeof createOfficeHostRuntime> | undefined
  let service = createPresentationService({ userDataPath: root })
  const request = vi.fn(async (body: unknown, signal?: AbortSignal) => {
    const text = Buffer.from(await service(body, signal ?? new AbortController().signal)).toString(
      'utf8',
    )
    return new Response(text, { status: JSON.parse(text).error ? 400 : 200 })
  })
  try {
    const sourceProposals = createStructuredProposalController()
    const source = createPresentationPlanningSkill({
      vfs: new InMemoryVfs(),
      available: () => true,
      documentId: async () => 'source-doc',
      lastProject: () => 'source-project',
      rememberProject: async () => {},
      request,
      proposals: sourceProposals,
      listChangeHistory: () => [
        {
          id: 'text:edit',
          kind: 'text',
          sequence: 1,
          legacy: false,
          record: {
            version: 1,
            changeId: 'edit',
            documentId: 'source-doc',
            projectId: 'source-project',
            requestId: 'source-request',
            artifactDigest: 'a'.repeat(64),
            pageId: 'page',
            hostSlideId: 'slide',
            shapeId: 'shape',
            before: '长标题',
            after: '短标题',
            state: 'applied',
          },
        },
      ],
    })
    const saved = await source.executeTool({
      id: 'save-source',
      name: 'save_presentation_preference',
      input: { project_id: 'source-project', change_id: 'edit', preference: '标题尽量简短' },
    })
    expect(saved.isError, saved.output).not.toBe(true)
    expect(request).not.toHaveBeenCalled()
    expect((await confirm(sourceProposals, JSON.parse(saved.output).proposalId)).status).toBe(
      'confirmed',
    )
    const invalidateQa = vi.fn(async () => {})
    vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
    runtime = createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        documentId: async () => 'target-doc',
        request,
        lastProject: () => 'target-project',
        rememberProject: async () => {},
        invalidateQa,
      },
    })
    expect(runtime.skill.tools.map((tool) => tool.name)).toContain('import_presentation_preference')
    const imported = await runtime.skill.executeTool({
      id: 'import',
      name: 'import_presentation_preference',
      input: {
        source_document_id: 'source-doc',
        source_project_id: 'source-project',
        source_change_id: 'edit',
        project_id: 'target-project',
      },
    })
    expect(imported.isError, imported.output).not.toBe(true)
    expect(JSON.parse(imported.output).status).toBe('awaiting_confirmation')
    expect(
      request.mock.calls.some(
        ([body]) => (body as { operation: string }).operation === 'preference_import',
      ),
    ).toBe(false)
    expect(
      await (
        await request({
          operation: 'preference_list',
          documentId: 'target-doc',
          projectId: 'target-project',
        })
      ).json(),
    ).toEqual({ preferences: [] })
    const proposalId = JSON.parse(imported.output).proposalId
    expect(
      (await confirm(runtime.proposals as StructuredProposalController, proposalId)).status,
    ).toBe('confirmed')
    expect(invalidateQa).not.toHaveBeenCalled()
    service = createPresentationService({ userDataPath: root })
    const listed = await runtime.skill.executeTool({
      id: 'list-target',
      name: 'list_presentation_preferences',
      input: { project_id: 'target-project' },
    })
    expect(listed.isError, listed.output).not.toBe(true)
    const preferences = JSON.parse(listed.output).preferences
    expect(preferences).toHaveLength(1)
    expect(preferences[0]).toMatchObject({
      projectId: 'target-project',
      text: '标题尽量简短',
      reuse: {
        version: 1,
        source: { documentId: 'source-doc', projectId: 'source-project', changeId: 'edit' },
        approvalId: proposalId,
      },
    })
    expect(new Date(preferences[0].reuse.approvedAt).toISOString()).toBe(
      preferences[0].reuse.approvedAt,
    )
    const plan = benchmarkPlan()
    plan.projectId = 'target-project'
    plan.brief.constraints.push(preferences[0].text)
    const result = await runtime.skill.executeTool({
      id: 'plan',
      name: 'save_presentation_plan',
      input: { expected_revision: 0, plan },
    })
    expect(result.isError, result.output).not.toBe(true)
    const read = await runtime.skill.executeTool({
      id: 'read-plan',
      name: 'read_presentation_plan',
      input: { project_id: 'target-project' },
    })
    expect(JSON.parse(read.output).plan.brief.constraints).toContain('标题尽量简短')
    expect(
      JSON.parse(read.output).plan.claims.every(
        (claim: { reviewStatus: string }) => claim.reviewStatus === 'needs_review',
      ),
    ).toBe(true)
    expect(
      await (await request({ operation: 'brand_kit_list', documentId: 'target-doc' })).json(),
    ).toEqual({ brandKits: [] })
    const deleted = await runtime.skill.executeTool({
      id: 'delete-target',
      name: 'delete_presentation_preference',
      input: { project_id: 'target-project', change_id: preferences[0].changeId },
    })
    expect(deleted.isError, deleted.output).not.toBe(true)
    expect(
      (
        await confirm(
          runtime.proposals as StructuredProposalController,
          JSON.parse(deleted.output).proposalId,
        )
      ).status,
    ).toBe('confirmed')
    expect(
      await (
        await request({
          operation: 'preference_list',
          documentId: 'target-doc',
          projectId: 'target-project',
        })
      ).json(),
    ).toEqual({ preferences: [] })
    expect(
      await (
        await request({
          operation: 'preference_get',
          documentId: 'source-doc',
          projectId: 'source-project',
          changeId: 'edit',
        })
      ).json(),
    ).toEqual({
      preference: { projectId: 'source-project', changeId: 'edit', text: '标题尽量简短' },
    })
  } finally {
    runtime?.dispose()
    vi.unstubAllGlobals()
    rmSync(root, { recursive: true, force: true })
  }
})

it('reconciles a lost actual PC import receipt through a new runtime proposal without duplicating or replacing the first approval', async () => {
  const root = mkdtempSync(join(tmpdir(), 'preference-import-lost-ack-'))
  let runtime: ReturnType<typeof createOfficeHostRuntime> | undefined
  let service = createPresentationService({ userDataPath: root }),
    loseReceipt = true
  const request = async (body: unknown, signal?: AbortSignal) => {
    const text = Buffer.from(await service(body, signal ?? new AbortController().signal)).toString(
      'utf8',
    )
    if ((body as { operation: string }).operation === 'preference_import' && loseReceipt) {
      loseReceipt = false
      throw new Error('Synthetic lost PC acknowledgement')
    }
    return new Response(text, { status: JSON.parse(text).error ? 400 : 200 })
  }
  try {
    await request({
      operation: 'preference_save',
      documentId: 'source-doc',
      preference: { projectId: 'source-project', changeId: 'edit', text: 'Short titles' },
    })
    vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
    const invalidateQa = vi.fn(async () => {})
    const open = () =>
      createOfficeHostRuntime('powerpoint', {
        presentation: {
          available: () => true,
          documentId: async () => 'target-doc',
          request,
          lastProject: () => 'target-project',
          rememberProject: async () => {},
          invalidateQa,
        },
      })
    runtime = open()
    const call = {
      id: 'copy',
      name: 'import_presentation_preference',
      input: {
        source_document_id: 'source-doc',
        source_project_id: 'source-project',
        source_change_id: 'edit',
        project_id: 'target-project',
      },
    }
    const first = await runtime.skill.executeTool(call)
    expect(first.isError, first.output).not.toBe(true)
    const firstApprovalId = JSON.parse(first.output).proposalId
    await expect(
      confirm(runtime.proposals as StructuredProposalController, firstApprovalId),
    ).rejects.toThrow('Synthetic lost PC acknowledgement')
    const copied = await (
      await request({
        operation: 'preference_list',
        documentId: 'target-doc',
        projectId: 'target-project',
      })
    ).json()
    expect(copied.preferences).toHaveLength(1)
    expect(copied.preferences[0].reuse.approvalId).toBe(firstApprovalId)
    runtime.dispose()
    service = createPresentationService({ userDataPath: root })
    runtime = open()
    const second = await runtime.skill.executeTool({ ...call, id: 'retry' })
    expect(second.isError, second.output).not.toBe(true)
    const secondApprovalId = JSON.parse(second.output).proposalId
    expect(secondApprovalId).not.toBe(firstApprovalId)
    expect(
      (await confirm(runtime.proposals as StructuredProposalController, secondApprovalId)).status,
    ).toBe('confirmed')
    const listed = await runtime.skill.executeTool({
      id: 'list',
      name: 'list_presentation_preferences',
      input: { project_id: 'target-project' },
    })
    expect(listed.isError, listed.output).not.toBe(true)
    expect(JSON.parse(listed.output)).toEqual(copied)
    expect(invalidateQa).not.toHaveBeenCalled()
  } finally {
    runtime?.dispose()
    vi.unstubAllGlobals()
    rmSync(root, { recursive: true, force: true })
  }
})
