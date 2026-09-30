import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, it, vi } from 'vitest'
import { createPresentationService } from '../../shell/src/main/presentation-service.js'
import { createPresentationManualObservationSkill } from '../src/skills/powerpoint/presentation-manual-observations.js'
import { createPresentationPlanningSkill } from '../src/skills/powerpoint/presentation-planning.js'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { createOfficeHostRuntime } from '../src/agent/host-runtime.js'
import { BrowserPresentationBaselineAdapter } from '../src/skills/powerpoint/browser-presentation-baseline-adapter.js'
import { InMemoryVfs } from '../src/skills/shared/vfs.js'
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

it('persists explicit before/after host observations across restarts, requires visible preference approval, and preserves origin through cross-project planning reuse', async () => {
  const root = mkdtempSync(join(tmpdir(), 'manual-observation-integration-'))
  let service = createPresentationService({ userDataPath: root })
  let documentId = 'source-doc'
  let loseResponseFor: string | undefined
  let failBeforeFor: string | undefined
  let shape = {
    id: 'shape',
    name: 'Title',
    type: 'TextBox',
    left: 10,
    top: 20,
    width: 100,
    height: 30,
    rotation: 0,
    text: '一个比较长的标题',
    font: {
      name: 'Arial',
      size: 20,
      color: '#000000',
      bold: false,
      italic: false,
      underline: 'None',
    },
  }
  const adapter = {
    readContext: async () => ({
      slideIds: ['slide'],
      selectedSlideIds: ['slide'],
      selectedShapeIds: ['shape'],
    }),
    readPage: async () => ({ slideId: 'slide', shapes: [structuredClone(shape)] }),
  }
  const request = async (body: unknown, signal?: AbortSignal) => {
    if ((body as { operation: string }).operation === failBeforeFor) {
      failBeforeFor = undefined
      throw new Error('request failed before PC commit')
    }
    const result = Buffer.from(
      await service(body, signal ?? new AbortController().signal),
    ).toString('utf8')
    if ((body as { operation: string }).operation === loseResponseFor) {
      loseResponseFor = undefined
      throw new Error('response lost after PC commit')
    }
    return new Response(result, { status: JSON.parse(result).error ? 400 : 200 })
  }
  const proposals = createStructuredProposalController()
  const make = () =>
    createPresentationManualObservationSkill({
      adapter,
      available: () => true,
      documentId: async () => documentId,
      request,
      proposals,
    })
  let skill = make()
  const invoke = (name: string, input: Record<string, unknown>) =>
    skill.executeTool({ id: name, name, input })
  try {
    loseResponseFor = 'manual_observation_begin'
    const before = await invoke('begin_presentation_edit_observation', {
      project_id: 'source-project',
      slide_id: 'slide',
      shape_id: 'shape',
    })
    expect(before.isError, before.output).not.toBe(true)
    const observationId = JSON.parse(before.output).observation.observationId
    skill.clear()
    service = createPresentationService({ userDataPath: root })
    skill = make()
    shape = { ...shape, text: '短标题', left: 30, rotation: 45, font: { ...shape.font, size: 24 } }
    loseResponseFor = 'manual_observation_complete'
    const completed = await invoke('complete_presentation_edit_observation', {
      project_id: 'source-project',
      observation_id: observationId,
    })
    expect(completed.isError, completed.output).not.toBe(true)
    const observation = JSON.parse(completed.output).observation
    expect(observation.before.shape.text).toBe('一个比较长的标题')
    expect(observation.after.shape.text).toBe('短标题')
    expect(observation.after.shape.left).toBe(30)
    expect(observation.after.shape.rotation).toBe(45)
    expect(observation.source).toBe('host_difference_unattributed')
    expect(observation.atomicSnapshot).toBe(false)
    const listPreferences = async () =>
      (
        await (
          await request({
            operation: 'preference_list',
            documentId: 'source-doc',
            projectId: 'source-project',
          })
        ).json()
      ).preferences
    expect(await listPreferences()).toEqual([])
    const proposal = await invoke('save_presentation_observed_preference', {
      project_id: 'source-project',
      observation_id: observationId,
      preference: '标题简短，字号24',
    })
    expect(proposal.isError, proposal.output).not.toBe(true)
    expect(await listPreferences()).toEqual([])
    failBeforeFor = 'preference_save_observation'
    await expect(confirm(proposals, JSON.parse(proposal.output).proposalId)).rejects.toThrow(
      'request failed before PC commit',
    )
    expect(await listPreferences()).toEqual([])
    const retryProposal = await invoke('save_presentation_observed_preference', {
      project_id: 'source-project',
      observation_id: observationId,
      preference: '标题简短，字号24',
    })
    expect(retryProposal.isError, retryProposal.output).not.toBe(true)
    loseResponseFor = 'preference_save_observation'
    expect((await confirm(proposals, JSON.parse(retryProposal.output).proposalId)).status).toBe(
      'confirmed',
    )
    const [saved] = await listPreferences()
    expect(saved.origin).toMatchObject({
      version: 1,
      observationId,
      beforeDigest: observation.before.digest,
      afterDigest: observation.after.digest,
    })
    const deletion = await invoke('delete_presentation_edit_observation', {
      project_id: 'source-project',
      observation_id: observationId,
    })
    expect(deletion.isError, deletion.output).not.toBe(true)
    loseResponseFor = 'manual_observation_delete'
    expect((await confirm(proposals, JSON.parse(deletion.output).proposalId)).status).toBe(
      'confirmed',
    )
    expect(
      (
        await (
          await request({
            operation: 'manual_observation_list',
            documentId: 'source-doc',
            projectId: 'source-project',
          })
        ).json()
      ).observations,
    ).toEqual([])
    expect(await listPreferences()).toEqual([saved])
    const rotationBefore = await invoke('begin_presentation_edit_observation', {
      project_id: 'source-project',
      slide_id: 'slide',
      shape_id: 'shape',
    })
    expect(rotationBefore.isError, rotationBefore.output).not.toBe(true)
    const rotationId = JSON.parse(rotationBefore.output).observation.observationId
    shape = { ...shape, rotation: 90 }
    const rotationComplete = await invoke('complete_presentation_edit_observation', {
      project_id: 'source-project',
      observation_id: rotationId,
    })
    expect(rotationComplete.isError, rotationComplete.output).not.toBe(true)
    const rotated = JSON.parse(rotationComplete.output).observation
    expect(rotated.before.shape.rotation).toBe(45)
    expect(rotated.after.shape.rotation).toBe(90)
    expect(rotated.after.shape.text).toBe(rotated.before.shape.text)
    expect(rotated.after.digest).not.toBe(rotated.before.digest)
    skill.clear()
    service = createPresentationService({ userDataPath: root })
    const persistedRotation = await (
      await request({
        operation: 'manual_observation_get',
        documentId: 'source-doc',
        projectId: 'source-project',
        observationId: rotationId,
      })
    ).json()
    expect(persistedRotation.observation.after.shape.rotation).toBe(90)
    documentId = 'target-doc'
    const targetProposals = createStructuredProposalController()
    const target = createPresentationPlanningSkill({
      vfs: new InMemoryVfs(),
      available: () => true,
      documentId: async () => documentId,
      lastProject: () => 'target-project',
      rememberProject: async () => {},
      request,
      proposals: targetProposals,
    })
    const imported = await target.executeTool({
      id: 'import',
      name: 'import_presentation_preference',
      input: {
        source_document_id: 'source-doc',
        source_project_id: 'source-project',
        source_change_id: saved.changeId,
        project_id: 'target-project',
      },
    })
    expect(imported.isError, imported.output).not.toBe(true)
    expect((await confirm(targetProposals, JSON.parse(imported.output).proposalId)).status).toBe(
      'confirmed',
    )
    const listed = await target.executeTool({
      id: 'list',
      name: 'list_presentation_preferences',
      input: { project_id: 'target-project' },
    })
    expect(listed.isError, listed.output).not.toBe(true)
    const [copy] = JSON.parse(listed.output).preferences
    expect(copy.origin).toEqual(saved.origin)
    expect(copy.reuse.source).toEqual({
      documentId: 'source-doc',
      projectId: 'source-project',
      changeId: saved.changeId,
    })
    const plan = structuredClone(benchmarkPlan())
    plan.projectId = 'target-project'
    plan.brief.constraints.push(copy.text)
    const planned = await target.executeTool({
      id: 'plan',
      name: 'save_presentation_plan',
      input: { expected_revision: 0, plan },
    })
    expect(planned.isError, planned.output).not.toBe(true)
    const loaded = await target.executeTool({
      id: 'read',
      name: 'read_presentation_plan',
      input: { project_id: JSON.parse(planned.output).projectId },
    })
    expect(loaded.isError, loaded.output).not.toBe(true)
    expect(JSON.parse(loaded.output).plan.brief.constraints).toContain(copy.text)
    // The test explicitly selects the suggestion; it does not claim autonomous learning or host QA.
    expect(
      (await (await request({ operation: 'brand_kit_list', documentId: 'target-doc' })).json())
        .brandKits,
    ).toEqual([])
    expect(
      JSON.parse(loaded.output).plan.claims.every(
        (claim: { reviewStatus: string }) => claim.reviewStatus === 'needs_review',
      ),
    ).toBe(true)
    expect(shape.text).toBe('短标题')
    expect(shape.left).toBe(30)
  } finally {
    skill.clear()
    rmSync(root, { recursive: true, force: true })
  }
})

it('registers the full observation workflow in the actual host runtime and keeps local approval outside host writes and QA invalidation', async () => {
  const root = mkdtempSync(join(tmpdir(), 'manual-observation-runtime-'))
  const service = createPresentationService({ userDataPath: root })
  let runtime: ReturnType<typeof createOfficeHostRuntime> | undefined
  let shape = {
    id: 'shape',
    name: 'Title',
    type: 'TextBox',
    left: 10,
    top: 20,
    width: 100,
    height: 30,
    text: 'Before',
  }
  vi.stubGlobal('Office', { context: { requirements: { isSetSupported: () => true } } })
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readContext').mockImplementation(
    async () => ({ slideIds: ['slide'], selectedSlideIds: ['slide'], selectedShapeIds: ['shape'] }),
  )
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readPage').mockImplementation(
    async () => ({ slideId: 'slide', shapes: [structuredClone(shape)] }),
  )
  const invalidateQa = vi.fn(async () => {})
  const request = async (body: unknown, signal?: AbortSignal) => {
    const result = Buffer.from(
      await service(body, signal ?? new AbortController().signal),
    ).toString('utf8')
    return new Response(result, { status: JSON.parse(result).error ? 400 : 200 })
  }
  const make = () =>
    createOfficeHostRuntime('powerpoint', {
      presentation: {
        available: () => true,
        documentId: async () => 'doc',
        request,
        lastProject: () => 'project',
        rememberProject: async () => {},
        invalidateQa,
      },
    })
  const call = (name: string, input: Record<string, unknown>) =>
    runtime!.skill.executeTool({ id: name, name, input })
  try {
    runtime = make()
    expect(runtime.skill.tools.map((tool) => tool.name)).toContain(
      'begin_presentation_edit_observation',
    )
    const result = await call('begin_presentation_edit_observation', {
      project_id: 'project',
      slide_id: 'slide',
      shape_id: 'shape',
    })
    expect(result.isError, result.output).not.toBe(true)
    const observationId = JSON.parse(result.output).observation.observationId
    runtime.dispose()
    runtime = make()
    shape = { ...shape, text: 'After' }
    const completed = await call('complete_presentation_edit_observation', {
      project_id: 'project',
      observation_id: observationId,
    })
    expect(completed.isError, completed.output).not.toBe(true)
    const saved = await call('save_presentation_observed_preference', {
      project_id: 'project',
      observation_id: observationId,
      preference: 'Short titles',
    })
    expect(saved.isError, saved.output).not.toBe(true)
    expect(
      (
        await confirm(
          runtime.proposals as StructuredProposalController,
          JSON.parse(saved.output).proposalId,
        )
      ).status,
    ).toBe('confirmed')
    expect(invalidateQa).not.toHaveBeenCalled()
    const listed = await call('list_presentation_preferences', { project_id: 'project' })
    expect(JSON.parse(listed.output).preferences).toHaveLength(1)
    expect(JSON.parse(listed.output).preferences[0].origin.observationId).toBe(observationId)
  } finally {
    runtime?.dispose()
    vi.restoreAllMocks()
    vi.unstubAllGlobals()
    rmSync(root, { recursive: true, force: true })
  }
})
