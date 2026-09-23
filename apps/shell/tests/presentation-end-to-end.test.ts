import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { openPptx } from '@wiswork/pptx-engine'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import { createPresentationService } from '../src/main/presentation-service'
import { createPresentationGenerationSkill } from '../../office-addin/src/skills/powerpoint/presentation-generation'
import { InMemoryVfs } from '../../office-addin/src/skills/shared/vfs'

describe('Taskpane to durable PC compilation', () => {
  it('delivers native eight-page PPTX and recovers a lost response without recompiling', async () => {
    const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-presentation-e2e-'))
    try {
      const compile = vi.fn(compilePresentationDeck)
      const service = createPresentationService({ userDataPath, compile })
      const signal = new AbortController().signal
      let loseResponse = true
      const request = async (body: unknown) => {
        const result = await service(body, signal)
        if (loseResponse) {
          loseResponse = false
          throw new Error('connection_lost')
        }
        return new Response(Buffer.from(result))
      }
      const deck = benchmarkDeck()
      const options = {
        available: () => true,
        documentId: async () => 'document-1',
        lastProject: () => deck.id,
        rememberProject: async () => {},
        request,
      }
      const first = createPresentationGenerationSkill({ ...options, vfs: new InMemoryVfs() })
      const call = {
        id: 'compile-1',
        name: 'compile_deck_with_pptxgenjs',
        input: { request_id: 'request-1', deck },
      }
      expect(await first.executeTool(call)).toMatchObject({ isError: true })
      const vfs = new InMemoryVfs()
      const restartedService = createPresentationService({ userDataPath, compile })
      const recreated = createPresentationGenerationSkill({
        ...options,
        vfs,
        request: async (body) => new Response(Buffer.from(await restartedService(body, signal))),
      })
      const outcome = await recreated.executeTool(call)
      expect(outcome.isError).not.toBe(true)
      expect(compile).toHaveBeenCalledOnce()
      const opened = await openPptx(vfs.readBytes(`/home/user/generated/${deck.id}.pptx`))
      expect(opened.deck.slides).toHaveLength(8)
      expect(opened.deck.slides[5]!.elements.some((element) => element.type === 'table')).toBe(true)
      expect(opened.deck.slides[6]!.elements.some((element) => element.type === 'chart')).toBe(true)
      expect(outcome.output).toContain('not_run')
      expect(outcome.output).not.toContain('pptxBase64')
      const wrong = new InMemoryVfs()
      const other = createPresentationGenerationSkill({
        ...options,
        vfs: wrong,
        documentId: async () => 'other-document',
        request: async (body) => new Response(Buffer.from(await restartedService(body, signal))),
      })
      expect(
        await other.executeTool({ id: 'restore', name: 'restore_presentation_project', input: {} }),
      ).toMatchObject({ isError: true, output: 'presentation_document_mismatch' })
      expect(wrong.list('/home/user')).toEqual([])
    } finally {
      rmSync(userDataPath, { recursive: true, force: true })
    }
  })
})

describe('Taskpane project recovery controls', () => {
  it('finds the first failed compile after reopening and resumes its saved input', async () => {
    const { createPresentationProjectController } =
      await import('../../office-addin/src/skills/powerpoint/presentation-project')
    const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-presentation-recovery-'))
    try {
      const compile = vi
        .fn(compilePresentationDeck)
        .mockRejectedValueOnce(new Error('temporary_failure'))
      let service = createPresentationService({ userDataPath, compile })
      let remembered: string | undefined
      const options = {
        available: () => true,
        documentId: async () => 'document-1',
        lastProject: () => remembered,
        rememberProject: async (id: string) => {
          remembered = id
        },
        request: async (body: unknown, signal?: AbortSignal) =>
          new Response(Buffer.from(await service(body, signal ?? new AbortController().signal))),
      }
      const original = createPresentationGenerationSkill({ ...options, vfs: new InMemoryVfs() })
      const deck = benchmarkDeck()
      expect(
        await original.executeTool({
          id: 'first',
          name: 'compile_deck_with_pptxgenjs',
          input: { request_id: 'request-1', deck },
        }),
      ).toMatchObject({ isError: true })
      expect(remembered).toBe(deck.id)
      service = createPresentationService({ userDataPath, compile })
      const vfs = new InMemoryVfs()
      const generation = createPresentationGenerationSkill({ ...options, vfs })
      const project = createPresentationProjectController({
        ...options,
        executeTool: generation.executeTool,
      })
      await project.refresh()
      expect(project.snapshot()).toMatchObject({
        phase: 'idle',
        project: { status: 'pending', latestRequestId: 'request-1', slideCount: 8 },
      })
      expect(project.snapshot().project?.checks).toBeUndefined()
      await project.resume('request-1')
      expect(project.snapshot()).toMatchObject({
        phase: 'idle',
        project: { status: 'compiled', checks: { render: 'not_run' } },
      })
      const opened = await openPptx(vfs.readBytes(`/home/user/generated/${deck.id}.pptx`))
      expect(opened.deck.slides).toHaveLength(8)
      expect(compile).toHaveBeenCalledTimes(2)
      await project.restore()
      expect(project.snapshot().error).toBeUndefined()
      expect(compile).toHaveBeenCalledTimes(2)
    } finally {
      rmSync(userDataPath, { recursive: true, force: true })
    }
  })
})

describe('planned production across restarts', () => {
  it('restores evidence and revision, compiles its native deck, and keeps old receipts after plan edits', async () => {
    const { benchmarkPlan, benchmarkPlannedDeck } =
      await import('../../../packages/pptx-engine/tests/fixtures/presentation-plan')
    const { createPresentationPlanningSkill } =
      await import('../../office-addin/src/skills/powerpoint/presentation-planning')
    const { createPresentationProjectController } =
      await import('../../office-addin/src/skills/powerpoint/presentation-project')
    const userDataPath = mkdtempSync(join(tmpdir(), 'wiswork-planned-production-'))
    try {
      const compile = vi.fn(compilePresentationDeck)
      let service = createPresentationService({ userDataPath, compile })
      let selected: string | undefined
      const options = {
        available: () => true,
        documentId: async () => 'doc-planned',
        lastProject: () => selected,
        rememberProject: async (id: string) => {
          selected = id
        },
        request: async (body: unknown, signal?: AbortSignal) =>
          new Response(Buffer.from(await service(body, signal ?? new AbortController().signal))),
      }
      const plan = benchmarkPlan()
      const first = createPresentationPlanningSkill({ ...options, vfs: new InMemoryVfs() })
      expect(
        await first.executeTool({
          id: 'save',
          name: 'save_presentation_plan',
          input: { expected_revision: 0, plan },
        }),
      ).toMatchObject({ mutated: false })
      expect(selected).toBe(plan.projectId)
      service = createPresentationService({ userDataPath, compile })
      const vfs = new InMemoryVfs()
      const planning = createPresentationPlanningSkill({ ...options, vfs })
      const restored = await planning.executeTool({
        id: 'read',
        name: 'read_presentation_plan',
        input: {},
      })
      expect(restored.isError).not.toBe(true)
      expect(JSON.parse(restored.output)).toMatchObject({ revision: 1, plan })
      const generation = createPresentationGenerationSkill({ ...options, vfs })
      const controller = createPresentationProjectController({
        ...options,
        executeTool: generation.executeTool,
      })
      await controller.refresh()
      expect(controller.snapshot().project?.status).toBe('planned')
      const call = {
        id: 'build',
        name: 'compile_deck_with_pptxgenjs',
        input: { request_id: 'build-1', plan_revision: 1, deck: benchmarkPlannedDeck() },
      }
      expect((await generation.executeTool(call)).isError).not.toBe(true)
      const before = vfs.readBytes(`/home/user/generated/${plan.projectId}.pptx`)
      expect((await openPptx(before)).deck.slides).toHaveLength(8)
      const next = { ...plan, brief: { ...plan.brief, objective: '修订汇报目标' } }
      expect(
        (
          await planning.executeTool({
            id: 'update',
            name: 'save_presentation_plan',
            input: { expected_revision: 1, plan: next },
          })
        ).isError,
      ).not.toBe(true)
      await controller.refresh()
      expect(controller.snapshot().project).toMatchObject({
        status: 'compiled',
        plan: { revision: 2 },
        requestPlanRevision: 1,
      })
      expect((await generation.executeTool(call)).isError).not.toBe(true)
      expect(compile).toHaveBeenCalledOnce()
      expect(vfs.readBytes(`/home/user/generated/${plan.projectId}.pptx`)).toEqual(before)
      expect(
        await generation.executeTool({ ...call, input: { ...call.input, request_id: 'build-2' } }),
      ).toMatchObject({ isError: true, output: 'presentation_revision_conflict' })
    } finally {
      rmSync(userDataPath, { recursive: true, force: true })
    }
  })
})
