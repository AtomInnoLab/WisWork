import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { PresentationStore } from '../../../packages/project-store/src/presentation-store.js'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import { presentationWorkflowSummary } from '../src/agent/presentation-workflow.js'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'

describe('persisted plan snapshots through the Office workbench', () => {
  it('rejects an unchanged stored revision paired with a different valid plan', async () => {
    const root = mkdtempSync(join(tmpdir(), 'wiswork-plan-snapshot-'))
    try {
      const plan = benchmarkPlan()
      const store = new PresentationStore(root)
      const saved = store.savePlan(plan.projectId, 'document', 0, plan)
      const altered = structuredClone(plan)
      altered.sources[0]!.excerpt = '与已保存修订不同的资料内容'
      const controller = createPresentationProjectController({
        available: () => true,
        lastProject: () => plan.projectId,
        documentId: async () => 'document',
        executeTool: async () => ({ output: '{}', summary: 'read' }),
        request: async () =>
          new Response(
            JSON.stringify({
              projectId: plan.projectId,
              title: plan.title,
              status: 'planned',
              slideCount: plan.slides.length,
              slides: plan.slides.map(({ id, title }) => ({ id, title })),
              history: [],
              plan: { revision: saved.revision, value: altered, revisions: saved.revisions },
            }),
          ),
      })
      await controller.refresh()
      expect(controller.snapshot().error).toBeTruthy()
      expect(controller.snapshot().project).toBeUndefined()
      controller.clear()
    } finally {
      rmSync(root, { recursive: true, force: true })
    }
  })

  it('reopens actual stored revisions and preserves same-count change evidence without approval claims', async () => {
    const root = mkdtempSync(join(tmpdir(), 'wiswork-plan-snapshot-'))
    try {
      const plan = benchmarkPlan()
      const store = new PresentationStore(root)
      store.savePlan(plan.projectId, 'document', 0, plan)
      const revised = structuredClone(plan)
      revised.sources[0]!.excerpt = '第二版合成数据，仅用于测试'
      revised.claims[0]!.statement = '第二版待核验结论'
      revised.slides[0]!.purpose = '解释修订后的研究结果'
      store.savePlan(plan.projectId, 'document', 1, revised)
      const saved = new PresentationStore(root).plan(plan.projectId, 'document')!
      const value = {
        projectId: plan.projectId,
        title: plan.title,
        status: 'planned',
        slideCount: revised.slides.length,
        slides: revised.slides.map(({ id, title }) => ({ id, title })),
        history: [],
        plan: { revision: saved.revision, value: saved.plan, revisions: saved.revisions },
      }
      const open = () =>
        createPresentationProjectController({
          available: () => true,
          lastProject: () => plan.projectId,
          documentId: async () => 'document',
          executeTool: async () => ({ output: '{}', summary: 'read' }),
          request: async () => new Response(JSON.stringify(value)),
        })
      const first = open()
      await first.refresh()
      expect(first.snapshot().error).toBeUndefined()
      const summary = presentationWorkflowSummary(first.snapshot().project, undefined, undefined)!
      expect(summary.timeline.find((event) => event.id === 'plan-2')?.text).toContain(
        '已登记资料、主张、逐页计划有变化',
      )
      expect(summary.timeline.find((event) => event.id === 'research-2')?.text).toContain(
        '尚需核对来源与结论',
      )
      first.clear()
      const reopened = open()
      await reopened.refresh()
      expect(
        presentationWorkflowSummary(reopened.snapshot().project, undefined, undefined)?.timeline,
      ).toEqual(summary.timeline)
      expect(summary.timeline.filter((event) => event.id === 'plan-2')).toHaveLength(1)
      reopened.clear()
    } finally {
      rmSync(root, { recursive: true, force: true })
    }
  })
})
