import { createHash } from 'node:crypto'
import { expect, it, vi } from 'vitest'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan.js'
import {
  canonicalPresentationValue,
  presentationPlanSnapshotInputs,
} from '@wiswork/project-store/presentation-canonical'
import { createPresentationProjectController } from '../src/skills/powerpoint/presentation-project.js'
function fixture() {
  const plan = benchmarkPlan(),
    hash = (value: string) => createHash('sha256').update(value).digest('hex')
  let documentId = 'doc',
    release!: () => void
  const gate = new Promise<void>((resolve) => {
    release = resolve
  })
  const calls: Record<string, unknown>[] = []
  let forged = false
  const controller = createPresentationProjectController({
    available: () => true,
    documentId: async () => documentId,
    lastProject: () => plan.projectId,
    executeTool: async () => ({ output: '', summary: '' }),
    request: async (body) => {
      const request = body as Record<string, unknown>
      if (request.operation === 'accept_plan') {
        calls.push(request)
        await gate
        return new Response(
          JSON.stringify({
            projectId: plan.projectId,
            documentId: 'doc',
            acceptance: {
              decisionId: request.decisionId,
              planRevision: 1,
              planDigest: forged ? 'f'.repeat(64) : hash(canonicalPresentationValue(plan)),
              styleDigest: hash(presentationPlanSnapshotInputs(plan).styleDigest),
              acceptedAt: '2026-09-29T00:00:00.000Z',
            },
          }),
        )
      }
      return new Response(
        JSON.stringify({
          projectId: plan.projectId,
          title: plan.title,
          status: 'planned',
          slideCount: plan.slides.length,
          slides: plan.slides.map(({ id, title }) => ({ id, title })),
          history: [],
          plan: { revision: 1, value: plan },
          planAcceptance: { version: 1, projectId: plan.projectId, documentId: 'doc', records: [] },
        }),
      )
    },
  })
  return {
    controller,
    calls,
    release,
    changeDocument: () => {
      documentId = 'foreign'
    },
    forge: () => {
      forged = true
    },
  }
}
it.each(['cancel', 'document', 'forged'] as const)(
  'does not adopt a late or mismatched acceptance after %s',
  async (mode) => {
    const f = fixture()
    await f.controller.refresh()
    const pending = f.controller.acceptPlan!(1)
    await vi.waitFor(() => expect(f.calls).toHaveLength(1))
    if (mode === 'cancel') f.controller.cancel()
    else if (mode === 'document') f.changeDocument()
    else f.forge()
    f.release()
    await pending
    expect(f.controller.snapshot().project?.planAcceptanceCurrent).toBeUndefined()
    expect(f.controller.snapshot().project?.planAcceptance?.records ?? []).toEqual([])
    if (mode !== 'cancel') expect(f.controller.snapshot().error).toBeTruthy()
    expect(f.calls).toHaveLength(1)
  },
)
