import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync, existsSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { createHash } from 'node:crypto'
import { PresentationStore, PresentationLifecycleStore } from '@wiswork/project-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { createOfficePresentationGovernanceProxy } from '../src/main/office-relay-client'
import { createPresentationProjectGovernanceController } from '../../office-addin/src/agent/presentation-project-governance'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'governance-office-integration-'))
  roots.push(root)
  const plan = benchmarkPlan(),
    documentId = 'private-office-document',
    scope = { documentId, projectId: plan.projectId }
  const store = new PresentationStore(root)
  store.savePlan(scope.projectId, documentId, 0, plan)
  const other = { ...plan, projectId: 'other-project' }
  store.savePlan(other.projectId, documentId, 0, other)
  const proxy = createOfficePresentationGovernanceProxy({ userDataPath: root })
  let saved: unknown,
    loseAcknowledgement = false
  const operations: string[] = []
  function controller(available = true) {
    return createPresentationProjectGovernanceController({
      available: () => available,
      documentId: async () => documentId,
      currentProjectId: () => scope.projectId,
      readAttempt: () => saved,
      writeAttempt: (_scope, value) => {
        saved = structuredClone(value)
      },
      request: async (body, signal) => {
        operations.push(String(body.operation))
        if (body.operation === 'project_deletion_confirm')
          expect(saved).toMatchObject({ scope, deletionId: body.deletionId })
        const result = await proxy(body, signal ?? new AbortController().signal)
        if (body.operation === 'project_deletion_confirm' && loseAcknowledgement)
          throw Error('synthetic_lost_ack')
        return new Response(new TextDecoder().decode(result))
      },
    })
  }
  return {
    root,
    scope,
    store,
    other,
    operations,
    controller,
    life: new PresentationLifecycleStore(root),
    loseAck: () => {
      loseAcknowledgement = true
    },
  }
}
it('actual Office controller uses the real PC factory for explicit policy, deletion, private control and anonymous audit', async () => {
  const f = fixture(),
    unavailable = f.controller(false)
  await unavailable.preview()
  expect(f.operations).toEqual([])
  const controller = f.controller()
  await controller.preview()
  expect(f.life.readControl(f.scope)).toBeUndefined()
  await controller.initializePolicy()
  expect(controller.snapshot().lifecycle?.revision).toBe(0)
  await controller.setPolicy({ contentRetentionDays: 30, auditRetentionDays: null })
  expect(controller.snapshot().lifecycle?.revision).toBe(1)
  await controller.preview()
  await controller.confirmDeletion()
  expect(controller.snapshot().phase).toBe('deleted')
  const path = join(
    f.root,
    'projects',
    'presentations',
    createHash('sha256').update(f.scope.projectId).digest('hex'),
  )
  expect(existsSync(path)).toBe(false)
  expect(f.store.plan(f.other.projectId, f.scope.documentId)?.plan).toEqual(f.other)
  expect(f.life.readControl(f.scope)?.state).toBe('deleted')
  await controller.exportAudit()
  const audit = JSON.stringify(controller.snapshot().audit)
  expect(audit).toContain('deletion_finished')
  expect(audit).not.toContain(f.scope.documentId)
  expect(audit).not.toContain(f.scope.projectId)
})
it('a lost real PC acknowledgement survives Office reopening and is resolved by control reads without replaying deletion', async () => {
  const f = fixture(),
    first = f.controller()
  await first.preview()
  f.loseAck()
  await first.confirmDeletion()
  expect(first.snapshot().phase).toBe('unknown')
  const deletionId = first.snapshot().attempt!.deletionId
  expect(f.life.readControl(f.scope)?.deletion?.deletionId).toBe(deletionId)
  expect(f.life.readControl(f.scope)?.state).toBe('deleted')
  const reopened = f.controller()
  await reopened.refresh()
  expect(reopened.snapshot().attempt?.deletionId).toBe(deletionId)
  await reopened.checkAttempt()
  expect(reopened.snapshot().phase).toBe('deleted')
  await reopened.confirmDeletion()
  await reopened.resumeDeletion()
  expect(f.operations.filter((op) => op === 'project_deletion_confirm')).toHaveLength(1)
  expect(f.operations.filter((op) => op === 'project_deletion_resume')).toHaveLength(0)
})
