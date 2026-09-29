import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationLifecycleStore, PresentationStore } from '@wiswork/project-store'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { capturePresentationProjectWriteLease } from '../src/main/presentation-project-write-lease'
import { handlePresentationSourceAudit } from '../src/main/presentation-source-audit-history'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function deferred() {
  let resolve!: () => void
  const promise = new Promise<void>((r) => {
    resolve = r
  })
  return { promise, resolve }
}
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'source-audit-write-fences-'))
  roots.push(root)
  const store = new PresentationStore(root),
    lifecycle = new PresentationLifecycleStore(root),
    plan = benchmarkPlan()
  plan.sources[0]!.uri = 'attachment:' + 'a'.repeat(64)
  const scope = { projectId: plan.projectId, documentId: 'doc' }
  store.savePlan(scope.projectId, scope.documentId, 0, plan)
  const lease = capturePresentationProjectWriteLease({
    store: lifecycle,
    scope,
    readExistingProject: (s) => store.projectScope(s.projectId, s.documentId),
  })
  const request = { ...scope, operation: 'audit_sources', auditId: 'audit' }
  const freeze = () =>
    lifecycle.beginDeletion(scope, lease.revision, {
      deletionId: 'delete',
      reason: 'user',
      resources: [{ resourceId: 'project', kind: 'project', ownership: 'project_exclusive' }],
    })
  return { store, scope, request, lease, freeze }
}
it('rejects a frozen audit before creating its initial receipt', async () => {
  const f = fixture(),
    attachments = vi.fn(async () => ({ attachmentId: 'a'.repeat(64), status: 'found', offset: 0 }))
  f.freeze()
  await expect(
    handlePresentationSourceAudit(
      f.request,
      f.store,
      attachments,
      new AbortController().signal,
      f.lease.assertWritable,
    ),
  ).rejects.toThrow('revision_conflict')
  expect(f.store.sourceAudit(f.scope.projectId, 'doc', 'audit')).toBeUndefined()
  expect(attachments).not.toHaveBeenCalled()
})
it.each([false, true])(
  'does not finalize a source audit after freeze, including failure catch (%s)',
  async (failure) => {
    const f = fixture(),
      entered = deferred(),
      gate = deferred()
    const attachments = vi.fn(async (body: Record<string, unknown>) => {
      if (body.operation === 'attachment_metadata') {
        entered.resolve()
        await gate.promise
        if (failure) throw Error('offline')
        return { attachmentId: 'a'.repeat(64) }
      }
      return { attachmentId: 'a'.repeat(64), status: 'found', offset: 0 }
    })
    const result = handlePresentationSourceAudit(
      f.request,
      f.store,
      attachments,
      new AbortController().signal,
      f.lease.assertWritable,
    )
    await entered.promise
    const before = f.store.sourceAudit(f.scope.projectId, 'doc', 'audit')
    expect(before?.state).toBe('running')
    f.freeze()
    gate.resolve()
    await expect(result).rejects.toThrow('revision_conflict')
    expect(f.store.sourceAudit(f.scope.projectId, 'doc', 'audit')).toEqual(before)
  },
)

it('rechecks the final write fence after dynamic error metadata freezes the lifecycle', async () => {
  const f = fixture(),
    error = new Error()
  let reads = 0
  Object.defineProperty(error, 'message', {
    get() {
      reads++
      // The inner attachment audit reads once before rethrowing; outer classification reads next.
      if (reads === 2) f.freeze()
      return 'offline'
    },
  })
  await expect(
    handlePresentationSourceAudit(
      f.request,
      f.store,
      async () => {
        throw error
      },
      new AbortController().signal,
      f.lease.assertWritable,
    ),
  ).rejects.toThrow('revision_conflict')
  expect(reads).toBeGreaterThan(1)
  expect(f.store.sourceAudit(f.scope.projectId, 'doc', 'audit')?.state).toBe('running')
})
