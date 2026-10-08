import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createPresentationService } from '../src/main/presentation-service'
vi.mock('@wiswork/pptx-engine/presentation-compiler', () => ({ compilePresentationDeck: vi.fn() }))
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'presentation-service-lifecycle-'))
  roots.push(root)
  const compile = vi.fn(async () => {
    throw Error('unexpected_compile')
  })
  const service = createPresentationService({ userDataPath: root, compile: compile as never })
  const call = async (operation: string, extra: Record<string, unknown> = {}) =>
    JSON.parse(
      new TextDecoder().decode(
        await service(
          { operation, projectId: 'project', documentId: 'document', ...extra },
          new AbortController().signal,
        ),
      ),
    )
  return { root, service, compile, call }
}
it('routes actual PC policy metadata and anonymous audit without compiling or deleting content', async () => {
  const f = fixture()
  expect(await f.call('project_lifecycle_read')).toEqual({ lifecycle: null })
  const created = await f.call('project_lifecycle_initialize')
  expect(created.lifecycle).toMatchObject({ revision: 0, state: 'active' })
  const policy = { contentRetentionDays: 30, auditRetentionDays: 365 }
  expect(
    await f.call('project_lifecycle_set_policy', { expectedRevision: 0, policy }),
  ).toMatchObject({ lifecycle: { revision: 1, policy } })
  const reopened = createPresentationService({ userDataPath: f.root })
  const result = JSON.parse(
    new TextDecoder().decode(
      await reopened(
        { operation: 'project_lifecycle_read', projectId: 'project', documentId: 'document' },
        new AbortController().signal,
      ),
    ),
  )
  expect(result.lifecycle).toMatchObject({ revision: 1, policy })
  const audit = await f.call('project_lifecycle_export_audit')
  expect(audit.audit.events.map((v: any) => v.action)).toEqual(['created', 'policy_updated'])
  expect(audit.audit).not.toHaveProperty('documentId')
  expect(audit.audit).not.toHaveProperty('projectId')
  expect(JSON.stringify(audit)).not.toContain(f.root)
  expect(f.compile).not.toHaveBeenCalled()
})
it('preserves lifecycle on cancelled or stale actual PC requests', async () => {
  const f = fixture()
  await f.call('project_lifecycle_initialize')
  expect(
    await f.call('project_lifecycle_set_policy', {
      expectedRevision: 99,
      policy: { contentRetentionDays: 30, auditRetentionDays: null },
    }),
  ).toEqual({ error: 'revision_conflict' })
  const response = await f.service(
    { operation: 'project_lifecycle_initialize', projectId: 'other', documentId: 'document' },
    AbortSignal.abort(),
  )
  expect(JSON.parse(new TextDecoder().decode(response))).toEqual({ error: 'aborted' })
  expect(await f.call('project_lifecycle_read', { projectId: 'other' })).toEqual({
    lifecycle: null,
  })
})
it('normalizes hostile exception metadata at the actual PC entry', async () => {
  const f = fixture()
  const error = new Error()
  Object.defineProperty(error, 'message', {
    get: () => {
      throw Error('private-token=secret')
    },
  })
  const input = {
    toJSON: () => {
      throw error
    },
  }
  const response = await f.service(input, new AbortController().signal)
  expect(JSON.parse(new TextDecoder().decode(response))).toEqual({ error: 'compile_failed' })
})
