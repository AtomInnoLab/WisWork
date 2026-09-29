import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync, readdirSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { PresentationLifecycleStore } from '@wiswork/project-store'
import {
  capturePresentationProjectWriteLease,
  capturePresentationProjectReadLease,
} from '../src/main/presentation-project-write-lease'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'project-write-lease-'))
  roots.push(root)
  const store = new PresentationLifecycleStore(root),
    scope = { projectId: 'p', documentId: 'doc' }
  return { root, store, scope, readExistingProject: () => ({ ...scope }) }
}
it('initializes only a confirmed historical project and fixes its original revision', () => {
  const f = fixture(),
    lease = capturePresentationProjectWriteLease(f)
  expect(lease.revision).toBe(0)
  lease.assertWritable()
  expect(new PresentationLifecycleStore(f.root).read(f.scope)?.revision).toBe(0)
  f.store.setPolicy(f.scope, 0, { contentRetentionDays: 30, auditRetentionDays: null })
  expect(() => lease.assertWritable()).toThrow('revision_conflict')
  expect(lease.revision).toBe(0)
})
it('does not initialize a missing or foreign project', () => {
  for (const value of [undefined, { projectId: 'p', documentId: 'foreign' }]) {
    const f = fixture()
    expect(() =>
      capturePresentationProjectWriteLease({ ...f, readExistingProject: () => value }),
    ).toThrow('project_not_found')
    expect(f.store.read(f.scope)).toBeUndefined()
  }
})
it('rejects deleting and reopened deleted tombstones without reviving them', () => {
  const f = fixture()
  f.store.initialize(f.scope)
  f.store.beginDeletion(f.scope, 0, {
    deletionId: 'delete',
    reason: 'user',
    resources: [{ resourceId: 'own', kind: 'project', ownership: 'project_exclusive' }],
  })
  expect(() => capturePresentationProjectWriteLease(f)).toThrow('project_deleting')
  f.store.recordDeletionResult(f.scope, 1, {
    deletionId: 'delete',
    resourceId: 'own',
    status: 'removed',
  })
  f.store.finishDeletion(f.scope, 2, 'delete')
  const reopened = new PresentationLifecycleStore(f.root)
  expect(() => capturePresentationProjectWriteLease({ ...f, store: reopened })).toThrow(
    'project_deleted',
  )
  expect(reopened.read(f.scope)?.state).toBe('deleted')
})
it('checks cancellation at admission and on every write assertion', () => {
  const f = fixture(),
    abort = new AbortController()
  const lease = capturePresentationProjectWriteLease({ ...f, signal: abort.signal })
  abort.abort()
  expect(() => lease.assertWritable()).toThrow('aborted')
  const empty = fixture()
  expect(() => capturePresentationProjectWriteLease({ ...empty, signal: abort.signal })).toThrow(
    'aborted',
  )
  expect(empty.store.read(empty.scope)).toBeUndefined()
})
it('owns caller scope/options aliases and cannot have its revision or scope changed', () => {
  const f = fixture(),
    original = { ...f.scope },
    lease = capturePresentationProjectWriteLease(f)
  f.scope.documentId = 'foreign'
  f.scope.projectId = 'other'
  f.store = new PresentationLifecycleStore(fixture().root)
  expect(lease.scope).toEqual(original)
  expect(Object.isFrozen(lease)).toBe(true)
  expect(Object.isFrozen(lease.scope)).toBe(true)
  lease.assertWritable()
})

it('does not initialize after project verification cancels admission', () => {
  const f = fixture(),
    abort = new AbortController()
  expect(() =>
    capturePresentationProjectWriteLease({
      ...f,
      signal: abort.signal,
      readExistingProject: () => {
        abort.abort()
        return { ...f.scope }
      },
    }),
  ).toThrow('aborted')
  expect(f.store.read(f.scope)).toBeUndefined()
})

it('captures an existing revision across reopen and rejects subsequent deletion', () => {
  const f = fixture()
  f.store.initialize(f.scope)
  f.store.setPolicy(f.scope, 0, { contentRetentionDays: 30, auditRetentionDays: null })
  const lease = capturePresentationProjectWriteLease({
    ...f,
    store: new PresentationLifecycleStore(f.root),
    readExistingProject: () => {
      throw Error('must not reinitialize')
    },
  })
  expect(lease.revision).toBe(1)
  lease.assertWritable()
  f.store.beginDeletion(f.scope, 1, {
    deletionId: 'delete',
    reason: 'user',
    resources: [{ resourceId: 'own', kind: 'project', ownership: 'project_exclusive' }],
  })
  expect(() => lease.assertWritable()).toThrow('revision_conflict')
  expect(f.store.read(f.scope)?.state).toBe('deleting')
})

it('reads an existing historical project without creating lifecycle files and fixes absence', () => {
  const f = fixture(),
    before = readdirSync(f.root),
    lease = capturePresentationProjectReadLease(f)
  expect(lease.revision).toBeUndefined()
  expect('assertWritable' in lease).toBe(false)
  lease.assertCurrent()
  expect(readdirSync(f.root)).toEqual(before)
  f.store.initialize(f.scope)
  expect(() => lease.assertCurrent()).toThrow('revision_conflict')
})
it('read leases preserve existing active revision and reject deleting or cancelled scopes', () => {
  const f = fixture(),
    controller = new AbortController()
  f.store.initialize(f.scope)
  const lease = capturePresentationProjectReadLease({ ...f, signal: controller.signal })
  expect(lease.revision).toBe(0)
  lease.assertCurrent()
  f.scope.projectId = 'foreign'
  expect(lease.scope.projectId).toBe('p')
  controller.abort()
  expect(() => lease.assertCurrent()).toThrow('aborted')
  const actual = { projectId: 'p', documentId: 'doc' }
  f.store.beginDeletion(actual, 0, {
    deletionId: 'delete',
    reason: 'user',
    resources: [{ resourceId: 'own', kind: 'project', ownership: 'project_exclusive' }],
  })
  expect(() => capturePresentationProjectReadLease({ ...f, scope: actual })).toThrow(
    'project_deleting',
  )
})
it('read leases do not establish ownership for missing projects', () => {
  const f = fixture()
  expect(() =>
    capturePresentationProjectReadLease({ ...f, readExistingProject: () => undefined }),
  ).toThrow('project_not_found')
  expect(readdirSync(f.root)).toEqual([])
})
