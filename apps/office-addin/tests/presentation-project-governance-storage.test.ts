import { expect, it } from 'vitest'
import { createPresentationGovernanceStorage } from '../src/agent/presentation-project-governance-storage'
it('keeps only original bounded scoped intent and distinguishes projects', () => {
  const values = new Map<string, string>()
  const storage = createPresentationGovernanceStorage({
    getItem: (k) => values.get(k) ?? null,
    setItem: (k, v) => {
      values.set(k, v)
    },
    removeItem: (k) => {
      values.delete(k)
    },
  })
  const scope = { documentId: '文'.repeat(4096), projectId: 'p' },
    attempt = {
      version: 1 as const,
      scope,
      expectedRevision: null,
      confirmationToken: 'a'.repeat(64),
      deletionId: 'deletion',
    }
  storage.write(scope, attempt)
  attempt.deletionId = 'changed'
  expect(storage.read(scope)?.deletionId).toBe('deletion')
  expect(storage.read({ ...scope, projectId: 'other' })).toBeUndefined()
  expect(() =>
    storage.write(scope, { ...attempt, scope: { ...scope, projectId: 'wrong' } }),
  ).toThrow()
  storage.write(scope, undefined)
  expect(storage.read(scope)).toBeUndefined()
})
