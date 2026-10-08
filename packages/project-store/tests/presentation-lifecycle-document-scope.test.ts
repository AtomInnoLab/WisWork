import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { PresentationLifecycleStore } from '../src/presentation-lifecycle'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
it('supports the existing research document identity budget without weakening scope binding', () => {
  const root = mkdtempSync(join(tmpdir(), 'lifecycle-research-scope-'))
  roots.push(root)
  const store = new PresentationLifecycleStore(root),
    scope = { projectId: 'research', documentId: 'd'.repeat(4096) }
  expect(store.initialize(scope).documentId).toBe(scope.documentId)
  expect(new PresentationLifecycleStore(root).read(scope)?.revision).toBe(0)
  expect(() => store.read({ ...scope, documentId: 'e'.repeat(4096) })).toThrow('document_mismatch')
  expect(() => store.initialize({ projectId: 'overflow', documentId: 'd'.repeat(4097) })).toThrow(
    'invalid_request',
  )
})
