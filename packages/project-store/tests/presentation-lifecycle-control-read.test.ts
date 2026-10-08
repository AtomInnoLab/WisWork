import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, mkdirSync, rmSync, writeFileSync, readdirSync, symlinkSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { createHash } from 'node:crypto'
import { PresentationLifecycleStore } from '../src/presentation-lifecycle.js'
const roots: string[] = []
afterEach(() => roots.splice(0).forEach((root) => rmSync(root, { recursive: true, force: true })))
it('observes its exact control without granting body ownership or initializing absent control', () => {
  const root = mkdtempSync(join(tmpdir(), 'control-read-'))
  roots.push(root)
  const scope = { projectId: 'p', documentId: 'doc' },
    store = new PresentationLifecycleStore(root)
  expect(store.readControl(scope)).toBeUndefined()
  expect(readdirSync(root)).toEqual([])
  store.initialize(scope)
  const dir = join(
    root,
    'projects',
    'presentations',
    createHash('sha256').update('p').digest('hex'),
  )
  mkdirSync(dir, { recursive: true })
  writeFileSync(join(dir, 'plan.json'), 'private')
  expect(store.readControl(scope)).toMatchObject({ revision: 0, state: 'active' })
  expect(() => store.read(scope)).toThrow('invalid_state')
  expect(() => store.assertActive(scope, 0)).toThrow('invalid_state')
  expect(() => store.initialize(scope)).toThrow('invalid_state')
  expect(() => store.readControl({ ...scope, documentId: 'foreign' })).toThrow('document_mismatch')
  const path = join(
    root,
    'presentation-project-lifecycles',
    createHash('sha256').update('p').digest('hex'),
    'lifecycle.json',
  )
  rmSync(path)
  const foreign = join(root, 'foreign')
  writeFileSync(foreign, 'private')
  symlinkSync(foreign, path)
  expect(() => store.readControl(scope)).toThrow('invalid_state')
})
