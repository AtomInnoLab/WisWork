import { afterEach, expect, it } from 'vitest'
import { mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync, symlinkSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { createHash } from 'node:crypto'
import { PresentationStore } from '../src/presentation-store.js'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'presentation-project-scope-'))
  roots.push(root)
  return { root, store: new PresentationStore(root) }
}
it('reads only proven project metadata without creating content or loading receipt bodies', () => {
  const f = fixture()
  expect(f.store.projectScope('p', 'doc')).toBeUndefined()
  expect(readdirSync(f.root)).toEqual([])
  f.store.begin('p', 'doc', 'request', { secret: 'private document content' })
  const directory = join(
    f.root,
    'projects',
    'presentations',
    createHash('sha256').update('p').digest('hex'),
  )
  const receipt = readdirSync(directory).find((name) => name !== 'project.json')!
  writeFileSync(join(directory, receipt), 'unparseable private receipt')
  const metadata = readFileSync(join(directory, 'project.json'))
  expect(new PresentationStore(f.root).projectScope('p', 'doc')).toEqual({
    projectId: 'p',
    documentId: 'doc',
  })
  expect(readFileSync(join(directory, 'project.json'))).toEqual(metadata)
  expect(readFileSync(join(directory, receipt), 'utf8')).toBe('unparseable private receipt')
  expect(() => f.store.projectScope('p', 'foreign')).toThrow('document_mismatch')
  expect(() => f.store.projectScope('../p', 'doc')).toThrow('invalid_request')
})
it('rejects corrupt and symlinked binding metadata', () => {
  const f = fixture()
  f.store.begin('p', 'doc', 'request', {})
  const path = join(
    f.root,
    'projects',
    'presentations',
    createHash('sha256').update('p').digest('hex'),
    'project.json',
  )
  writeFileSync(path, JSON.stringify({ version: 1, projectId: 'foreign', documentId: 'doc' }))
  expect(() => f.store.projectScope('p', 'doc')).toThrow('invalid_state')
  rmSync(path)
  const foreign = join(f.root, 'foreign.json')
  writeFileSync(foreign, JSON.stringify({ version: 1, projectId: 'p', documentId: 'doc' }))
  symlinkSync(foreign, path)
  expect(() => f.store.projectScope('p', 'doc')).toThrow('invalid_state')
})
it('reads the 4096 character governance binding while ordinary writes retain their 2048 limit', () => {
  const f = fixture(),
    documentId = 'd'.repeat(4096)
  f.store.begin('p', 'doc', 'request', {})
  const path = join(
    f.root,
    'projects',
    'presentations',
    createHash('sha256').update('p').digest('hex'),
    'project.json',
  )
  writeFileSync(path, JSON.stringify({ version: 1, projectId: 'p', documentId }))
  const before = readFileSync(path)
  expect(f.store.projectScope('p', documentId)).toEqual({ projectId: 'p', documentId })
  expect(f.store.projectScope('absent', documentId)).toBeUndefined()
  expect(() => f.store.projectScope('p', documentId + 'd')).toThrow('invalid_request')
  expect(() => f.store.projectScope('p', 'x'.repeat(4096))).toThrow('document_mismatch')
  expect(() => f.store.begin('p', documentId, 'next', {})).toThrow('invalid_request')
  expect(readFileSync(path)).toEqual(before)
})
