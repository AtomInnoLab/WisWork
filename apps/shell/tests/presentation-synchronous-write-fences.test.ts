import { afterEach, expect, it, vi } from 'vitest'
import { mkdtempSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { benchmarkPlan } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { PresentationCommentLibrary } from '../src/main/presentation-comments'
import { PresentationManualObservationLibrary } from '../src/main/presentation-manual-observations'
import { PresentationPreferenceLibrary } from '../src/main/presentation-preferences'
const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'sync-project-fences-'))
  roots.push(root)
  return {
    root,
    comments: new PresentationCommentLibrary(root),
    observations: new PresentationManualObservationLibrary(root),
    preferences: new PresentationPreferenceLibrary(root),
  }
}
const shape = {
  id: 'shape',
  name: 'Title',
  type: 'TextBox',
  left: 1,
  top: 2,
  width: 300,
  height: 40,
  text: 'before',
}
const deny = (_scope?: Readonly<{ documentId: string; projectId: string }>) => {
  throw Error('project_deleting')
}
const digest = (text: string) => createHash('sha256').update(text).digest('hex')
const files = (root: string) =>
  readdirSync(root, { recursive: true, withFileTypes: true })
    .filter((v) => v.isFile())
    .map((v) => [
      join(v.parentPath, v.name),
      readFileSync(join(v.parentPath, v.name)).toString('base64'),
    ])
const approval = '12345678-1234-4123-8123-123456789abc'
it('refuses comment creation and resolution without changing body or creating a namespace', () => {
  const f = fixture(),
    plan = benchmarkPlan(),
    comment = {
      id: 'c',
      targetKind: 'slide',
      targetId: plan.slides[0]!.id,
      authorLabel: 'Reviewer',
      text: 'Check',
    }
  const guard = vi.fn(deny)
  expect(() =>
    f.comments.add('doc', plan.projectId, 0, 1, comment, { revision: 1, plan }, guard),
  ).toThrow('project_deleting')
  expect(guard).toHaveBeenCalledWith({ documentId: 'doc', projectId: plan.projectId })
  expect(Object.isFrozen(guard.mock.calls[0]![0])).toBe(true)
  expect(readdirSync(f.root)).toEqual([])
  f.comments.add('doc', plan.projectId, 0, 1, comment, { revision: 1, plan })
  const before = files(f.root)
  expect(() => f.comments.resolve('doc', plan.projectId, 1, 'c', deny)).toThrow('project_deleting')
  expect(files(f.root)).toEqual(before)
})
it('refuses observation begin, completion and deletion while retaining original snapshots', () => {
  const f = fixture()
  expect(() => f.observations.begin('doc', 'p', 'o', 'slide', shape, deny)).toThrow(
    'project_deleting',
  )
  expect(readdirSync(f.root)).toEqual([])
  const before = f.observations.begin('doc', 'p', 'o', 'slide', shape),
    disk = files(f.root)
  expect(() =>
    f.observations.complete(
      'doc',
      'p',
      'o',
      before.before.digest,
      { ...shape, text: 'after' },
      deny,
    ),
  ).toThrow('project_deleting')
  expect(() => f.observations.delete('doc', 'p', 'o', deny)).toThrow('project_deleting')
  expect(files(f.root)).toEqual(disk)
})
it('validates the nested preference target and fences save/delete without namespace creation', () => {
  const f = fixture(),
    input = { projectId: 'nested-project', changeId: 'change', text: 'Brief titles' },
    guard = vi.fn(deny)
  expect(() => f.preferences.save('doc', input, guard)).toThrow('project_deleting')
  expect(guard).toHaveBeenCalledWith({ documentId: 'doc', projectId: 'nested-project' })
  expect(readdirSync(f.root)).toEqual([])
  f.preferences.save('doc', input)
  const before = files(f.root)
  expect(() => f.preferences.delete('doc', input.projectId, input.changeId, deny)).toThrow(
    'project_deleting',
  )
  expect(files(f.root)).toEqual(before)
})
it('uses actual observation scope for preference save and never creates a preference directory after denial', () => {
  const f = fixture(),
    before = f.observations.begin('doc', 'observation-project', 'o', 'slide', shape)
  const observation = f.observations.complete(
      'doc',
      'observation-project',
      'o',
      before.before.digest,
      { ...shape, text: 'after' },
    ),
    disk = files(f.root),
    guard = vi.fn(deny)
  expect(() => f.preferences.saveObservation('doc', observation, 'Brief titles', guard)).toThrow(
    'project_deleting',
  )
  expect(guard).toHaveBeenCalledWith({ documentId: 'doc', projectId: 'observation-project' })
  expect(files(f.root)).toEqual(disk)
})
it.each(['source', 'target'])(
  'keeps import source and target guards distinct and writes neither on %s denial',
  (which) => {
    const f = fixture(),
      source = { documentId: 'source-doc', projectId: 'source-project', changeId: 'original' },
      input = { projectId: source.projectId, changeId: source.changeId, text: 'Brief titles' }
    f.preferences.save(source.documentId, input)
    const before = files(f.root),
      sourceGuard = vi.fn(which === 'source' ? deny : () => {}),
      targetGuard = vi.fn(which === 'target' ? deny : () => {})
    expect(() =>
      f.preferences.import(
        'target-doc',
        'target-project',
        source,
        digest(input.text),
        approval,
        undefined,
        { assertWritable: targetGuard, assertSourceCurrent: sourceGuard },
      ),
    ).toThrow('project_deleting')
    expect(sourceGuard).toHaveBeenCalledWith({
      documentId: source.documentId,
      projectId: source.projectId,
    })
    if (which === 'target')
      expect(targetGuard).toHaveBeenCalledWith({
        documentId: 'target-doc',
        projectId: 'target-project',
      })
    expect(files(f.root)).toEqual(before)
  },
)
it('snapshots validated preference inputs before a guard can mutate caller aliases', () => {
  const f = fixture(),
    input = { projectId: 'p', changeId: 'change', text: 'original' }
  const saved = f.preferences.save('doc', input, () => {
    input.projectId = 'foreign'
    input.text = 'mutated'
  })
  expect(saved).toEqual({ projectId: 'p', changeId: 'change', text: 'original' })
  expect(f.preferences.get('doc', 'p', 'change')).toEqual(saved)
  expect(f.preferences.list('doc', 'foreign')).toEqual([])
})
it.each(['comments', 'observations', 'preferences'])(
  'checks %s again before publishing and cleans only its staged write on denial',
  (kind) => {
    const f = fixture(),
      plan = benchmarkPlan(),
      comment = {
        id: 'c',
        targetKind: 'slide',
        targetId: plan.slides[0]!.id,
        authorLabel: 'Reviewer',
        text: 'Check',
      }
    if (kind === 'comments')
      f.comments.add('doc', plan.projectId, 0, 1, comment, { revision: 1, plan })
    else if (kind === 'observations') f.observations.begin('doc', 'p', 'o', 'slide', shape)
    else f.preferences.save('doc', { projectId: 'p', changeId: 'a', text: 'original' })
    const before = files(f.root),
      guard = () => {
        if (files(f.root).some(([path]) => path!.endsWith('.tmp'))) deny()
      }
    expect(() => {
      if (kind === 'comments') f.comments.resolve('doc', plan.projectId, 1, 'c', guard)
      else if (kind === 'observations') f.observations.delete('doc', 'p', 'o', guard)
      else f.preferences.delete('doc', 'p', 'a', guard)
    }).toThrow('project_deleting')
    expect(files(f.root)).toEqual(before)
  },
)
it('keeps missing readonly namespaces absent and supports legacy independent document/project labels', () => {
  const f = fixture()
  expect(f.comments.list('doc', 'p').comments).toEqual([])
  expect(f.observations.list('doc', 'p')).toEqual([])
  expect(f.preferences.get('doc', 'p', 'missing')).toBeUndefined()
  expect(f.preferences.list('doc', 'p')).toEqual([])
  expect(readdirSync(f.root)).toEqual([])
  f.preferences.save('a', { projectId: 'same-label', changeId: 'c', text: 'A' })
  f.preferences.save('b', { projectId: 'same-label', changeId: 'c', text: 'B' })
  expect(f.preferences.get('a', 'same-label', 'c')?.text).toBe('A')
  expect(f.preferences.get('b', 'same-label', 'c')?.text).toBe('B')
})
