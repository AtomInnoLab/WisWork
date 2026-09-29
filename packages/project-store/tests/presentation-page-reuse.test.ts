import { mkdtempSync, rmSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { afterEach, describe, expect, it } from 'vitest'
import { PresentationStore } from '../src/presentation-store.js'
import { presentationPageInput } from '../src/presentation-page-input.js'

const roots: string[] = []
afterEach(() => {
  for (const root of roots.splice(0)) rmSync(root, { recursive: true, force: true })
})
function input() {
  const deck = {
    version: 1,
    id: 'project',
    title: '计划',
    style: { fontFace: 'Arial' },
    assets: [] as { id: string; base64: string }[],
    claims: ['a', 'b', 'c'].map((id) => ({ id, text: id })),
    slides: ['a', 'b', 'c'].map((id) => ({
      id,
      title: id,
      claimIds: [id],
      elements: [] as { kind: string; assetId: string }[],
    })),
  }
  const plan = {
    projectId: 'project',
    style: deck.style,
    brief: { objective: '测试' },
    claims: ['a', 'b', 'c'].map((id) => ({ id, statement: id, sourceIds: [id] })),
    sources: ['a', 'b', 'c'].map((id) => ({ id, excerpt: id })),
    slides: ['a', 'b', 'c'].map((id) => ({
      id,
      title: id,
      claimIds: [id],
      dependsOn: [] as string[],
    })),
  }
  return { deck, binding: { revision: 1, plan } }
}
function setup() {
  const root = mkdtempSync(join(tmpdir(), 'ppt-plan-reuse-'))
  roots.push(root)
  const store = new PresentationStore(root)
  const original = input()
  let record = store.beginProduction('project', 'doc', 'parent', original.deck, original.binding)
  for (const page of record.pages) {
    record = store.updateProductionPage(record, page.pageId, { state: 'building', attempt: 1 })
    record = store.updateProductionPage(record, page.pageId, {
      state: 'compiled',
      attempt: 1,
      result: { pptxBase64: 'UEsDBAAAAAA=', sourceSlideId: '256#', report: { page: page.pageId } },
    })
  }
  return { root, store, original, parent: record }
}
describe('compiled page reuse after plan revisions', () => {
  it('retains compiled pages when locking and prevents changed production, compile, or rebuild inputs', () => {
    const f = setup()
    f.store.savePlan('project', 'doc', 0, f.original.binding.plan)
    const locked = f.store.setPlanPageLock('project', 'doc', 1, 'b', true)
    expect(locked.pageLocks?.[0]?.productionInputDigest).toMatch(/^[a-f0-9]{64}$/)
    const binding = { revision: locked.revision, plan: locked.plan }
    const reused = f.store.beginProduction(
      'project',
      'doc',
      'locked',
      f.original.deck,
      binding,
      true,
    )
    expect(reused.pages.every((page) => page.state === 'compiled')).toBe(true)
    const changed = structuredClone(f.original.deck)
    changed.slides[1]!.title += ' changed'
    expect(() =>
      f.store.beginProduction('project', 'doc', 'changed', changed, binding, true),
    ).toThrow('page_locked')
    expect(() => f.store.begin('project', 'doc', 'compile-changed', changed, binding)).toThrow(
      'page_locked',
    )
    expect(() => f.store.begin('project', 'doc', 'unplanned', changed)).toThrow('page_locked')
    expect(() =>
      f.store.deriveProduction('project', 'doc', 'parent', 'rebuild-locked', 'b', f.original.deck),
    ).toThrow('page_locked')
    expect(f.store.production('project', 'doc', 'changed')).toBeUndefined()
  })
  it('pins the first reserved production input when a page is locked before production', () => {
    const f = setup()
    const plan = structuredClone(f.original.binding.plan)
    plan.slides[1]!.title = 'new plan'
    f.store.savePlan('project', 'doc', 0, plan)
    const locked = f.store.setPlanPageLock('project', 'doc', 1, 'b', true)
    expect(locked.pageLocks?.[0]?.productionInputDigest).toBeUndefined()
    const binding = { revision: 2, plan: locked.plan }
    const deck = structuredClone(f.original.deck)
    deck.slides[1]!.title = 'new plan'
    f.store.beginProduction('project', 'doc', 'first-reservation', deck, binding)
    const reopened = new PresentationStore(f.root)
    expect(reopened.plan('project', 'doc')?.pageLocks?.[0]?.productionInputDigest).toMatch(
      /^[a-f0-9]{64}$/,
    )
    deck.slides[1]!.elements.push({ kind: 'image', assetId: 'changed' })
    expect(() =>
      reopened.beginProduction('project', 'doc', 'changed-reservation', deck, binding),
    ).toThrow('page_locked')
  })
  it('prevents rebuilding an unlocked predecessor that would invalidate a locked dependent page', () => {
    const f = setup()
    const dependent = structuredClone(f.original)
    dependent.binding.revision = 2
    dependent.binding.plan.slides[1]!.dependsOn = ['a']
    let record = f.store.beginProduction(
      'project',
      'doc',
      'dependent',
      dependent.deck,
      dependent.binding,
    )
    for (const page of record.pages) {
      record = f.store.updateProductionPage(record, page.pageId, { state: 'building', attempt: 1 })
      record = f.store.updateProductionPage(record, page.pageId, {
        state: 'compiled',
        attempt: 1,
        result: f.parent.pages.find((original) => original.pageId === page.pageId)!.result!,
      })
    }
    f.store.savePlan('project', 'doc', 0, dependent.binding.plan)
    f.store.setPlanPageLock('project', 'doc', 1, 'b', true)
    const changed = structuredClone(dependent.deck)
    changed.slides[0]!.title += ' changed'
    expect(() =>
      f.store.deriveProduction('project', 'doc', 'dependent', 'indirect', 'a', changed),
    ).toThrow('page_locked')
  })
  it('keeps compiled inputs unchanged when only the production concurrency changes', () => {
    const original = input()
    const changed = structuredClone(original)
    Object.assign(changed.binding.plan, { parallelism: 2 })
    expect(presentationPageInput(changed.deck, changed.binding, 'b')).toBe(
      presentationPageInput(original.deck, original.binding, 'b'),
    )
  })
  it('invalidates transitive dependents when rebuilding a page inside the same plan', () => {
    const f = setup()
    const input = structuredClone(f.original)
    input.binding.revision = 2
    input.binding.plan.slides[1]!.dependsOn = ['a']
    input.binding.plan.slides[2]!.dependsOn = ['b']
    let parent = f.store.beginProduction('project', 'doc', 'dependent', input.deck, input.binding)
    for (const page of parent.pages) {
      parent = f.store.updateProductionPage(parent, page.pageId, { state: 'building', attempt: 1 })
      parent = f.store.updateProductionPage(parent, page.pageId, {
        state: 'compiled',
        attempt: 1,
        result: f.parent.pages.find((original) => original.pageId === page.pageId)!.result!,
      })
    }
    input.deck.slides[0]!.title = '修改依赖起点'
    const child = f.store.deriveProduction(
      'project',
      'doc',
      'dependent',
      'rebuilt',
      'a',
      input.deck,
    )
    expect(child.pages.map((page) => page.state)).toEqual(['pending', 'pending', 'pending'])
    expect(new PresentationStore(f.root).production('project', 'doc', 'rebuilt')).toEqual(child)
  })
  it('reuses unchanged pages with durable provenance and leaves parent immutable', () => {
    const f = setup()
    const changed = structuredClone(f.original)
    changed.binding.revision = 2
    changed.deck.slides[1]!.title = '新标题'
    changed.binding.plan.slides[1]!.title = '新标题'
    const child = f.store.beginProduction(
      'project',
      'doc',
      'child',
      changed.deck,
      changed.binding,
      true,
    )
    expect(child.pages.map((page) => page.state)).toEqual(['compiled', 'pending', 'compiled'])
    expect(child.pages[0]!.reusedFrom).toEqual({
      requestId: 'parent',
      inputDigest: f.parent.inputDigest,
      planDigest: f.parent.planDigest,
    })
    expect(new PresentationStore(f.root).production('project', 'doc', 'child')).toEqual(child)
    expect(f.store.production('project', 'doc', 'parent')).toEqual(f.parent)
    expect(
      f.store.beginProduction('project', 'doc', 'child', changed.deck, changed.binding, true),
    ).toEqual(child)
    expect(() => f.store.production('project', 'other', 'child')).toThrow('document_mismatch')
  })

  it('invalidates cited evidence, assets and transitive dependency pages', () => {
    const original = input()
    original.binding.plan.slides[1]!.dependsOn = ['a']
    original.binding.plan.slides[2]!.dependsOn = ['b']
    const changed = structuredClone(original)
    changed.binding.plan.sources[0]!.excerpt = '新证据'
    for (const id of ['a', 'b', 'c'])
      expect(presentationPageInput(changed.deck, changed.binding, id)).not.toBe(
        presentationPageInput(original.deck, original.binding, id),
      )
    const assets = input()
    assets.deck.assets = [{ id: 'image', base64: 'before' }]
    assets.deck.slides[1]!.elements = [{ kind: 'image', assetId: 'image' }]
    const updated = structuredClone(assets)
    updated.deck.assets[0]!.base64 = 'after'
    expect(presentationPageInput(updated.deck, updated.binding, 'a')).toBe(
      presentationPageInput(assets.deck, assets.binding, 'a'),
    )
    expect(presentationPageInput(updated.deck, updated.binding, 'b')).not.toBe(
      presentationPageInput(assets.deck, assets.binding, 'b'),
    )
  })

  it('keeps unaffected pages after deletion/reorder but invalidates shared style', () => {
    const f = setup()
    const changed = structuredClone(f.original)
    changed.binding.revision = 2
    changed.deck.slides = [changed.deck.slides[0]!, changed.deck.slides[2]!]
    changed.binding.plan.slides = [changed.binding.plan.slides[0]!, changed.binding.plan.slides[2]!]
    expect(
      f.store
        .beginProduction('project', 'doc', 'deleted', changed.deck, changed.binding, true)
        .pages.every((page) => page.state === 'compiled'),
    ).toBe(true)
    changed.binding.revision = 3
    changed.deck.style.fontFace = 'Another'
    changed.binding.plan.style.fontFace = 'Another'
    expect(
      f.store
        .beginProduction('project', 'doc', 'styled', changed.deck, changed.binding, true)
        .pages.every((page) => page.state === 'pending'),
    ).toBe(true)
  })

  it('finds completed pages behind a newer pending request and rejects cyclic/missing references', () => {
    const f = setup()
    const changed = structuredClone(f.original)
    changed.binding.revision = 2
    f.store.beginProduction('project', 'doc', 'pending', changed.deck, changed.binding)
    changed.binding.revision = 3
    expect(
      f.store
        .beginProduction('project', 'doc', 'next', changed.deck, changed.binding, true)
        .pages.every((page) => page.reusedFrom?.requestId === 'parent'),
    ).toBe(true)
    changed.binding.plan.slides[0]!.dependsOn = ['a']
    expect(presentationPageInput(changed.deck, changed.binding, 'a')).toBeUndefined()
    changed.binding.plan.slides[0]!.dependsOn = ['missing']
    expect(presentationPageInput(changed.deck, changed.binding, 'a')).toBeUndefined()
  })

  it('rejects tampered provenance and copied results after reopening', () => {
    const f = setup()
    const changed = structuredClone(f.original)
    changed.binding.revision = 2
    f.store.beginProduction('project', 'doc', 'child', changed.deck, changed.binding, true)
    const hash = (value: string) => createHash('sha256').update(value).digest('hex')
    const path = join(
      f.root,
      'projects',
      'presentations',
      hash('project'),
      `production-${hash('child')}.json`,
    )
    const saved = readFileSync(path, 'utf8')
    const forged = JSON.parse(saved)
    forged.pages[0].reusedFrom.inputDigest = 'f'.repeat(64)
    writeFileSync(path, JSON.stringify(forged))
    expect(() => new PresentationStore(f.root).production('project', 'doc', 'child')).toThrow(
      'invalid_state',
    )
    writeFileSync(path, saved)
    forged.pages[0].reusedFrom.inputDigest = f.parent.inputDigest
    forged.pages[0].attempt += 1
    writeFileSync(path, JSON.stringify(forged))
    expect(() => new PresentationStore(f.root).production('project', 'doc', 'child')).toThrow(
      'invalid_state',
    )
  })
})
