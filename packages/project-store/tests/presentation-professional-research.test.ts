import { expect, it, afterEach } from 'vitest'
import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import { researchDraft } from './fixtures/presentation-research'
import { parsePresentationResearchDraft } from '../src/presentation-research'
import { PresentationResearchStore } from '../src/presentation-research-store'
import { canonicalPresentationValue } from '../src/presentation-canonical'
const roots: string[] = []
afterEach(() => {
  for (const p of roots.splice(0)) rmSync(p, { recursive: true, force: true })
})
it('preserves full professional context, canonical digest and exact archived record after restart', async () => {
  const path = mkdtempSync(join(tmpdir(), 'research-professional-'))
  roots.push(path)
  const store = new PresentationResearchStore(path),
    draft = researchDraft()
  draft.facts[0]!.professionalContext = {
    domain: 'finance',
    materialKind: 'financial_statement',
    reportingPeriod: '2025全年',
    asOf: '2025-12-31',
    currency: 'CNY',
    unit: '百万元',
    accountingBasis: '声明口径',
    formula: '收入/成本',
    limitations: '未审计示例',
  }
  draft.facts[0]!.asOf = '2026-01-01'
  const begun = await store.begin('doc', 'project', 0, 'A', draft)
  expect(begun.record.draftDigest).toBe(
    createHash('sha256').update(canonicalPresentationValue(draft)).digest('hex'),
  )
  const completed = await store.finish('doc', 'project', 'A', {
    state: 'completed',
    sources: [{ sourceId: 'original', status: 'missing', provenance: 'unavailable' }],
  })
  const restored = await new PresentationResearchStore(path).read('doc', 'project', 'A')
  expect(restored).toEqual(completed)
  expect(restored!.draft).toEqual(draft)
  expect(() =>
    store.begin('doc', 'project', 0, 'A', {
      ...draft,
      facts: draft.facts.map((f) => ({ ...f, professionalContext: undefined })),
    }),
  ).toThrow('invalid_state')
})
it('keeps absent professional fields and old canonical SHA shape unchanged', async () => {
  const draft = researchDraft(),
    parsed = parsePresentationResearchDraft(draft)
  expect(parsed).toEqual(draft)
  expect(parsed.facts[0]).not.toHaveProperty('professionalContext')
  const path = mkdtempSync(join(tmpdir(), 'research-old-'))
  roots.push(path)
  const result = await new PresentationResearchStore(path).begin('doc', 'project', 0, 'old', draft)
  expect(result.record.draftDigest).toBe(
    createHash('sha256').update(canonicalPresentationValue(draft)).digest('hex'),
  )
})
it('rejects cross-domain or own undefined context before persisting', async () => {
  const path = mkdtempSync(join(tmpdir(), 'research-invalid-'))
  roots.push(path)
  const store = new PresentationResearchStore(path)
  for (const professionalContext of [undefined, { domain: 'law', sample: '跨领域' }]) {
    const draft = researchDraft()
    Object.assign(draft.facts[0]!, { professionalContext })
    expect(() => store.begin('doc', 'project', 0, 'bad', draft)).toThrow('invalid_state')
    expect((await store.summary('doc', 'project')).revision).toBe(0)
  }
})
