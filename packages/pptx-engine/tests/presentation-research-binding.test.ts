import { describe, expect, it } from 'vitest'
import { parsePresentationPlan } from '../src/presentation-plan'
import {
  assertPresentationResearchBinding,
  presentationResearchBindingFindings,
} from '../src/presentation-research-binding'
import { researchFixture } from './fixtures/presentation-research'

describe('exact research plan binding', () => {
  it('accepts explicitly aliased IDs and preserves unavailable evidence findings', () => {
    const { plan, record } = researchFixture()
    expect(parsePresentationPlan(plan).research).toEqual(plan.research)
    expect(() => assertPresentationResearchBinding(plan, record)).not.toThrow()
    expect(presentationResearchBindingFindings(plan, record)).toEqual([
      {
        code: 'source_unavailable',
        claimId: 'source-1',
        researchClaimId: 'original-claim',
        sourceId: 'original-source',
      },
    ])
  })
  it.each(['statement', 'asOf', 'jurisdiction', 'type'])('rejects altered claim %s', (field) => {
    const { plan, record } = researchFixture()
    Object.assign(plan.claims[0]!, { [field]: field === 'type' ? 'judgment' : 'changed' })
    expect(() => assertPresentationResearchBinding(plan, record)).toThrow(
      'research_binding_invalid',
    )
  })
  it.each(['uri', 'snapshotAttachmentId', 'excerpt', 'locator', 'asOf'])(
    'rejects altered source %s',
    (field) => {
      const { plan, record } = researchFixture()
      Object.assign(plan.sources[0]!, {
        [field]: field === 'snapshotAttachmentId' ? 'b'.repeat(64) : 'changed',
      })
      expect(() => assertPresentationResearchBinding(plan, record)).toThrow()
    },
  )
  it('rejects nonexistent and duplicate plan mapping targets', () => {
    const { plan } = researchFixture()
    plan.research!.claims[0]!.claimId = 'missing'
    expect(() => parsePresentationPlan(plan)).toThrow()
    plan.research!.claims = [
      { claimId: 'source-1', researchClaimId: 'original-claim' },
      { claimId: 'source-1', researchClaimId: 'other' },
    ]
    expect(() => parsePresentationPlan(plan)).toThrow()
  })
  it('preserves unselected references, omitted conflict partners and unmapped plan claims', () => {
    const { plan, record } = researchFixture()
    for (let i = 2; i <= 4; i++) {
      const sourceId = `original-source-${i}`
      record.draft.sources.push({ ...record.draft.sources[0]!, id: sourceId })
      record.sources!.push({ sourceId, status: 'missing', provenance: 'unavailable' })
      record.draft.facts[0]!.sourceRefs.push(sourceId)
    }
    record.draft.facts.push({
      ...record.draft.facts[0]!,
      claimId: 'conflict',
      statement: 'Contrary position',
      conflictsWith: ['original-claim'],
    })
    plan.claims.push({ ...plan.claims[0]!, id: 'unmapped', statement: 'A separate plan judgment' })
    const findings = presentationResearchBindingFindings(plan, record)
    expect(
      findings
        .filter((finding) => finding.code === 'unselected_source_ref')
        .map((finding) => finding.sourceId),
    ).toEqual(['original-source-2', 'original-source-3', 'original-source-4'])
    expect(findings).toContainEqual({
      code: 'omitted_conflict_partner',
      claimId: 'source-1',
      researchClaimId: 'original-claim',
      relatedResearchClaimId: 'conflict',
    })
    expect(findings).toContainEqual({ code: 'unmapped_claim', claimId: 'unmapped' })
    expect(presentationResearchBindingFindings(plan, record)).toEqual(findings)
  })
  it('accepts later arithmetic reproduction and changed declared confidence but rejects calculation basis changes', () => {
    const { plan, record } = researchFixture()
    const claim = plan.claims[0]!,
      original = record.draft.facts[0]!
    claim.type = original.type = 'calculation'
    original.calculation = { formula: 'a + 2', inputs: ['input a'], unit: 'count', currency: 'CNY' }
    claim.calculation = {
      ...original.calculation,
      reproduction: {
        bindings: [{ name: 'a', inputIndex: 0, value: 1, sourceId: 'source' }],
        expected: 3,
      },
    }
    claim.confidence = 'high'
    expect(() => assertPresentationResearchBinding(plan, record)).not.toThrow()
    claim.calculation.unit = 'different'
    expect(() => assertPresentationResearchBinding(plan, record)).toThrow(
      'research_binding_invalid',
    )
  })
  it.each(['id', 'sequence', 'draftDigest', 'projectId'])(
    'rejects wrong record identity %s',
    (key) => {
      const { plan, record } = researchFixture()
      Object.assign(record, {
        [key]: key === 'sequence' ? 2 : key === 'draftDigest' ? 'b'.repeat(64) : 'other',
      })
      expect(() => assertPresentationResearchBinding(plan, record)).toThrow(
        'research_binding_invalid',
      )
    },
  )
  it('requires complete archived research and mapped source membership', () => {
    const { plan, record } = researchFixture()
    record.state = 'running'
    expect(() => assertPresentationResearchBinding(plan, record)).toThrow(
      'research_binding_invalid',
    )
    record.state = 'completed'
    record.draft.facts[0]!.sourceRefs = []
    expect(() => assertPresentationResearchBinding(plan, record)).toThrow(
      'research_binding_invalid',
    )
  })
})
