import { describe, expect, it } from 'vitest'
import { parsePresentationDeck } from '../src/presentation'
import {
  parsePresentationPlan,
  presentationPlanClaims,
  assertDeckMatchesPresentationPlan,
} from '../src/presentation-plan'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { benchmarkDeck } from './fixtures/presentation-benchmark'

describe('durable presentation plan', () => {
  it('clones a valid plan and maps the benchmark deck exactly', () => {
    const plan = benchmarkPlan()
    expect(parsePresentationPlan(plan)).toEqual(plan)
    expect(parsePresentationPlan(plan)).not.toBe(plan)
    expect(presentationPlanClaims(plan)).toEqual(benchmarkDeck().claims)
    expect(() => assertDeckMatchesPresentationPlan(benchmarkDeck(), plan)).not.toThrow()
    expect(parsePresentationDeck(benchmarkDeck()).id).toBe(plan.projectId)
  })
  it('rejects forged verification, missing sources and invalid calculations', () => {
    for (const change of [
      { reviewStatus: 'verified' },
      { type: 'fact', sourceIds: [] },
      { type: 'quote', sourceIds: [] },
      { type: 'calculation' },
      { type: 'calculation', calculation: { formula: '1+1', inputs: [] } },
    ]) {
      const plan = benchmarkPlan()
      Object.assign(plan.claims[0]!, change)
      expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:')
    }
    const plan = benchmarkPlan()
    Object.assign(plan.claims[0]!, {
      type: 'calculation',
      calculation: { formula: 'x * 2', inputs: ['x=1'], unit: 'count' },
    })
    expect(parsePresentationPlan(plan)).toEqual(plan)
  })
  it('rejects duplicate IDs/references and dangling references', () => {
    const changes = [
      (p: ReturnType<typeof benchmarkPlan>) => p.slides.push(p.slides[0]!),
      (p: ReturnType<typeof benchmarkPlan>) => p.sources.push(p.sources[0]!),
      (p: ReturnType<typeof benchmarkPlan>) => p.claims.push(p.claims[0]!),
      (p: ReturnType<typeof benchmarkPlan>) => p.claims[0]!.sourceIds.push('source'),
      (p: ReturnType<typeof benchmarkPlan>) => p.claims[0]!.sourceIds.push('missing'),
      (p: ReturnType<typeof benchmarkPlan>) => p.slides[0]!.claimIds.push('missing'),
      (p: ReturnType<typeof benchmarkPlan>) => p.slides[0]!.claimIds.push('source-1'),
    ]
    for (const change of changes) {
      const plan = benchmarkPlan()
      change(plan)
      expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:')
    }
  })
  it('bounds fields, slide count and serialized UTF-8 bytes', () => {
    for (const change of [
      (p: ReturnType<typeof benchmarkPlan>) => {
        p.projectId = 'a'.repeat(81)
      },
      (p: ReturnType<typeof benchmarkPlan>) => {
        p.sources[0]!.uri = 'x'.repeat(501)
      },
      (p: ReturnType<typeof benchmarkPlan>) => {
        p.sources[0]!.locator = 'x'.repeat(201)
      },
      (p: ReturnType<typeof benchmarkPlan>) => {
        p.slides = []
      },
      (p: ReturnType<typeof benchmarkPlan>) => {
        p.slides = Array.from({ length: 33 }, (_, i) => ({ ...p.slides[0]!, id: `s${i}` }))
      },
      (p: ReturnType<typeof benchmarkPlan>) => {
        p.sources = Array.from({ length: 30 }, (_, i) => ({
          ...p.sources[0]!,
          id: `source${i}`,
          excerpt: '汉'.repeat(3000),
        }))
        p.claims = []
        p.slides.forEach((s) => {
          s.claimIds = []
        })
      },
    ]) {
      const plan = benchmarkPlan()
      change(plan)
      expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:')
    }
    const plan = benchmarkPlan()
    plan.slides = Array.from({ length: 32 }, (_, i) => ({ ...plan.slides[0]!, id: `s${i}` }))
    expect(parsePresentationPlan(plan).slides).toHaveLength(32)
  })
  it('rejects expanded claims beyond the deck text budget even when the plan fits its byte budget', () => {
    const plan = benchmarkPlan()
    plan.sources = Array.from({ length: 3 }, (_, i) => ({
      id: `source-${i}`,
      title: 'Source',
      uri: 'u'.repeat(499),
      excerpt: '',
    }))
    plan.claims = Array.from({ length: 256 }, (_, i) => ({
      ...plan.claims[0]!,
      id: `claim-${i}`,
      statement: 'Claim',
      sourceIds: plan.sources.map((source) => source.id),
    }))
    plan.slides.forEach((slide) => {
      slide.claimIds = []
    })
    expect(new TextEncoder().encode(JSON.stringify(plan)).byteLength).toBeLessThan(192 * 1024)
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:text_budget')
  })
  it('rejects the character budget boundary when transport bytes already exceed their limit', () => {
    const plan = benchmarkPlan()
    plan.sources = Array.from({ length: 3 }, (_, i) => ({
      id: `source-${i}`,
      title: 'Source',
      uri: 'u'.repeat(400),
      locator: 'l'.repeat(10),
      excerpt: '',
    }))
    plan.claims = Array.from({ length: 200 }, (_, i) => ({
      ...plan.claims[0]!,
      id: `claim-${i}`,
      statement: 'C',
      sourceIds: plan.sources.map((source) => source.id),
    }))
    plan.slides.forEach((slide) => {
      slide.claimIds = []
      slide.title = 't'.repeat(175)
    })
    expect(() => parsePresentationPlan(plan)).toThrow(
      'presentation_plan_invalid:compiled_byte_budget',
    )
    plan.slides[0]!.title += 't'
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:text_budget')
  })
  it('rejects multibyte source expansion below the deck character limit', () => {
    const plan = benchmarkPlan()
    plan.sources = Array.from({ length: 3 }, (_, i) => ({
      id: `s${i}`,
      title: 'Source',
      uri: '汉'.repeat(499),
      excerpt: '',
    }))
    plan.claims = Array.from({ length: 100 }, (_, i) => ({
      ...plan.claims[0]!,
      id: `c${i}`,
      statement: 'C',
      sourceIds: plan.sources.map((source) => source.id),
    }))
    plan.slides.forEach((slide) => {
      slide.claimIds = []
    })
    expect(new TextEncoder().encode(JSON.stringify(plan)).byteLength).toBeLessThan(192 * 1024)
    expect(() => parsePresentationPlan(plan)).toThrow(
      'presentation_plan_invalid:compiled_byte_budget',
    )
  })
  it('accepts exactly 192 KiB of mandatory deck JSON and rejects one additional byte', () => {
    const plan = benchmarkPlan()
    plan.sources = Array.from({ length: 3 }, (_, i) => ({
      id: `s${i}`,
      title: 'Source',
      uri: 'u'.repeat(400),
      excerpt: '',
    }))
    plan.claims = Array.from({ length: 120 }, (_, i) => ({
      ...plan.claims[0]!,
      id: `c${i}`,
      statement: 'C',
      sourceIds: plan.sources.map((source) => source.id),
    }))
    plan.slides.forEach((slide) => {
      slide.claimIds = []
    })
    const mandatoryDeck = {
      version: 1,
      id: plan.projectId,
      title: plan.title,
      style: plan.style,
      assets: [],
      claims: presentationPlanClaims(plan),
      slides: plan.slides.map((slide) => ({
        id: slide.id,
        title: slide.title,
        claimIds: slide.claimIds,
        elements: [],
      })),
    }
    let remaining = 192 * 1024 - new TextEncoder().encode(JSON.stringify(mandatoryDeck)).byteLength
    expect(remaining).toBeGreaterThan(0)
    for (const claim of plan.claims) {
      const padding = Math.min(remaining, 11998)
      claim.statement += 'x'.repeat(padding)
      remaining -= padding
    }
    expect(remaining).toBe(0)
    expect(() => parsePresentationPlan(plan)).not.toThrow()
    plan.claims[0]!.statement += 'x'
    expect(() => parsePresentationPlan(plan)).toThrow(
      'presentation_plan_invalid:compiled_byte_budget',
    )
  })
  it('preserves unverified provenance and joins only nonempty locators', () => {
    const plan = benchmarkPlan()
    plan.sources.push({
      id: 'other',
      title: 'other',
      uri: 'https://example.test/private',
      excerpt: '',
    })
    plan.claims[0]!.sourceIds.push('other')
    expect(presentationPlanClaims(plan)[0]).toEqual({
      ...benchmarkDeck().claims[0],
      source: '研究报告（合成基准） ; https://example.test/private',
    })
    plan.claims[0]!.sourceIds = []
    expect(presentationPlanClaims(plan)[0]).toEqual({
      id: 'source-1',
      text: '示例数据仅用于测试',
      source: '未核验：assumption',
    })
  })
  it('rejects content/style drift while ignoring unclaimed visual QA', () => {
    const changes = [
      (d: ReturnType<typeof benchmarkDeck>) => {
        d.id = 'other'
      },
      (d: ReturnType<typeof benchmarkDeck>) => {
        d.title = 'other'
      },
      (d: ReturnType<typeof benchmarkDeck>) => {
        d.style.accentColor = '000000'
      },
      (d: ReturnType<typeof benchmarkDeck>) => {
        d.slides.reverse()
      },
      (d: ReturnType<typeof benchmarkDeck>) => {
        d.slides[0]!.title = 'other'
      },
      (d: ReturnType<typeof benchmarkDeck>) => {
        d.slides[0]!.claimIds = []
      },
      (d: ReturnType<typeof benchmarkDeck>) => {
        d.claims[0]!.text = 'other'
      },
      (d: ReturnType<typeof benchmarkDeck>) => {
        d.claims[0]!.source = 'other'
      },
      (d: ReturnType<typeof benchmarkDeck>) => {
        delete d.claims[0]!.locator
      },
    ]
    for (const change of changes) {
      const deck = benchmarkDeck()
      change(deck)
      expect(() => assertDeckMatchesPresentationPlan(deck, benchmarkPlan())).toThrow(
        'presentation_plan_mismatch:',
      )
    }
  })
})
