import { describe, expect, it } from 'vitest'
import { parsePresentationDeck } from '../src/presentation'
import {
  parsePresentationPlan,
  presentationSourceAttachmentId,
  PRESENTATION_DOMAIN_PROFILES,
  presentationPlanClaims,
  assertDeckMatchesPresentationPlan,
  assertBrandKitRevision,
} from '../src/presentation-plan'
import { benchmarkPlan } from './fixtures/presentation-plan'
import { benchmarkDeck } from './fixtures/presentation-benchmark'

describe('durable presentation plan', () => {
  it('bounds opt-in parallelism and requires backward-only page dependencies', () => {
    const plan = benchmarkPlan()
    plan.parallelism = 2
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:page_dependency')
    plan.slides.forEach((slide) => {
      slide.dependsOn = []
    })
    plan.slides[1]!.dependsOn = [plan.slides[0]!.id]
    expect(() => parsePresentationPlan(plan)).not.toThrow()
    plan.slides[0]!.dependsOn = [plan.slides[1]!.id]
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:page_dependency')
    plan.slides[0]!.dependsOn = []
    plan.slides[1]!.dependsOn = [plan.slides[0]!.id, plan.slides[0]!.id]
    expect(() => parsePresentationPlan(plan)).toThrow(
      'presentation_plan_invalid:duplicate_page_dependency',
    )
    plan.slides[1]!.dependsOn = []
    plan.parallelism = 3 as 2
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:schema')
  })
  it('requires the selected domain workflow sections while leaving generic plans valid', () => {
    const plan = benchmarkPlan()
    expect(() => parsePresentationPlan(plan)).not.toThrow()
    plan.domain = 'research'
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:domain_section')
    const sections = ['question', 'method', 'results', 'limitations', 'references'] as const
    plan.slides = sections.map((section, index) => ({
      ...plan.slides[0]!,
      id: `domain-${index}`,
      domainSection: section,
    }))
    expect(() => parsePresentationPlan(plan)).not.toThrow()
    plan.slides[1]!.domainSection = 'customer_problem'
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:domain_section')
  })
  it('accepts each of the five domain workflows and rejects omitted sections', () => {
    for (const [domain, profile] of Object.entries(PRESENTATION_DOMAIN_PROFILES)) {
      const plan = benchmarkPlan()
      plan.domain = domain as keyof typeof PRESENTATION_DOMAIN_PROFILES
      plan.slides = profile.sections.map((domainSection, index) => ({
        ...plan.slides[0]!,
        id: `section-${index}`,
        domainSection,
      }))
      expect(() => parsePresentationPlan(plan)).not.toThrow()
      plan.slides.pop()
      expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:domain_section')
    }
  })
  it('pins reusable layout slots to native object types and geometry', () => {
    const plan = benchmarkPlan()
    plan.brandKit = {
      id: 'research-brand',
      revision: 1,
      name: 'Research',
      allowedColors: ['FFFFFF', '172033', '2255AA'],
      layoutComponents: [
        {
          id: 'title-body',
          name: 'Title and body',
          layout: 'content',
          slots: [
            { id: 'title', kind: 'text', x: 1, y: 1, w: 10, h: 1 },
            { id: 'body', kind: 'text', x: 1, y: 2.5, w: 10, h: 3 },
          ],
        },
      ],
    }
    plan.slides[0]!.layoutComponentId = 'title-body'
    const deck = benchmarkDeck()
    expect(() => assertDeckMatchesPresentationPlan(deck, parsePresentationPlan(plan))).not.toThrow()
    deck.slides[0]!.elements[1]!.x = 1.1
    expect(() => assertDeckMatchesPresentationPlan(deck, plan)).toThrow(
      'presentation_plan_mismatch:layout_component',
    )
    deck.slides[0]!.elements[1]!.x = 1
    deck.slides[0]!.elements[1] = {
      kind: 'shape',
      shape: 'rect',
      id: 'body',
      x: 1,
      y: 2.5,
      w: 10,
      h: 3,
    }
    expect(() => assertDeckMatchesPresentationPlan(deck, plan)).toThrow(
      'presentation_plan_mismatch:layout_component',
    )
    plan.slides[0]!.layout = 'cover'
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:layout_component')
    plan.slides[0]!.layout = 'content'
    plan.slides[0]!.layoutComponentId = 'missing'
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:layout_component')
    plan.slides[0]!.layoutComponentId = 'title-body'
    const revised = structuredClone(plan)
    revised.brandKit!.layoutComponents![0]!.slots[0]!.x = 2
    expect(() => assertBrandKitRevision(plan, revised)).toThrow(
      'presentation_plan_invalid:brand_kit_revision',
    )
    revised.brandKit!.revision = 2
    expect(() => assertBrandKitRevision(plan, revised)).not.toThrow()
    plan.brandKit.layoutComponents![0]!.slots[1]!.w = 13
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:brand_kit')
  })
  it('enforces an optional brand palette and required cover logo in the compiled deck', () => {
    const plan = benchmarkPlan()
    plan.brandKit = {
      id: 'research-brand',
      revision: 1,
      name: '研究品牌',
      allowedColors: ['FFFFFF', '172033', '2255AA'],
      logo: { assetId: 'pixel', assetDigest: 'a'.repeat(64), placement: 'cover' },
    }
    plan.slides[0]!.layout = 'cover'
    const deck = benchmarkDeck()
    deck.slides[0]!.elements.push({
      kind: 'image',
      id: 'brand-logo',
      x: 11,
      y: 0.3,
      w: 1,
      h: 1,
      assetId: 'pixel',
    })
    expect(parsePresentationPlan(plan).brandKit).toEqual(plan.brandKit)
    expect(() => assertDeckMatchesPresentationPlan(deck, plan)).not.toThrow()
    deck.slides[0]!.elements.pop()
    expect(() => assertDeckMatchesPresentationPlan(deck, plan)).toThrow(
      'presentation_plan_mismatch:brand_logo',
    )
    deck.slides[0]!.elements.push({
      kind: 'image',
      id: 'brand-logo',
      x: 11,
      y: 0.3,
      w: 1,
      h: 1,
      assetId: 'pixel',
    })
    deck.slides[0]!.elements[0] = {
      ...deck.slides[0]!.elements[0]!,
      kind: 'text',
      text: '标题',
      color: 'FF0000',
    }
    expect(() => assertDeckMatchesPresentationPlan(deck, plan)).toThrow(
      'presentation_plan_mismatch:brand_color',
    )
    deck.slides[0]!.elements[0] = {
      ...deck.slides[0]!.elements[0]!,
      kind: 'text',
      text: '标题',
      color: '172033',
    }
    plan.brandKit.allowedColors = ['FFFFFF', 'FFFFFF', '2255AA']
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:brand_kit')
    plan.brandKit.allowedColors = ['FFFFFF', '172033', '2255AA']
    plan.slides[0]!.layout = 'content'
    expect(() => parsePresentationPlan(plan)).toThrow(
      'presentation_plan_invalid:brand_kit_logo_scope',
    )
    plan.brandKit.logo!.placement = 'all'
    expect(() => parsePresentationPlan(plan)).not.toThrow()
    expect(() => assertDeckMatchesPresentationPlan(deck, plan)).toThrow(
      'presentation_plan_mismatch:brand_logo',
    )
  })
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

it('resolves source snapshots while preserving original URIs and legacy plans', () => {
  const plan = benchmarkPlan()
  const source = plan.sources[0]!
  expect(presentationSourceAttachmentId(source)).toBeUndefined()
  source.uri = 'https://example.com/original'
  source.snapshotAttachmentId = 'a'.repeat(64)
  expect(parsePresentationPlan(plan).sources[0]).toEqual(source)
  expect(presentationSourceAttachmentId(source)).toBe(source.snapshotAttachmentId)
  expect(presentationPlanClaims(plan)[0]!.source).toBe(source.uri)
  source.uri = `attachment:${source.snapshotAttachmentId}`
  expect(() => parsePresentationPlan(plan)).not.toThrow()
  delete source.snapshotAttachmentId
  expect(presentationSourceAttachmentId(source)).toBe('a'.repeat(64))
  source.snapshotAttachmentId = 'b'.repeat(64)
  expect(() => parsePresentationPlan(plan)).toThrow(
    'presentation_plan_invalid:source_snapshot_conflict',
  )
  for (const value of ['', 'A'.repeat(64), '../file', 'a'.repeat(63)]) {
    source.snapshotAttachmentId = value
    expect(() => parsePresentationPlan(plan)).toThrow('presentation_plan_invalid:')
  }
})
