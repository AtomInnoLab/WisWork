import { describe, expect, it } from 'vitest'
import {
  compileSlidesAcceptance,
  verifySlidesAcceptance,
  type SlidesAcceptanceAuthority,
  type SlidesAcceptanceIntent,
} from '../src/renderer/ai/task-acceptance'

const revision = (digit: string) => `sha256:${digit.repeat(64)}`

function authority(): SlidesAcceptanceAuthority {
  const elements = (slide: number) => [
    {
      targetToken: `slide-${slide}:title`,
      role: 'title' as const,
      locked: false,
      properties: {
        text: `Title ${slide}`,
        color: slide === 6 ? '#111111' : '#222222',
        x: slide === 6 ? 100 : 140,
        y: slide === 6 ? 60 : 90,
        width: slide === 6 ? 600 : 560,
        height: 72,
        fill_color: '#FFFFFF',
        stroke_color: '#000000',
      },
    },
    {
      targetToken: `slide-${slide}:body`,
      role: 'body' as const,
      locked: false,
      properties: { color: '#333333' },
    },
    {
      targetToken: `slide-${slide}:emphasis`,
      role: 'emphasis' as const,
      locked: false,
      properties: { color: '#444444' },
    },
  ]
  return {
    documentToken: 'document-1',
    sessionToken: 'session-1',
    revision: revision('a'),
    slides: [6, 7, 8].map((number) => ({
      number,
      slideToken: `slide-${number}`,
      backgroundColor: '#FFFFFF',
      elements: elements(number),
    })),
  }
}

const goldenIntent: SlidesAcceptanceIntent = {
  taskId: 'golden-pages-6-8',
  affectedSlides: [6, 7, 8],
  changes: [
    { kind: 'set_property', slides: [6, 7, 8], role: 'title', property: 'color', value: '#2457A7' },
    { kind: 'set_property', slides: [6, 7, 8], role: 'body', property: 'color', value: '#172033' },
    {
      kind: 'set_property',
      slides: [6, 7, 8],
      role: 'emphasis',
      property: 'color',
      value: '#18A0A6',
    },
    {
      kind: 'match_reference',
      slides: [7, 8],
      referenceSlide: 6,
      role: 'title',
      properties: ['x', 'y', 'width', 'height'],
      tolerance: 0.5,
    },
  ],
  maxCorrectionPasses: 2,
}

describe('PC Slides acceptance compiler', () => {
  it('compiles the pages 6-8 golden request into deterministic durable checks', () => {
    const result = compileSlidesAcceptance(goldenIntent, authority())
    expect(result.status).toBe('compiled')
    if (result.status !== 'compiled') throw new Error('expected compiled')
    expect(result.contract).toMatchObject({
      taskId: 'golden-pages-6-8',
      documentToken: 'document-1',
      sessionToken: 'session-1',
      baseRevision: revision('a'),
      affectedSlides: [6, 7, 8],
      referenceSlides: [6],
    })
    expect(result.contract.checks).toHaveLength(17)
    expect(result.contract.checks).toContainEqual({
      id: 'check-001',
      kind: 'element_property',
      slide: 6,
      roleOrTarget: { kind: 'target', targetToken: 'slide-6:title' },
      property: 'color',
      expected: '#2457A7',
    })
    expect(result.contract.checks).toContainEqual({
      id: 'check-010',
      kind: 'element_property',
      slide: 7,
      roleOrTarget: { kind: 'target', targetToken: 'slide-7:title' },
      property: 'x',
      expected: 100,
    })
  })

  it('asks for clarification when a reference role is ambiguous', () => {
    const current = authority()
    current.slides[0]!.elements.push({
      targetToken: 'slide-6:title-2',
      role: 'title',
      locked: false,
      properties: { x: 1 },
    })
    expect(compileSlidesAcceptance(goldenIntent, current)).toEqual({
      status: 'needs_clarification',
      code: 'ambiguous_reference',
      slide: 6,
      role: 'title',
    })
  })

  it('requires the user for a locked target and reports a proved no-op', () => {
    const current = authority()
    current.slides[1]!.elements[0]!.locked = true
    expect(compileSlidesAcceptance(goldenIntent, current)).toMatchObject({
      status: 'needs_user',
      code: 'locked_target',
      slide: 7,
    })

    const noOp = compileSlidesAcceptance(
      {
        taskId: 'noop',
        affectedSlides: [6],
        changes: [
          { kind: 'set_property', slides: [6], role: 'body', property: 'color', value: '#333333' },
        ],
        maxCorrectionPasses: 0,
      },
      authority(),
    )
    expect(noOp).toEqual({ status: 'unchanged', taskId: 'noop', affectedSlides: [6] })
  })

  it('rejects unauthorized scope expansion', () => {
    expect(() =>
      compileSlidesAcceptance(
        {
          ...goldenIntent,
          affectedSlides: [6, 7],
        },
        authority(),
      ),
    ).toThrow(/outside frozen affected scope/)
  })
})

describe('PC Slides deterministic verifier', () => {
  it('passes supported text/color/geometry/fill/stroke/background checks without mutation', () => {
    const current = authority()
    const before = structuredClone(current)
    const intent: SlidesAcceptanceIntent = {
      taskId: 'all-supported',
      affectedSlides: [6],
      maxCorrectionPasses: 0,
      changes: [
        { kind: 'set_property', slides: [6], role: 'title', property: 'text', value: 'Title 6' },
        {
          kind: 'set_property',
          slides: [6],
          role: 'title',
          property: 'fill_color',
          value: '#FFFFFF',
        },
        {
          kind: 'set_property',
          slides: [6],
          role: 'title',
          property: 'stroke_color',
          value: '#000000',
        },
        {
          kind: 'set_property',
          slides: [6],
          role: 'title',
          property: 'x',
          value: 100,
          tolerance: 0.1,
        },
        { kind: 'set_background', slides: [6], color: '#FFFFFF' },
      ],
    }
    const compiled = compileSlidesAcceptance(intent, current)
    expect(compiled.status).toBe('unchanged')
    // Force verification coverage from a non-no-op contract by changing compile snapshot only.
    current.slides[0]!.elements[0]!.properties.text = 'before'
    current.slides[0]!.elements[0]!.properties.fill_color = '#101010'
    current.slides[0]!.elements[0]!.properties.stroke_color = '#101010'
    current.slides[0]!.elements[0]!.properties.x = 99
    current.slides[0]!.backgroundColor = '#101010'
    const planned = compileSlidesAcceptance(intent, current)
    if (planned.status !== 'compiled') throw new Error('expected compiled')
    const verificationAuthority = authority()
    const verificationBefore = structuredClone(verificationAuthority)
    const result = verifySlidesAcceptance(planned.contract, verificationAuthority)
    expect(result.every((entry) => entry.status === 'pass')).toBe(true)
    expect(verificationAuthority).toEqual(verificationBefore)
    expect(current).not.toEqual(before)
  })

  it('fails stale authority and scope expansion, and marks unsupported checks unavailable', () => {
    const compiled = compileSlidesAcceptance(goldenIntent, authority())
    if (compiled.status !== 'compiled') throw new Error('expected compiled')
    const stale = authority()
    stale.revision = revision('b')
    expect(
      verifySlidesAcceptance(compiled.contract, stale).every(
        (r) => r.status === 'fail' && r.code === 'stale_revision',
      ),
    ).toBe(true)

    const postWrite = authority()
    postWrite.baseRevision = revision('a')
    postWrite.revision = revision('c')
    postWrite.slides[0]!.elements[0]!.properties.color = '#2457A7'
    expect(verifySlidesAcceptance(compiled.contract, postWrite)[0]).toEqual({
      checkId: 'check-001',
      status: 'pass',
    })

    const expanded = authority()
    expanded.slides.push({
      number: 9,
      slideToken: 'slide-9',
      backgroundColor: '#FFFFFF',
      elements: [],
    })
    expanded.mutatedTargetTokens = ['slide-9:intruder']
    expect(verifySlidesAcceptance(compiled.contract, expanded)[0]).toMatchObject({
      status: 'fail',
      code: 'scope_mismatch',
    })

    const unsupported = structuredClone(compiled.contract)
    const firstCheck = unsupported.checks[0]!
    if (firstCheck.kind !== 'element_property') throw new Error('expected element check')
    unsupported.checks[0] = { ...firstCheck, property: 'font_family' }
    expect(verifySlidesAcceptance(unsupported, authority())[0]).toEqual({
      checkId: 'check-001',
      status: 'unavailable',
      code: 'unsupported_check',
    })
  })
})
