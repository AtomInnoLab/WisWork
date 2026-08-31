import { describe, expect, it } from 'vitest'
import {
  compileSlidesAcceptance,
  parseSlidesAcceptanceAuthority,
  parseSlidesAcceptanceIntent,
  verifySlidesAcceptance,
  verifyAndBrandSlidesAcceptanceAuthority,
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
    expect(result.contract.checks).toHaveLength(11)
    expect(result.plannedMutationTargets).toHaveLength(9)
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
      kind: 'reference_match',
      slide: 7,
      referenceSlide: 6,
      role: 'title',
      properties: ['x', 'y', 'width', 'height'],
      tolerance: 0.5,
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

  it('rejects set_property tolerance', () => {
    expect(() =>
      compileSlidesAcceptance(
        {
          taskId: 'bad-tolerance',
          affectedSlides: [6],
          maxCorrectionPasses: 0,
          changes: [
            {
              kind: 'set_property',
              slides: [6],
              role: 'title',
              property: 'x',
              value: 100,
              tolerance: 0.1,
            },
          ],
        } as unknown as SlidesAcceptanceIntent,
        authority(),
      ),
    ).toThrow(/invalid/)
  })

  it('clarifies structurally ambiguous emphasis targets', () => {
    const current = authority()
    current.slides[0]!.elements.push({
      targetToken: 'slide-6:emphasis-2',
      role: 'emphasis',
      locked: false,
      properties: { color: '#444444' },
    })
    expect(
      compileSlidesAcceptance(
        {
          taskId: 'ambiguous-emphasis',
          affectedSlides: [6],
          maxCorrectionPasses: 0,
          changes: [
            {
              kind: 'set_property',
              slides: [6],
              role: 'emphasis',
              property: 'color',
              value: '#FFFFFF',
            },
          ],
        },
        current,
      ),
    ).toEqual({
      status: 'needs_clarification',
      code: 'ambiguous_target',
      slide: 6,
      role: 'emphasis',
    })
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
        },
        { kind: 'set_background', slides: [6], color: '#FFFFFF' },
      ],
    }
    const compiled = compileSlidesAcceptance(intent, current)
    expect(compiled.status).toBe('unchanged')
    // Force verification coverage from a non-no-op contract by changing compile snapshot only.
    current.slides[0]!.elements[0]!.properties.fill_color = '#101010'
    current.slides[0]!.elements[0]!.properties.stroke_color = '#101010'
    current.slides[0]!.elements[0]!.properties.x = 99
    current.slides[0]!.backgroundColor = '#101010'
    const planned = compileSlidesAcceptance(intent, current)
    if (planned.status !== 'compiled') throw new Error('expected compiled')
    const verificationAuthority = authority()
    const verificationBefore = structuredClone(verificationAuthority)
    const result = verifySlidesAcceptance(planned.contract, verificationAuthority, {
      mode: 'postwrite',
      mutatedTargetTokens: ['slide-6:title', 'slide-6'],
      plannedMutationTargets: planned.plannedMutationTargets,
    })
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
      verifySlidesAcceptance(compiled.contract, stale, { mode: 'prewrite' }).every(
        (r) => r.status === 'fail' && r.code === 'stale_revision',
      ),
    ).toBe(true)

    const postWrite = authority()
    postWrite.baseRevision = revision('a')
    postWrite.revision = revision('c')
    postWrite.slides[0]!.elements[0]!.properties.color = '#2457A7'
    expect(verifySlidesAcceptance(compiled.contract, postWrite, { mode: 'prewrite' })[0]).toEqual({
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
    expect(
      verifySlidesAcceptance(compiled.contract, expanded, {
        mode: 'postwrite',
        mutatedTargetTokens: ['slide-9:intruder'],
        plannedMutationTargets: compiled.plannedMutationTargets,
      })[0],
    ).toMatchObject({
      status: 'fail',
      code: 'scope_mismatch',
    })

    const unsupported = structuredClone(compiled.contract)
    const firstCheck = unsupported.checks[0]!
    if (firstCheck.kind !== 'element_property') throw new Error('expected element check')
    unsupported.checks[0] = { ...firstCheck, property: 'font_family' }
    expect(verifySlidesAcceptance(unsupported, authority(), { mode: 'prewrite' })[0]).toEqual({
      checkId: 'check-001',
      status: 'unavailable',
      code: 'unsupported_check',
    })
  })

  it('uses exact reference tolerance and never passes missing facts', () => {
    const current = authority()
    const compiled = compileSlidesAcceptance(
      {
        taskId: 'tolerance',
        affectedSlides: [7],
        maxCorrectionPasses: 0,
        changes: [
          {
            kind: 'match_reference',
            slides: [7],
            referenceSlide: 6,
            role: 'title',
            properties: ['x'],
            tolerance: 0.1,
          },
        ],
      },
      current,
    )
    if (compiled.status !== 'compiled') throw new Error('expected compiled')
    current.slides[1]!.elements[0]!.properties.x = 100.4
    expect(verifySlidesAcceptance(compiled.contract, current, { mode: 'prewrite' })[0]).toEqual({
      checkId: 'check-001',
      status: 'fail',
      code: 'value_mismatch',
    })
    delete current.slides[0]!.elements[0]!.properties.x
    delete current.slides[1]!.elements[0]!.properties.x
    expect(verifySlidesAcceptance(compiled.contract, current, { mode: 'prewrite' })[0]).toEqual({
      checkId: 'check-001',
      status: 'unavailable',
      code: 'unsupported_check',
    })
  })

  it('keeps equal missing facts verifiable and requires exact post-write target proof', () => {
    const current = authority()
    current.slides[0]!.elements[1]!.properties.bold = null
    const compiled = compileSlidesAcceptance(
      {
        taskId: 'missing',
        affectedSlides: [6],
        maxCorrectionPasses: 0,
        changes: [
          { kind: 'set_property', slides: [6], role: 'body', property: 'bold', value: false },
        ],
      },
      current,
    )
    expect(compiled.status).toBe('compiled')
    if (compiled.status !== 'compiled') throw new Error('expected compiled')
    expect(verifySlidesAcceptance(compiled.contract, current, { mode: 'prewrite' })[0].status).toBe(
      'unavailable',
    )
    for (const targets of [[], ['slide-6:body', 'extra']])
      expect(
        verifySlidesAcceptance(compiled.contract, current, {
          mode: 'postwrite',
          mutatedTargetTokens: targets,
          plannedMutationTargets: compiled.plannedMutationTargets,
        })[0],
      ).toMatchObject({ status: 'fail', code: 'scope_mismatch' })
  })

  it('does not require checked no-op targets to mutate and rejects extra writes', () => {
    const current = authority()
    const compiled = compileSlidesAcceptance(
      {
        taskId: 'mixed',
        affectedSlides: [6],
        maxCorrectionPasses: 0,
        changes: [
          { kind: 'set_property', slides: [6], role: 'title', property: 'x', value: 100 },
          { kind: 'set_property', slides: [6], role: 'body', property: 'color', value: '#ABCDEF' },
        ],
      },
      current,
    )
    if (compiled.status !== 'compiled') throw new Error('expected compiled')
    expect(compiled.plannedMutationTargets).toEqual(['slide-6:body'])
    current.slides[0]!.elements[1]!.properties.color = '#ABCDEF'
    expect(
      verifySlidesAcceptance(compiled.contract, current, {
        mode: 'postwrite',
        mutatedTargetTokens: ['slide-6:body'],
        plannedMutationTargets: compiled.plannedMutationTargets,
      }).every((result) => result.status === 'pass'),
    ).toBe(true)
    expect(
      verifySlidesAcceptance(compiled.contract, current, {
        mode: 'postwrite',
        mutatedTargetTokens: ['slide-6:body', 'slide-6:title'],
        plannedMutationTargets: compiled.plannedMutationTargets,
      })[0],
    ).toMatchObject({ status: 'fail', code: 'scope_mismatch' })
  })

  it('rejects forged text matches and accepts only verified branded proofs', async () => {
    const current = authority()
    delete current.slides[0]!.elements[0]!.properties.text
    current.leaseToken = 'lease:test'
    current.textMatches = {
      'check-001': { targetToken: 'slide-6:title', matches: true, proof: revision('d') },
    }
    const intent: SlidesAcceptanceIntent = {
      taskId: 'text-digest',
      affectedSlides: [6],
      maxCorrectionPasses: 0,
      changes: [
        { kind: 'set_property', slides: [6], role: 'title', property: 'text', value: 'Expected' },
      ],
    }
    expect(() => compileSlidesAcceptance(intent, current)).toThrow(/verified authority/)
    const verifiedNoop = await verifyAndBrandSlidesAcceptanceAuthority(
      intent,
      current,
      async (request) => request.proof === revision('d'),
    )
    expect(compileSlidesAcceptance(intent, verifiedNoop)).toEqual({
      status: 'unchanged',
      taskId: 'text-digest',
      affectedSlides: [6],
    })
    const intentB: SlidesAcceptanceIntent = {
      ...intent,
      changes: [
        { kind: 'set_property', slides: [6], role: 'title', property: 'text', value: 'Different' },
      ],
    }
    expect(() => compileSlidesAcceptance(intentB, verifiedNoop)).toThrow(/exact source/)
    expect(compileSlidesAcceptance(intent, verifiedNoop).status).toBe('unchanged')
    verifiedNoop.textMatches!['check-001']!.matches = false
    const verifiedMismatch = await verifyAndBrandSlidesAcceptanceAuthority(
      intent,
      verifiedNoop,
      async () => true,
    )
    const compiled = compileSlidesAcceptance(intent, verifiedMismatch)
    if (compiled.status !== 'compiled') throw new Error('expected compiled')
    expect(compiled.plannedMutationTargets).toEqual(['slide-6:title'])
    verifiedMismatch.textMatches!['check-001']!.matches = true
    const verifiedPostwrite = await verifyAndBrandSlidesAcceptanceAuthority(
      compiled.contract,
      verifiedMismatch,
      async () => true,
    )
    const proof = {
      mode: 'postwrite' as const,
      mutatedTargetTokens: ['slide-6:title'],
      plannedMutationTargets: compiled.plannedMutationTargets,
    }
    expect(verifySlidesAcceptance(compiled.contract, verifiedPostwrite, proof)[0].status).toBe(
      'pass',
    )
    const mutatedContract = structuredClone(compiled.contract)
    const textCheck = mutatedContract.checks[0]!
    if (textCheck.kind !== 'element_property') throw new Error('expected text check')
    textCheck.expected = 'Different'
    expect(() => verifySlidesAcceptance(mutatedContract, verifiedPostwrite, proof)).toThrow(
      /exact source/,
    )
    expect(verifySlidesAcceptance(compiled.contract, verifiedPostwrite, proof)[0].status).toBe(
      'pass',
    )
    ;(proof as Record<string, unknown>).expectedTextDigests = { 'check-001': revision('e') }
    verifiedPostwrite.textMatches!['check-001']!.matches = false
    expect(() => verifySlidesAcceptance(compiled.contract, verifiedPostwrite, proof)).toThrow(
      /verified authority/,
    )
    await expect(
      verifyAndBrandSlidesAcceptanceAuthority(
        compiled.contract,
        verifiedPostwrite,
        async () => false,
      ),
    ).rejects.toThrow(/proof/)
    const verifiedFailure = await verifyAndBrandSlidesAcceptanceAuthority(
      compiled.contract,
      verifiedPostwrite,
      async () => true,
    )
    expect(verifySlidesAcceptance(compiled.contract, verifiedFailure, proof)[0]).toMatchObject({
      status: 'fail',
      code: 'value_mismatch',
    })
  })
})

describe('PC Slides runtime input parsers', () => {
  it('rejects prototypes, accessors, symbols, unknown fields and bounds before iteration', () => {
    expect(parseSlidesAcceptanceIntent(goldenIntent)).toEqual(goldenIntent)
    expect(parseSlidesAcceptanceAuthority(authority())).toEqual(authority())
    expect(() =>
      parseSlidesAcceptanceIntent(Object.assign(Object.create({}), goldenIntent)),
    ).toThrow()
    const accessor = { ...goldenIntent }
    Object.defineProperty(accessor, 'changes', { enumerable: true, get: () => [] })
    expect(() => parseSlidesAcceptanceIntent(accessor)).toThrow()
    expect(() => parseSlidesAcceptanceIntent({ ...goldenIntent, unexpected: true })).toThrow()
    expect(() =>
      parseSlidesAcceptanceIntent({
        ...goldenIntent,
        changes: Array(51).fill(goldenIntent.changes[0]),
      }),
    ).toThrow()
    expect(() =>
      parseSlidesAcceptanceAuthority({ ...authority(), [Symbol('unsafe')]: true }),
    ).toThrow()
  })
})
