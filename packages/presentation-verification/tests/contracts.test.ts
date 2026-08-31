import { describe, expect, it } from 'vitest'
import {
  PRESENTATION_VERIFICATION_LIMITS,
  digestPresentationAcceptanceContract,
  parsePresentationAcceptanceContract,
  parsePresentationCompletionReceipt,
  parsePresentationRenderingFacts,
  parseVisualReviewResult,
  presentationVerificationFlags,
  renderPresentationCompletionFacts,
  PRESENTATION_GOLDEN_CASES,
  PRESENTATION_CONSISTENCY_GOLDEN,
  parsePresentationTelemetryEvent,
  emitPresentationTelemetry,
} from '../src/index'

describe('cross-host golden contract', () => {
  it('contains executable authority, page, operation and check-accounting facts', () => {
    expect(PRESENTATION_GOLDEN_CASES).toHaveLength(12)
    expect(PRESENTATION_CONSISTENCY_GOLDEN.pages).toHaveLength(8)
    expect(PRESENTATION_CONSISTENCY_GOLDEN.operations).toHaveLength(
      PRESENTATION_CONSISTENCY_GOLDEN.expectedCheckCount,
    )
    expect(
      PRESENTATION_CONSISTENCY_GOLDEN.operations.every(({ slide }) => slide >= 6 && slide <= 8),
    ).toBe(true)
  })

  it('only accepts privacy-safe telemetry dimensions', () => {
    expect(
      parsePresentationTelemetryEvent({
        host: 'pc',
        phase: 'complete',
        outcome: 'success',
        code: 'verified',
        count: 7,
        durationMs: 24,
      }),
    ).toMatchObject({ code: 'verified' })
    for (const privateKey of [
      'prompt',
      'text',
      'slideId',
      'session',
      'receipt',
      'fingerprint',
      'screenshot',
    ])
      expect(() =>
        parsePresentationTelemetryEvent({
          host: 'office',
          phase: 'visual',
          outcome: 'failed',
          code: 'pass',
          count: 1,
          durationMs: 1,
          [privateKey]: 'secret',
        }),
      ).toThrow(/private fields/)
  })
  it('keeps diagnostics fail-open when a validated sink throws', () => {
    expect(() =>
      emitPresentationTelemetry(
        () => {
          throw new Error('sink down')
        },
        {
          host: 'pc',
          phase: 'complete',
          outcome: 'success',
          code: 'verified',
          count: 1,
          durationMs: 1,
        },
      ),
    ).not.toThrow()
  })
})

const contract = () => ({
  version: 1,
  taskId: 'task-1',
  documentToken: 'doc-token',
  sessionToken: 'session-token',
  baseRevision: 'sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
  affectedSlides: [6, 7, 8],
  referenceSlides: [6],
  checks: [
    {
      id: 'check-title-color',
      kind: 'element_property',
      slide: 7,
      roleOrTarget: { kind: 'role', role: 'title' },
      property: 'color',
      expected: '#123456',
    },
    {
      id: 'check-title-position',
      kind: 'reference_match',
      slide: 7,
      referenceSlide: 6,
      role: 'title',
      properties: ['x', 'y'],
      tolerance: 0.5,
    },
    {
      id: 'check-render',
      kind: 'render_quality',
      slide: 7,
      rules: ['no_overflow', 'no_overlap'],
    },
  ],
  maxCorrectionPasses: 2,
})

const receipt = () => ({
  version: 1,
  taskId: 'task-1',
  status: 'applied_unverified',
  mutationReceiptIds: ['mutation-1'],
  passedCheckIds: ['check-title-color'],
  failedCheckIds: [],
  unavailableCheckIds: ['check-render'],
  correctionPasses: 1,
  affectedSlides: [6, 7],
  rollbackId: 'rollback-1',
  safeCode: 'screenshot_unavailable',
})

const verifiedReceipt = () => ({
  version: 1,
  taskId: 'task-1',
  status: 'verified',
  mutationReceiptIds: ['mutation-1'],
  passedCheckIds: ['check-title-color', 'check-title-position', 'check-render'],
  failedCheckIds: [],
  unavailableCheckIds: [],
  correctionPasses: 0,
  affectedSlides: [6, 7, 8],
})

describe('presentation acceptance contracts', () => {
  it('uses safe rollout defaults and independent rollback switches', () => {
    expect(presentationVerificationFlags({})).toEqual({
      planning: true,
      verifiedCompletion: true,
      visualReview: true,
      autoCorrection: false,
    })
    expect(
      presentationVerificationFlags({
        WISWORK_PRESENTATION_PLANNING: '0',
        WISWORK_PRESENTATION_VERIFIED_COMPLETION: '0',
        WISWORK_PRESENTATION_VISUAL_REVIEW: '0',
        WISWORK_PRESENTATION_AUTO_CORRECTION: '1',
      }),
    ).toEqual({
      planning: false,
      verifiedCompletion: false,
      visualReview: false,
      autoCorrection: true,
    })
    expect(() =>
      presentationVerificationFlags({ WISWORK_PRESENTATION_VISUAL_REVIEW: 'false' }),
    ).toThrow('invalid_presentation_verification_flags')
  })
  it('parses a bounded contract and returns detached data', () => {
    const input = contract()
    const parsed = parsePresentationAcceptanceContract(input)
    expect(parsed).toEqual(input)
    expect(parsed).not.toBe(input)
    expect(parsed.checks).not.toBe(input.checks)
  })

  it('canonicalizes semantically equivalent contracts before digesting', async () => {
    const left = contract()
    const right = {
      checks: [...left.checks].reverse(),
      referenceSlides: [...left.referenceSlides].reverse(),
      affectedSlides: [...left.affectedSlides].reverse(),
      baseRevision: left.baseRevision,
      sessionToken: left.sessionToken,
      documentToken: left.documentToken,
      taskId: left.taskId,
      maxCorrectionPasses: left.maxCorrectionPasses,
      version: left.version,
    }
    expect(await digestPresentationAcceptanceContract(left)).toBe(
      await digestPresentationAcceptanceContract(right),
    )
  })

  it.each([
    ['unknown object key', { ...contract(), prompt: 'raw slide text' }],
    ['unknown check key', { ...contract(), checks: [{ ...contract().checks[0], text: 'secret' }] }],
    ['wrong version', { ...contract(), version: 2 }],
    ['duplicate slides', { ...contract(), affectedSlides: [6, 6] }],
    [
      'duplicate check ids',
      { ...contract(), checks: [contract().checks[0], contract().checks[0]] },
    ],
    ['slide outside scope', { ...contract(), checks: [{ ...contract().checks[0], slide: 9 }] }],
    [
      'reference outside scope',
      { ...contract(), checks: [{ ...contract().checks[1], referenceSlide: 9 }] },
    ],
    ['unsafe token', { ...contract(), taskId: '../private/path' }],
    [
      'unbounded tolerance',
      { ...contract(), checks: [{ ...contract().checks[1], tolerance: 1_001 }] },
    ],
    [
      'non-finite scalar',
      { ...contract(), checks: [{ ...contract().checks[0], expected: Infinity }] },
    ],
  ])('rejects %s', (_label, input) => {
    expect(() => parsePresentationAcceptanceContract(input)).toThrow(TypeError)
  })

  it('enforces global check and page bounds', () => {
    const checks = Array.from(
      { length: PRESENTATION_VERIFICATION_LIMITS.maxChecks + 1 },
      (_, index) => ({ ...contract().checks[0], id: `check-${index}` }),
    )
    expect(() => parsePresentationAcceptanceContract({ ...contract(), checks })).toThrow(/bounds/i)
    expect(() =>
      parsePresentationAcceptanceContract({
        ...contract(),
        affectedSlides: Array.from(
          { length: PRESENTATION_VERIFICATION_LIMITS.maxAffectedSlides + 1 },
          (_, index) => index + 1,
        ),
      }),
    ).toThrow(/bounds/i)
  })

  it('rejects symbols, accessors, exotic prototypes, holes, and extended arrays without reading them', () => {
    let reads = 0
    const accessor = contract()
    Object.defineProperty(accessor, 'prompt', {
      enumerable: true,
      get: () => {
        reads += 1
        return 'secret'
      },
    })
    expect(() => parsePresentationAcceptanceContract(accessor)).toThrow(/accessor/i)
    expect(reads).toBe(0)

    const symbol = contract() as Record<PropertyKey, unknown>
    symbol[Symbol('secret')] = true
    expect(() => parsePresentationAcceptanceContract(symbol)).toThrow(/symbol/i)

    expect(() => parsePresentationAcceptanceContract(Object.create(contract()))).toThrow(/plain/i)

    const sparse = contract()
    sparse.affectedSlides = Array(2) as number[]
    expect(() => parsePresentationAcceptanceContract(sparse)).toThrow(/holes/i)

    const extended = contract()
    Object.defineProperty(extended.checks, 'secret', { enumerable: true, value: 'raw' })
    expect(() => parsePresentationAcceptanceContract(extended)).toThrow(/extra field/i)
  })

  it('rejects own unsafe prototype keys', () => {
    const input = JSON.parse(JSON.stringify(contract()))
    Object.defineProperty(input, '__proto__', { enumerable: true, value: {} })
    expect(() => parsePresentationAcceptanceContract(input)).toThrow(/unsafe/i)
  })
})

describe('safe visual and completion data', () => {
  it('parses bounded visual review output', () => {
    expect(
      parseVisualReviewResult({
        status: 'needs_fix',
        failedCheckIds: ['check-render'],
        observations: [{ code: 'overflow', severity: 'error', checkId: 'check-render', slide: 7 }],
        fixIntents: [
          {
            checkId: 'check-render',
            kind: 'set_property',
            roleOrTarget: { kind: 'role', role: 'body' },
            property: 'font_size',
            value: 18,
          },
        ],
      }),
    ).toMatchObject({ status: 'needs_fix' })
  })

  it('rejects raw narrative and contradictory visual outputs', () => {
    expect(() =>
      parseVisualReviewResult({
        status: 'pass',
        failedCheckIds: ['check-render'],
        observations: [],
        fixIntents: [],
      }),
    ).toThrow()
    expect(() =>
      parseVisualReviewResult({
        status: 'cannot_verify',
        failedCheckIds: [],
        observations: [{ code: 'overflow', severity: 'error', message: 'raw slide text' }],
        fixIntents: [],
      }),
    ).toThrow(/unknown/i)
  })

  it.each([
    ['numeric geometry', 'x', '12'],
    ['boolean styling', 'bold', 1],
    ['canonical color', 'fill_color', 'red'],
    ['string text', 'text', false],
  ])('validates fix intent values as strictly as %s checks', (_label, property, value) => {
    expect(() =>
      parseVisualReviewResult({
        status: 'needs_fix',
        failedCheckIds: ['check-render'],
        observations: [],
        fixIntents: [
          {
            checkId: 'check-render',
            kind: 'set_property',
            roleOrTarget: { kind: 'role', role: 'body' },
            property,
            value,
          },
        ],
      }),
    ).toThrow(/value/i)
  })

  it('parses only bounded, non-content rendering facts', () => {
    const facts = parsePresentationRenderingFacts({
      contractDigest: 'sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
      revision: 'sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb',
      screenshots: [
        { slide: 6, role: 'reference', mediaToken: 'media-1', bytes: 1024 },
        { slide: 7, role: 'affected', mediaToken: 'media-2', bytes: 2048 },
      ],
      deterministicResults: [
        { checkId: 'check-title-color', status: 'pass' },
        { checkId: 'check-render', status: 'unavailable', code: 'unsupported_check' },
      ],
    })
    expect(facts.screenshots).toHaveLength(2)

    expect(() =>
      parsePresentationRenderingFacts({
        ...facts,
        screenshots: Array.from(
          { length: PRESENTATION_VERIFICATION_LIMITS.maxScreenshotsPerPass + 1 },
          (_, index) => ({
            slide: index + 1,
            role: 'affected',
            mediaToken: `media-${index}`,
            bytes: 1,
          }),
        ),
      }),
    ).toThrow(/bounds/i)
    expect(() =>
      parsePresentationRenderingFacts({ ...facts, screenshot: 'data:image/png;base64,...' }),
    ).toThrow()
  })

  it('rejects unknown fields before measuring the bounded envelope', () => {
    expect(() =>
      parsePresentationRenderingFacts({
        contractDigest: 'sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
        revision: 'sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb',
        screenshots: [{ slide: 7, role: 'affected', mediaToken: 'media-1', bytes: 1 }],
        deterministicResults: [],
        padding: 'x'.repeat(PRESENTATION_VERIFICATION_LIMITS.maxVisualRequestBytes),
      }),
    ).toThrow(/unknown field padding/i)
  })

  it('rejects hostile rendering shapes promptly without unbounded traversal or accessor reads', () => {
    const hugeSparse: unknown[] = []
    hugeSparse.length = 0xffff_ffff
    expect(() =>
      parsePresentationRenderingFacts({
        contractDigest: 'sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
        revision: 'sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb',
        screenshots: hugeSparse,
        deterministicResults: [],
      }),
    ).toThrow(/screenshots length.*bounds/i)

    let reads = 0
    const deepUnknown: Record<string, unknown> = {}
    let cursor = deepUnknown
    for (let index = 0; index < 10_000; index += 1) {
      const next: Record<string, unknown> = {}
      cursor.next = next
      cursor = next
    }
    Object.defineProperty(cursor, 'secret', {
      enumerable: true,
      get: () => {
        reads += 1
        return 'raw content'
      },
    })
    expect(() =>
      parsePresentationRenderingFacts({
        contractDigest: 'sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
        revision: 'sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb',
        screenshots: [],
        deterministicResults: [],
        unknown: deepUnknown,
      }),
    ).toThrow(/unknown field unknown/i)
    expect(reads).toBe(0)
  })

  it('counts the envelope in addition to declared screenshot payload bytes', () => {
    expect(() =>
      parsePresentationRenderingFacts({
        contractDigest: 'sha256:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa',
        revision: 'sha256:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb',
        screenshots: [
          {
            slide: 7,
            role: 'affected',
            mediaToken: 'media-1',
            bytes: PRESENTATION_VERIFICATION_LIMITS.maxVisualRequestBytes,
          },
        ],
        deterministicResults: [],
      }),
    ).toThrow(/serialized visual request.*bounds/i)
  })

  it('parses receipts and derives deterministic, privacy-safe response facts', () => {
    const appliedContract = {
      ...contract(),
      affectedSlides: [6, 7],
      checks: [contract().checks[0], contract().checks[2]],
    }
    const parsed = parsePresentationCompletionReceipt(receipt(), appliedContract)
    expect(renderPresentationCompletionFacts(parsed, appliedContract)).toEqual({
      status: 'applied_unverified',
      affectedSlides: [6, 7],
      passedCount: 1,
      failedCount: 0,
      unavailableCount: 1,
      correctionPasses: 1,
      rollbackAvailable: true,
      safeCode: 'screenshot_unavailable',
    })
  })

  it('binds receipt identity, affected slides, correction limit, and exact check accounting', () => {
    expect(() =>
      parsePresentationCompletionReceipt({ ...verifiedReceipt(), taskId: 'task-2' }, contract()),
    ).toThrow(/task/i)
    expect(() =>
      parsePresentationCompletionReceipt(
        { ...verifiedReceipt(), affectedSlides: [6, 7] },
        contract(),
      ),
    ).toThrow(/affected/i)
    expect(() =>
      parsePresentationCompletionReceipt(
        { ...verifiedReceipt(), passedCheckIds: ['check-title-color', 'check-render'] },
        contract(),
      ),
    ).toThrow(/account/i)
    expect(() =>
      parsePresentationCompletionReceipt(
        { ...verifiedReceipt(), passedCheckIds: [...verifiedReceipt().passedCheckIds, 'unknown'] },
        contract(),
      ),
    ).toThrow(/account/i)
    expect(() =>
      parsePresentationCompletionReceipt(
        { ...verifiedReceipt(), correctionPasses: 1 },
        { ...contract(), maxCorrectionPasses: 0 },
      ),
    ).toThrow(/correction/i)
  })

  it.each([
    ['verified without mutation', { ...verifiedReceipt(), mutationReceiptIds: [] }],
    [
      'verified with failed check',
      {
        ...verifiedReceipt(),
        passedCheckIds: ['check-title-color', 'check-title-position'],
        failedCheckIds: ['check-render'],
      },
    ],
    [
      'failed with mutation',
      { ...verifiedReceipt(), status: 'failed', mutationReceiptIds: ['m-1'] },
    ],
    [
      'unchanged with unavailable check',
      {
        ...verifiedReceipt(),
        status: 'unchanged',
        mutationReceiptIds: [],
        passedCheckIds: ['check-title-color', 'check-title-position'],
        unavailableCheckIds: ['check-render'],
      },
    ],
    [
      'unchanged after correction',
      { ...verifiedReceipt(), status: 'unchanged', mutationReceiptIds: [], correctionPasses: 1 },
    ],
    [
      'applied unverified without unproved checks',
      { ...verifiedReceipt(), status: 'applied_unverified' },
    ],
    [
      'needs user without confirmation evidence',
      {
        ...verifiedReceipt(),
        status: 'needs_user',
        mutationReceiptIds: [],
        passedCheckIds: ['check-title-color', 'check-title-position'],
        unavailableCheckIds: ['check-render'],
      },
    ],
  ])('rejects incoherent contract-bound status: %s', (_label, input) => {
    expect(() => parsePresentationCompletionReceipt(input, contract())).toThrow()
  })

  it('accepts coherent verified, unchanged, failed, applied-unverified, and needs-user receipts', () => {
    expect(parsePresentationCompletionReceipt(verifiedReceipt(), contract()).status).toBe(
      'verified',
    )
    expect(
      parsePresentationCompletionReceipt(
        { ...verifiedReceipt(), status: 'unchanged', mutationReceiptIds: [] },
        contract(),
      ).status,
    ).toBe('unchanged')
    expect(
      parsePresentationCompletionReceipt(
        {
          ...verifiedReceipt(),
          status: 'failed',
          mutationReceiptIds: [],
          passedCheckIds: ['check-title-color'],
          failedCheckIds: ['check-title-position', 'check-render'],
          safeCode: 'mutation_failed',
        },
        contract(),
      ).status,
    ).toBe('failed')
    expect(
      parsePresentationCompletionReceipt(
        {
          ...verifiedReceipt(),
          status: 'applied_unverified',
          passedCheckIds: ['check-title-color'],
          unavailableCheckIds: ['check-title-position', 'check-render'],
          safeCode: 'screenshot_unavailable',
        },
        contract(),
      ).status,
    ).toBe('applied_unverified')
    expect(
      parsePresentationCompletionReceipt(
        {
          ...verifiedReceipt(),
          status: 'needs_user',
          mutationReceiptIds: [],
          passedCheckIds: ['check-title-color'],
          failedCheckIds: ['check-title-position'],
          unavailableCheckIds: ['check-render'],
          safeCode: 'confirmation_required',
        },
        contract(),
      ).status,
    ).toBe('needs_user')
  })

  it.each([
    ['verified failure code', { ...verifiedReceipt(), safeCode: 'mutation_failed' }],
    [
      'applied with only failed checks',
      {
        ...verifiedReceipt(),
        status: 'applied_unverified',
        passedCheckIds: ['check-title-color', 'check-title-position'],
        failedCheckIds: ['check-render'],
        safeCode: 'unsupported_check',
      },
    ],
    [
      'applied uncertain office state',
      {
        ...verifiedReceipt(),
        status: 'applied_unverified',
        passedCheckIds: ['check-title-color'],
        unavailableCheckIds: ['check-title-position', 'check-render'],
        safeCode: 'office_state_uncertain',
      },
    ],
    [
      'applied mutation failure code',
      {
        ...verifiedReceipt(),
        status: 'applied_unverified',
        passedCheckIds: ['check-title-color'],
        unavailableCheckIds: ['check-title-position', 'check-render'],
        safeCode: 'mutation_failed',
      },
    ],
    [
      'needs user screenshot code',
      {
        ...verifiedReceipt(),
        status: 'needs_user',
        mutationReceiptIds: [],
        passedCheckIds: ['check-title-color'],
        unavailableCheckIds: ['check-title-position', 'check-render'],
        safeCode: 'screenshot_unavailable',
      },
    ],
    [
      'failed screenshot code',
      {
        ...verifiedReceipt(),
        status: 'failed',
        mutationReceiptIds: [],
        passedCheckIds: ['check-title-color'],
        failedCheckIds: ['check-title-position', 'check-render'],
        safeCode: 'screenshot_unavailable',
      },
    ],
    [
      'unchanged failure code',
      {
        ...verifiedReceipt(),
        status: 'unchanged',
        mutationReceiptIds: [],
        safeCode: 'mutation_failed',
      },
    ],
  ])('rejects status/safe-code contradiction: %s', (_label, input) => {
    expect(() => parsePresentationCompletionReceipt(input, contract())).toThrow(
      /status|code|applied/i,
    )
  })

  it.each([
    ['verified', verifiedReceipt()],
    [
      'applied screenshot unavailable',
      {
        ...verifiedReceipt(),
        status: 'applied_unverified',
        passedCheckIds: ['check-title-color'],
        unavailableCheckIds: ['check-title-position', 'check-render'],
        safeCode: 'screenshot_unavailable',
      },
    ],
    [
      'applied unsupported check',
      {
        ...verifiedReceipt(),
        status: 'applied_unverified',
        passedCheckIds: ['check-title-color'],
        unavailableCheckIds: ['check-title-position', 'check-render'],
        safeCode: 'unsupported_check',
      },
    ],
    [
      'needs user unsupported',
      {
        ...verifiedReceipt(),
        status: 'needs_user',
        mutationReceiptIds: [],
        passedCheckIds: ['check-title-color'],
        unavailableCheckIds: ['check-title-position', 'check-render'],
        safeCode: 'unsupported_check',
      },
    ],
    [
      'failed uncertain office state',
      {
        ...verifiedReceipt(),
        status: 'failed',
        mutationReceiptIds: [],
        passedCheckIds: ['check-title-color'],
        unavailableCheckIds: ['check-title-position', 'check-render'],
        safeCode: 'office_state_uncertain',
      },
    ],
    [
      'unchanged cancelled',
      { ...verifiedReceipt(), status: 'unchanged', mutationReceiptIds: [], safeCode: 'cancelled' },
    ],
  ])('accepts coherent status/safe-code pair: %s', (_label, input) => {
    expect(parsePresentationCompletionReceipt(input, contract()).status).toBe(input.status)
  })

  it.each([
    ['raw content', { ...receipt(), message: 'secret text' }],
    ['over-correction', { ...receipt(), correctionPasses: 3 }],
    ['duplicate accounting', { ...receipt(), failedCheckIds: ['check-title-color'] }],
    ['verified unavailable', { ...receipt(), status: 'verified' }],
    ['unchanged mutation', { ...receipt(), status: 'unchanged' }],
    ['free-form safe code', { ...receipt(), safeCode: 'slide-content-goes-here' }],
  ])('rejects unsafe receipt: %s', (_label, input) => {
    expect(() => parsePresentationCompletionReceipt(input, contract())).toThrow()
  })
})
