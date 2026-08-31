import { describe, expect, it } from 'vitest'
import {
  PRESENTATION_VERIFICATION_LIMITS,
  digestPresentationAcceptanceContract,
  parsePresentationAcceptanceContract,
  parsePresentationCompletionReceipt,
  parsePresentationRenderingFacts,
  parseVisualReviewResult,
  renderPresentationCompletionFacts,
} from '../src/index'

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

describe('presentation acceptance contracts', () => {
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

  it('parses receipts and derives deterministic, privacy-safe response facts', () => {
    const parsed = parsePresentationCompletionReceipt(receipt())
    expect(renderPresentationCompletionFacts(parsed)).toEqual({
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

  it.each([
    ['raw content', { ...receipt(), message: 'secret text' }],
    ['over-correction', { ...receipt(), correctionPasses: 3 }],
    ['duplicate accounting', { ...receipt(), failedCheckIds: ['check-title-color'] }],
    ['verified unavailable', { ...receipt(), status: 'verified' }],
    ['unchanged mutation', { ...receipt(), status: 'unchanged' }],
    ['free-form safe code', { ...receipt(), safeCode: 'slide-content-goes-here' }],
  ])('rejects unsafe receipt: %s', (_label, input) => {
    expect(() => parsePresentationCompletionReceipt(input)).toThrow()
  })
})
