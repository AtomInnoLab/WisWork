import { describe, expect, it, vi } from 'vitest'
import {
  canonicalPowerPointVerificationBinding,
  createBrowserPowerPointVerificationAuthority,
  createOfficePowerPointVerification,
  type OfficePowerPointVisualReviewer,
} from '../src/skills/powerpoint/powerpoint-verification.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import {
  BrowserPowerPointAdapter,
  type PowerPointAdapter,
} from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { PRESENTATION_CONSISTENCY_GOLDEN } from '@wiswork/presentation-verification'

const authority = (overrides = {}) => ({
  acquire: vi.fn().mockResolvedValue({
    documentToken: 'doc-1',
    sessionToken: 'session-1',
    revision: `sha256:${'1'.repeat(64)}`,
  }),
  current: vi.fn().mockResolvedValue({
    documentToken: 'doc-1',
    sessionToken: 'session-1',
    revision: `sha256:${'2'.repeat(64)}`,
  }),
  readShape: vi.fn().mockResolvedValue({
    slideId: 'slide-1',
    shapeId: 'shape-1',
    text: 'After',
    color: '#112233',
    left: 10,
    top: 20,
    width: 100,
    height: 30,
  }),
  readSlide: vi.fn().mockResolvedValue({ slideId: 'slide-1', backgroundColor: '#FFFFFF' }),
  captureScreenshot: vi.fn().mockResolvedValue({ mime: 'image/png', base64: 'iVBORw0KGgoAAAA=' }),
  ...overrides,
})

const textCall = {
  id: 'c1',
  name: 'edit_slide_text',
  input: { slide_index: 0, shape_id: 'shape-1', text: 'After' },
}

it('removes the production verification hook when verified completion is rolled back', () => {
  const skill = createPowerPointSkill({
    adapter: productionAdapter({ text: 'Before', left: 5 }),
    proposals: createStructuredProposalController(),
    verificationAuthority: authority(),
    presentationFlags: {
      planning: true,
      verifiedCompletion: false,
      visualReview: true,
      autoCorrection: false,
    },
  })
  expect(skill.presentation).toBeUndefined()
})
const verificationBinding = (call: typeof textCall, targets = ['slide-1/shape-1']) =>
  canonicalPowerPointVerificationBinding(call, targets)

function productionAdapter(state: { text: string; left: number }): PowerPointAdapter {
  return {
    inspectSlideMasters: vi.fn().mockResolvedValue({ masters: [] }),
    executeMasterOperations: vi.fn(),
    screenshotSlide: vi.fn(),
    listSlideShapes: vi.fn().mockImplementation(() =>
      Promise.resolve({
        slideId: 'slide-1',
        slideIndex: 0,
        shapes: [
          {
            id: 'shape-1',
            name: 'Title',
            type: 'TextBox',
            left: state.left,
            top: 20,
            width: 100,
            height: 30,
          },
        ],
      }),
    ),
    readSlideText: vi.fn().mockImplementation(() =>
      Promise.resolve({
        slideId: 'slide-1',
        shapeId: 'shape-1',
        text: state.text,
        paragraphs: [state.text],
      }),
    ),
    verifySlides: vi.fn().mockResolvedValue({ slideWidth: 960, slideHeight: 540, slides: [] }),
    snapshotSlide: vi
      .fn()
      .mockImplementation(() =>
        Promise.resolve({ slideId: 'slide-1', fingerprint: `slide-1:${state.text}:${state.left}` }),
      ),
    editSlideText: vi.fn().mockImplementation(async (_slide, _shape, text) => {
      state.text = text
    }),
    duplicateSlide: vi.fn(),
    exportSlidePackage: vi.fn(),
    replaceSlidePackage: vi.fn(),
    executeDeclarative: vi.fn().mockImplementation(async (operations) => {
      for (const operation of operations) {
        if (operation.op === 'set_shape_text') state.text = operation.text
        if (operation.op === 'set_shape_geometry') state.left = operation.left
      }
      return { createdShapeIds: [] }
    }),
  }
}

describe('Office PowerPoint presentation verification', () => {
  async function visualSubject(review: OfficePowerPointVisualReviewer['review'], overrides = {}) {
    const call = { ...textCall, id: 'visual-call', input: { ...textCall.input, text: 'Final' } }
    let applied = false
    const source = authority({
      readShape: vi.fn().mockImplementation(() =>
        Promise.resolve({
          slideId: 'slide-1',
          shapeId: 'shape-1',
          text: applied ? 'Final' : 'Before',
          left: 10,
          top: 20,
          width: 100,
          height: 30,
        }),
      ),
      ...overrides,
    })
    const reviewer = { review: vi.fn().mockImplementation(review) }
    const subject = createOfficePowerPointVerification({
      authority: source,
      reviewer,
      taskId: () => 'task-visual',
      flags: { planning: true, verifiedCompletion: true, visualReview: true, autoCorrection: true },
    })
    const enrolled = await subject.enroll([call], undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    applied = true
    const proposal = {
      id: 'visual-proposal',
      toolName: 'edit_slide_text',
      fingerprint: 'state',
      targets: ['slide-1/shape-1'],
      verificationBinding: canonicalPowerPointVerificationBinding(call, ['slide-1/shape-1']),
    }
    subject.recordProposal(proposal)
    subject.recordSettlement({ id: proposal.id, status: 'confirmed' })
    return { subject, contract: enrolled.contract, reviewer, source }
  }

  it('binds postwrite screenshots to authority and sends only strict isolated review facts', async () => {
    const run = await visualSubject(async () => ({
      status: 'pass',
      failedCheckIds: [],
      observations: [],
      fixIntents: [],
    }))
    await expect(
      run.subject.complete({
        contract: run.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({ kind: 'receipt', receipt: { status: 'verified' } })
    expect(run.source.current).toHaveBeenCalledTimes(3)
    const request = run.reviewer.review.mock.calls[0]![0]
    expect(request.isolation).toEqual({ tools: [], maxTurns: 1 })
    expect(JSON.stringify(request.facts)).not.toContain('Final')
  })

  it('never reports an applied write verified when visual review is disabled', async () => {
    const run = await visualSubject(async () => ({
      status: 'pass',
      failedCheckIds: [],
      observations: [],
      fixIntents: [],
    }))
    let applied = false
    const subject = createOfficePowerPointVerification({
      authority: authority({
        readShape: vi.fn().mockImplementation(async () => ({
          slideId: 'slide-1',
          shapeId: 'shape-1',
          text: applied ? 'Final' : 'Before',
          left: 10,
          top: 20,
          width: 100,
          height: 30,
        })),
      }),
      reviewer: run.reviewer,
      taskId: () => 'task-visual-disabled',
      flags: {
        planning: true,
        verifiedCompletion: true,
        visualReview: false,
        autoCorrection: false,
      },
    })
    const call = { ...textCall, id: 'visual-disabled', input: { ...textCall.input, text: 'Final' } }
    const enrolled = await subject.enroll([call], undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    applied = true
    const proposal = {
      id: 'visual-disabled-proposal',
      toolName: call.name,
      fingerprint: 'state',
      targets: ['slide-1/shape-1'],
      verificationBinding: canonicalPowerPointVerificationBinding(call, ['slide-1/shape-1']),
    }
    subject.recordProposal(proposal)
    subject.recordSettlement({ id: proposal.id, status: 'confirmed' })
    await expect(
      subject.complete({
        contract: enrolled.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode: 'visual_disabled' },
    })
    expect(run.reviewer.review).not.toHaveBeenCalled()
  })

  it('requests one bounded safe correction, stops at two, and rejects unsafe fixes', async () => {
    const makeReview: (value?: string) => OfficePowerPointVisualReviewer['review'] =
      (value = 'Final') =>
      async (request) => ({
        status: 'needs_fix' as const,
        failedCheckIds: [request.facts.deterministicResults[0].checkId],
        observations: [],
        fixIntents: [
          {
            checkId: request.facts.deterministicResults[0].checkId,
            kind: 'set_property' as const,
            roleOrTarget: { kind: 'target' as const, targetToken: 'target-0' },
            property: 'text' as const,
            value,
          },
        ],
      })
    const safe = await visualSubject(makeReview())
    await expect(
      safe.subject.complete({
        contract: safe.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({ kind: 'correct' })
    const capped = await visualSubject(makeReview())
    await expect(
      capped.subject.complete({
        contract: capped.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 2,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'needs_user', correctionPasses: 2 },
    })
    const converges = await visualSubject(
      vi.fn().mockImplementationOnce(makeReview()).mockResolvedValueOnce({
        status: 'pass',
        failedCheckIds: [],
        observations: [],
        fixIntents: [],
      }),
    )
    await expect(
      converges.subject.complete({
        contract: converges.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({ kind: 'correct' })
    const correctionCall = {
      ...textCall,
      id: 'correction-applied',
      input: { ...textCall.input, text: 'Final' },
    }
    await converges.subject.enroll([correctionCall], converges.contract)
    converges.subject.recordProposal({
      id: 'correction-proposal',
      toolName: 'edit_slide_text',
      fingerprint: 'state-2',
      targets: ['slide-1/shape-1'],
      verificationBinding: canonicalPowerPointVerificationBinding(correctionCall, [
        'slide-1/shape-1',
      ]),
    })
    converges.subject.recordSettlement({ id: 'correction-proposal', status: 'confirmed' })
    await expect(
      converges.subject.complete({
        contract: converges.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 1,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'verified', correctionPasses: 1 },
    })
    await expect(
      safe.subject.enroll(
        [
          {
            ...textCall,
            id: 'correction-call',
            input: { ...textCall.input, text: 'Final' },
          },
        ],
        safe.contract,
      ),
    ).resolves.toMatchObject({ kind: 'ready', contract: safe.contract })
    await expect(
      safe.subject.enroll(
        [
          {
            ...textCall,
            id: 'expanded-call',
            input: { ...textCall.input, shape_id: 'other', text: 'Final' },
          },
        ],
        safe.contract,
      ),
    ).resolves.toMatchObject({ kind: 'clarify' })
    const unsafe = await visualSubject(makeReview('Different'))
    await expect(
      unsafe.subject.complete({
        contract: unsafe.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'needs_user', safeCode: 'confirmation_required' },
    })
  })

  it.each([
    {
      id: 'delete',
      name: 'execute_office_js',
      input: {
        program: {
          version: 1,
          operations: [{ op: 'delete_shape', slide_index: 0, shape_id: 'shape-1' }],
        },
      },
    },
    { id: 'duplicate', name: 'duplicate_slide', input: { slide_index: 0 } },
    { id: 'master', name: 'edit_slide_master', input: { program: {} } },
    { id: 'chart', name: 'edit_slide_chart', input: { slide_index: 0, program: {} } },
    { id: 'xml', name: 'edit_slide_xml', input: { slide_index: 0, program: {} } },
  ])('rejects unsupported correction tool $name before dispatch', async (unsafeCall) => {
    const run = await visualSubject(async () => ({
      status: 'pass',
      failedCheckIds: [],
      observations: [],
      fixIntents: [],
    }))
    await expect(run.subject.enroll([unsafeCall], run.contract)).resolves.toMatchObject({
      kind: 'clarify',
    })
  })

  it('rejects a mixed supported and unsupported correction batch atomically', async () => {
    const run = await visualSubject(async () => ({
      status: 'pass',
      failedCheckIds: [],
      observations: [],
      fixIntents: [],
    }))
    await expect(
      run.subject.enroll(
        [
          { ...textCall, id: 'safe-correction', input: { ...textCall.input, text: 'Final' } },
          { id: 'unsafe-correction', name: 'duplicate_slide', input: { slide_index: 0 } },
        ],
        run.contract,
      ),
    ).resolves.toMatchObject({ kind: 'clarify' })
  })

  it.each([
    {
      status: 'needs_fix',
      failedCheckIds: ['unknown'],
      observations: [],
      fixIntents: [
        {
          checkId: 'unknown',
          kind: 'set_property',
          roleOrTarget: { kind: 'target', targetToken: 'target-0' },
          property: 'text',
          value: 'Final',
        },
      ],
    },
    {
      status: 'needs_fix',
      failedCheckIds: ['visual-call-op0-text'],
      observations: [],
      fixIntents: [
        {
          checkId: 'visual-call-op0-text',
          kind: 'set_property',
          roleOrTarget: { kind: 'target', targetToken: 'target-0' },
          property: 'text',
          value: 'Final',
        },
        {
          checkId: 'other',
          kind: 'set_property',
          roleOrTarget: { kind: 'target', targetToken: 'target-0' },
          property: 'text',
          value: 'Final',
        },
      ],
    },
    {
      status: 'cannot_verify',
      failedCheckIds: ['visual-call-op0-text', 'visual-call-op0-text'],
      observations: [],
      fixIntents: [],
    },
  ])('fails closed on invalid reviewer check accounting', async (review) => {
    const run = await visualSubject(async () => review as never)
    await expect(
      run.subject.complete({
        contract: run.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode: 'verification_invalid' },
    })
  })

  it.each([
    [
      'screenshot failure',
      { captureScreenshot: vi.fn().mockRejectedValue(new Error('capture')) },
      'screenshot_unavailable',
    ],
    [
      'stale screenshot authority',
      {
        current: vi
          .fn()
          .mockResolvedValueOnce({
            documentToken: 'doc-1',
            sessionToken: 'session-1',
            revision: `sha256:${'2'.repeat(64)}`,
          })
          .mockResolvedValueOnce({
            documentToken: 'doc-1',
            sessionToken: 'session-1',
            revision: `sha256:${'2'.repeat(64)}`,
          })
          .mockResolvedValueOnce({
            documentToken: 'doc-1',
            sessionToken: 'session-1',
            revision: `sha256:${'3'.repeat(64)}`,
          }),
      },
      'stale_authority',
    ],
  ])('preserves applied truth on %s', async (_name, overrides, safeCode) => {
    const run = await visualSubject(
      async () => ({ status: 'pass', failedCheckIds: [], observations: [], fixIntents: [] }),
      overrides,
    )
    await expect(
      run.subject.complete({
        contract: run.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode, mutationReceiptIds: ['visual-proposal'] },
    })
  })

  it('normalizes real mixed declarative proposal targets and closes from actual confirmation receipts', async () => {
    const state = { text: 'Before', left: 5 }
    const adapter = productionAdapter(state)
    const proposals = createStructuredProposalController()
    const verificationAuthority = authority({
      readShape: vi.fn().mockImplementation(() =>
        Promise.resolve({
          slideId: 'slide-1',
          shapeId: 'shape-1',
          text: state.text,
          left: state.left,
          top: 20,
          width: 100,
          height: 30,
        }),
      ),
    })
    const skill = createPowerPointSkill({ adapter, proposals, verificationAuthority })
    const call = {
      id: 'mixed',
      name: 'execute_office_js',
      input: {
        program: {
          version: 1,
          operations: [
            { op: 'set_shape_text', slide_index: 0, shape_id: 'shape-1', text: 'After' },
            {
              op: 'set_shape_geometry',
              slide_index: 0,
              shape_id: 'shape-1',
              left: 10,
              top: 20,
              width: 100,
              height: 30,
            },
          ],
        },
      },
    }
    const enrolled = await skill.presentation!.enroll!([call], undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    await skill.executeTool(call)
    expect(proposals.pending()?.impact.targets).toEqual(['slide-1/shape-1', 'slide-1/shape-1'])
    await proposals.confirm(proposals.pending()!.id)
    await expect(
      skill.presentation!.complete({
        contract: enrolled.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'verified', mutationReceiptIds: [expect.any(String)] },
    })
  })

  it('verifies a real confirmation-gated declarative text color through durable authority', async () => {
    let color = '#000000'
    const adapter = productionAdapter({ text: 'Title', left: 10 })
    adapter.executeDeclarative = vi.fn().mockImplementation(async (operations) => {
      color = operations[0].color
      return { createdShapeIds: [] }
    })
    adapter.readShapeTextStyle = vi.fn().mockImplementation(async () => ({ color }))
    const verificationAuthority = authority({
      readShape: vi.fn().mockImplementation(async () => ({
        slideId: 'slide-1',
        shapeId: 'shape-1',
        text: 'Title',
        color,
        left: 10,
        top: 20,
        width: 100,
        height: 30,
      })),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter, proposals, verificationAuthority })
    const call = {
      id: 'style-color',
      name: 'execute_office_js',
      input: {
        program: {
          version: 1,
          operations: [
            { op: 'set_shape_text_style', slide_index: 0, shape_id: 'shape-1', color: '#2457A7' },
          ],
        },
      },
    }
    const enrolled = await skill.presentation!.enroll!([call], undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    await skill.executeTool!(call)
    await proposals.confirm(proposals.pending()!.id)
    await expect(
      skill.presentation!.complete({
        contract: enrolled.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'verified', passedCheckIds: [expect.stringContaining('color')] },
    })
    expect(color).toBe('#2457A7')
  })

  it('executes the shared pages 6-8 consistency golden through skill, proposal, authority and screenshots', async () => {
    const state = new Map<string, Record<string, string | number>>(
      PRESENTATION_CONSISTENCY_GOLDEN.pages.flatMap((page) =>
        page.shapes.map((shape) => [`${page.slide}:${shape.role}`, { ...shape }]),
      ),
    )
    const outside = JSON.stringify([...state].filter(([key]) => /^[1-5]:/.test(key)))
    const slidesItems = PRESENTATION_CONSISTENCY_GOLDEN.pages.map((page) => {
      const items = page.shapes.map((entry) => {
        const value = state.get(`${page.slide}:${entry.role}`)!
        const font = { load: vi.fn(), name: 'Arial', size: 18, bold: false, italic: false }
        Object.defineProperty(font, 'color', {
          get: () => value.color,
          set: (next) => {
            value.color = next
          },
          enumerable: true,
        })
        const textRange = { text: entry.role, load: vi.fn(), font }
        const shape = {
          id: entry.role,
          name: entry.role,
          type: 'TextBox',
          load: vi.fn(),
          textFrame: { hasText: true, load: vi.fn(), textRange },
        }
        for (const [field, source] of [
          ['left', 'x'],
          ['top', 'y'],
          ['width', 'width'],
          ['height', 'height'],
        ] as const)
          Object.defineProperty(shape, field, {
            get: () => value[source] ?? (field === 'width' ? 100 : field === 'height' ? 30 : 0),
            set: (next) => {
              value[source] = next
            },
            enumerable: true,
          })
        return shape
      })
      const shapes = {
        items,
        load: vi.fn(),
        getItem: vi.fn((id: string) => items.find((item) => item.id === id)),
      }
      return {
        id: `slide-${page.slide}`,
        load: vi.fn(),
        shapes,
        getImageAsBase64: vi.fn(() => ({ value: 'iVBORw0KGgoAAAA=' })),
      }
    })
    const slides = {
      items: slidesItems,
      load: vi.fn(),
      getCount: vi.fn(() => ({ value: 8 })),
      getItemAt: vi.fn((index: number) => slidesItems[index]),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          document: { url: 'https://example/golden.pptx' },
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({
            presentation: {
              slides,
              pageSetup: { slideWidth: 960, slideHeight: 540, load: vi.fn() },
            },
            sync: vi.fn().mockResolvedValue(undefined),
          }),
      },
    })
    const adapter = new BrowserPowerPointAdapter()
    const verificationAuthority = createBrowserPowerPointVerificationAuthority(adapter)
    const proposals = createStructuredProposalController()
    const reviewer = {
      review: vi.fn().mockResolvedValue({
        status: 'pass',
        failedCheckIds: [],
        observations: [],
        fixIntents: [],
      }),
    }
    const telemetry = vi.fn()
    const skill = createPowerPointSkill({
      adapter,
      proposals,
      verificationAuthority,
      visualReviewer: reviewer,
      presentationTelemetry: telemetry,
    })
    const operations = PRESENTATION_CONSISTENCY_GOLDEN.operations
      .filter((operation) => operation.property === 'color' || operation.property === 'x')
      .map((operation) =>
        operation.property === 'color'
          ? {
              op: 'set_shape_text_style',
              slide_index: operation.slide - 1,
              shape_id: operation.targetToken.split(':')[1],
              color: operation.value,
            }
          : {
              op: 'set_shape_geometry',
              slide_index: operation.slide - 1,
              shape_id: 'title',
              left: operation.value,
              top: 48,
              width: 816,
              height: 54,
            },
      )
    const call = {
      id: 'golden-pages-6-8',
      name: 'execute_office_js',
      input: { program: { version: 1, operations } },
    }
    const enrolled = await skill.presentation!.enroll!([call], undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    await skill.executeTool!(call)
    await proposals.confirm(proposals.pending()!.id)
    const completed = await skill.presentation!.complete({
      contract: enrolled.contract,
      mutated: true,
      cancelled: false,
      correctionPasses: 0,
    })
    expect(completed).toMatchObject({
      kind: 'receipt',
      receipt: { status: 'verified', passedCheckIds: { length: 17 } },
    })
    expect(reviewer.review).toHaveBeenCalledOnce()
    expect(reviewer.review.mock.calls[0]![0].facts.screenshots).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ slide: 6, role: 'affected' }),
        expect.objectContaining({ slide: 6, role: 'reference' }),
      ]),
    )
    expect(telemetry).toHaveBeenCalledWith({
      host: 'office',
      phase: 'complete',
      outcome: 'success',
      code: 'verified',
      count: 3,
      durationMs: expect.any(Number),
    })
    expect(telemetry.mock.calls.map(([event]) => event.phase)).toEqual([
      'plan',
      'dispatch',
      'deterministic',
      'visual',
      'complete',
    ])
    expect(JSON.stringify(telemetry.mock.calls)).not.toMatch(/slide-|golden-|title|body|emphasis/)
    expect(JSON.stringify([...state].filter(([key]) => /^[1-5]:/.test(key)))).toBe(outside)
  })

  it.each(['text', 'declarative'] as const)(
    'verifies two real sequential %s proposals from evolving authoritative state',
    async (kind) => {
      const state = { text: 'Start', left: 5 }
      const adapter = productionAdapter(state)
      const proposals = createStructuredProposalController()
      const verificationAuthority = authority({
        readShape: vi.fn().mockImplementation(() =>
          Promise.resolve({
            slideId: 'slide-1',
            shapeId: 'shape-1',
            text: state.text,
            left: state.left,
            top: 20,
            width: 100,
            height: 30,
          }),
        ),
      })
      const skill = createPowerPointSkill({ adapter, proposals, verificationAuthority })
      const calls =
        kind === 'text'
          ? [
              { ...textCall, id: 'first-write', input: { ...textCall.input, text: 'Middle' } },
              { ...textCall, id: 'second-write', input: { ...textCall.input, text: 'Final' } },
            ]
          : [
              {
                id: 'first-program',
                name: 'execute_office_js',
                input: {
                  program: {
                    version: 1,
                    operations: [
                      {
                        op: 'set_shape_geometry',
                        slide_index: 0,
                        shape_id: 'shape-1',
                        left: 8,
                        top: 20,
                        width: 100,
                        height: 30,
                      },
                    ],
                  },
                },
              },
              {
                id: 'second-program',
                name: 'execute_office_js',
                input: {
                  program: {
                    version: 1,
                    operations: [
                      {
                        op: 'set_shape_geometry',
                        slide_index: 0,
                        shape_id: 'shape-1',
                        left: 12,
                        top: 20,
                        width: 100,
                        height: 30,
                      },
                    ],
                  },
                },
              },
            ]
      const enrolled = await skill.presentation!.enroll!(calls, undefined)
      if (enrolled.kind !== 'ready') throw new Error('not ready')
      for (const call of calls) {
        await skill.executeTool(call)
        await proposals.confirm(proposals.pending()!.id)
      }
      await expect(
        skill.presentation!.complete({
          contract: enrolled.contract,
          mutated: true,
          cancelled: false,
          correctionPasses: 0,
        }),
      ).resolves.toMatchObject({
        kind: 'receipt',
        receipt: {
          status: 'verified',
          mutationReceiptIds: [expect.any(String), expect.any(String)],
        },
      })
      expect(kind === 'text' ? state.text : state.left).toBe(kind === 'text' ? 'Final' : 12)
    },
  )

  it('rejects an A/B fingerprint-target exchange even when global multisets match', async () => {
    const calls = [
      { ...textCall, id: 'call-a', input: { ...textCall.input, shape_id: 'shape-a', text: 'A' } },
      { ...textCall, id: 'call-b', input: { ...textCall.input, shape_id: 'shape-b', text: 'B' } },
    ]
    let applied = false
    const source = authority({
      readShape: vi.fn().mockImplementation((_slide, shapeId) =>
        Promise.resolve({
          slideId: 'slide-1',
          shapeId,
          text: applied ? (shapeId === 'shape-a' ? 'A' : 'B') : 'Before',
          left: 10,
          top: 20,
          width: 100,
          height: 30,
        }),
      ),
    })
    const subject = createOfficePowerPointVerification({
      authority: source,
      taskId: () => 'task-swap',
    })
    const enrolled = await subject.enroll(calls, undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    applied = true
    const bindingA = canonicalPowerPointVerificationBinding(calls[0]!, ['slide-1/shape-a'])
    const bindingB = canonicalPowerPointVerificationBinding(calls[1]!, ['slide-1/shape-b'])
    subject.recordProposal({
      id: 'a',
      toolName: 'edit_slide_text',
      fingerprint: 'state-a',
      targets: bindingA.targets,
      verificationBinding: { ...bindingA, fingerprint: bindingB.fingerprint },
    })
    subject.recordProposal({
      id: 'b',
      toolName: 'edit_slide_text',
      fingerprint: 'state-b',
      targets: bindingB.targets,
      verificationBinding: { ...bindingB, fingerprint: bindingA.fingerprint },
    })
    subject.recordSettlement({ id: 'a', status: 'confirmed' })
    subject.recordSettlement({ id: 'b', status: 'confirmed' })
    await expect(
      subject.complete({
        contract: enrolled.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode: 'verification_invalid' },
    })
  })

  it('skips a real supported no-op before creating a proposal or dispatching a mutation', async () => {
    const state = { text: 'After', left: 10 }
    const adapter = productionAdapter(state)
    const proposals = createStructuredProposalController()
    const verificationAuthority = authority({
      readShape: vi.fn().mockResolvedValue({
        slideId: 'slide-1',
        shapeId: 'shape-1',
        text: 'After',
        left: 10,
        top: 20,
        width: 100,
        height: 30,
      }),
    })
    const skill = createPowerPointSkill({ adapter, proposals, verificationAuthority })
    const enrolled = await skill.presentation!.enroll!([textCall], undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    await expect(skill.executeTool(textCall)).resolves.toMatchObject({
      mutated: false,
      output: '{"status":"unchanged"}',
    })
    expect(proposals.pending()).toBeUndefined()
    expect(adapter.editSlideText).not.toHaveBeenCalled()
  })

  it.each([
    [
      [
        { op: 'set_shape_text', slide_index: 0, shape_id: 'shape-1', text: 'After' },
        {
          op: 'add_text_box',
          slide_index: 0,
          name: 'New',
          text: 'x',
          left: 0,
          top: 0,
          width: 10,
          height: 10,
        },
      ],
    ],
    [
      [
        { op: 'set_shape_text', slide_index: 0, shape_id: 'shape-1', text: 'After' },
        { op: 'delete_shape', slide_index: 0, shape_id: 'shape-2' },
      ],
    ],
    [[{ op: 'duplicate_slide', slide_index: 0 }]],
  ])(
    'atomically bypasses verification for a real unsupported declarative program',
    async (operations) => {
      const adapter = productionAdapter({ text: 'Before', left: 5 })
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({
        adapter,
        proposals,
        verificationAuthority: authority(),
      })
      const call = {
        id: 'mixed-unsupported',
        name: 'execute_office_js',
        input: { program: { version: 1, operations } },
      }
      await expect(skill.presentation!.enroll!([call], undefined)).resolves.toEqual({
        kind: 'bypass',
      })
      await expect(skill.executeTool(call)).resolves.toMatchObject({ mutated: false })
      expect(proposals.pending()).toBeDefined()
    },
  )

  it('atomically bypasses a supported plus elevated multi-tool batch', async () => {
    const subject = createOfficePowerPointVerification({ authority: authority() })
    await expect(
      subject.enroll(
        [
          textCall,
          {
            id: 'master',
            name: 'edit_slide_master',
            input: { program: { version: 2, operations: [] } },
          },
        ],
        undefined,
      ),
    ).resolves.toEqual({ kind: 'bypass' })
  })

  it('collapses sequential properties to last-write-wins while retaining independent axes', async () => {
    const subject = createOfficePowerPointVerification({
      authority: authority(),
      taskId: () => 'task-last',
    })
    const enrolled = await subject.enroll(
      [
        { ...textCall, id: 'first', input: { ...textCall.input, text: 'After' } },
        { ...textCall, id: 'last', input: { ...textCall.input, text: 'Final' } },
        {
          id: 'geometry',
          name: 'execute_office_js',
          input: {
            program: {
              version: 1,
              operations: [
                {
                  op: 'set_shape_geometry',
                  slide_index: 0,
                  shape_id: 'shape-1',
                  left: 10,
                  top: 21,
                  width: 100,
                  height: 30,
                },
                {
                  op: 'set_shape_geometry',
                  slide_index: 0,
                  shape_id: 'shape-1',
                  left: 12,
                  top: 21,
                  width: 100,
                  height: 30,
                },
              ],
            },
          },
        },
      ],
      undefined,
    )
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    expect(enrolled.contract.checks).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ id: 'last-op0-text', expected: 'Final' }),
        expect.objectContaining({ id: 'geometry-op1-x', expected: 12 }),
        expect.objectContaining({ id: 'geometry-op1-y', expected: 21 }),
      ]),
    )
    expect(enrolled.contract.checks).toHaveLength(5)
    expect(subject.shouldSkip({ ...textCall, id: 'first' })).toBe(true)
    expect(
      subject.shouldSkip({ ...textCall, id: 'last', input: { ...textCall.input, text: 'Final' } }),
    ).toBe(false)
  })

  it('refuses verified status for the same targets with a substituted fingerprint', async () => {
    const subject = createOfficePowerPointVerification({
      authority: authority(),
      taskId: () => 'task-fingerprint',
    })
    const enrolled = await subject.enroll([textCall], undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    subject.recordProposal({
      id: 'replacement',
      toolName: 'edit_slide_text',
      fingerprint: 'different',
      targets: ['slide-1/shape-1'],
      verificationBinding: { ...verificationBinding(textCall), fingerprint: 'different' },
    })
    subject.recordSettlement({ id: 'replacement', status: 'confirmed' })
    await expect(
      subject.complete({
        contract: enrolled.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode: 'verification_invalid' },
    })
  })

  it('receives privacy-safe proposal lineage after post-dispatch confirmation', async () => {
    const controller = createStructuredProposalController()
    const events: unknown[] = []
    controller.subscribeAudit?.((event) => events.push(event))
    const proposal = controller.propose({
      operation: 'edit_slide_text',
      title: 'Edit',
      preview: { text: 'secret' },
      impact: { host: 'powerpoint', targets: ['slide-1/shape-1'], count: 1 },
      fingerprint: 'hash-1',
      validate: () => true,
      execute: () => undefined,
    })
    await controller.confirm(proposal.id)
    expect(events).toEqual([
      { kind: 'proposed', id: proposal.id, fingerprint: 'hash-1', targets: ['slide-1/shape-1'] },
      { kind: 'settled', id: proposal.id, status: 'confirmed' },
    ])
    expect(JSON.stringify(events)).not.toContain('secret')
  })

  it('leases authority, compiles exact native checks, and verifies with bounded convergence', async () => {
    const source = authority({
      readShape: vi
        .fn()
        .mockResolvedValueOnce({
          slideId: 'slide-1',
          shapeId: 'shape-1',
          text: 'Before',
          left: 9,
          top: 20,
          width: 100,
          height: 30,
        })
        .mockResolvedValue({
          slideId: 'slide-1',
          shapeId: 'shape-1',
          text: 'After',
          left: 10,
          top: 20,
          width: 100,
          height: 30,
        }),
    })
    const subject = createOfficePowerPointVerification({
      authority: source,
      taskId: () => 'task-1',
      delay: vi.fn(),
    })
    expect(await subject.prepare('edit')).toEqual({ kind: 'bypass' })
    const enrolled = await subject.enroll(
      [
        textCall,
        {
          id: 'c2',
          name: 'execute_office_js',
          input: {
            program: {
              version: 1,
              operations: [
                {
                  op: 'set_shape_geometry',
                  slide_index: 0,
                  shape_id: 'shape-1',
                  left: 10,
                  top: 20,
                  width: 100,
                  height: 30,
                },
              ],
            },
          },
        },
      ],
      undefined,
    )
    expect(enrolled).toMatchObject({
      kind: 'ready',
      contract: { documentToken: 'doc-1', sessionToken: 'session-1', affectedSlides: [1] },
    })
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    subject.recordProposal({
      id: 'proposal-1',
      toolName: 'edit_slide_text',
      fingerprint: 'hash',
      targets: ['slide-1/shape-1'],
      verificationBinding: verificationBinding(textCall),
    })
    subject.recordSettlement({ id: 'proposal-1', status: 'confirmed' })
    await expect(
      subject.complete({
        contract: enrolled.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: {
        status: 'verified',
        mutationReceiptIds: ['proposal-1'],
        passedCheckIds: ['c1-op0-text', 'c2-op0-x', 'c2-op0-y', 'c2-op0-width', 'c2-op0-height'],
      },
    })
    expect(source.readShape).toHaveBeenCalledTimes(10)
  })

  it('fails closed for stale authority, target expansion, session replacement, uncertainty, and cancellation truth', async () => {
    const subject = createOfficePowerPointVerification({
      authority: authority(),
      taskId: () => 'task-2',
      delay: vi.fn(),
    })
    const enrolled = await subject.enroll([textCall], undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    subject.recordProposal({
      id: 'proposal-2',
      toolName: 'edit_slide_text',
      fingerprint: 'hash',
      targets: ['other-target'],
      verificationBinding: verificationBinding(textCall, ['other-target']),
    })
    subject.recordSettlement({ id: 'proposal-2', status: 'confirmed' })
    await expect(
      subject.complete({
        contract: enrolled.contract,
        mutated: true,
        cancelled: true,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode: 'cancelled_after_apply' },
    })
  })

  it.each([
    ['target expansion', authority(), { targets: ['other-target'] }, 'verification_invalid'],
    [
      'session replacement',
      authority({
        current: vi.fn().mockResolvedValue({
          documentToken: 'doc-1',
          sessionToken: 'session-2',
          revision: `sha256:${'2'.repeat(64)}`,
        }),
      }),
      { targets: ['slide-1/shape-1'] },
      'stale_authority',
    ],
  ])('reports applied-unverified for %s', async (_name, source, proposalOverride, safeCode) => {
    const subject = createOfficePowerPointVerification({
      authority: source,
      taskId: () => 'task-x',
      delay: vi.fn(),
    })
    const enrolled = await subject.enroll([textCall], undefined)
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    subject.recordProposal({
      id: 'proposal-x',
      toolName: 'edit_slide_text',
      fingerprint: 'hash',
      ...proposalOverride,
      verificationBinding: verificationBinding(textCall, proposalOverride.targets),
    })
    subject.recordSettlement({ id: 'proposal-x', status: 'confirmed' })
    await expect(
      subject.complete({
        contract: enrolled.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode },
    })
  })

  it('distinguishes no-op, rejected stale proposal, and uncertain applied state', async () => {
    for (const [mutated, settlement, status, safeCode] of [
      [false, undefined, 'unchanged', undefined],
      [false, { id: 'p', status: 'rejected' as const }, 'needs_user', 'confirmation_required'],
      [
        false,
        { id: 'p', status: 'failed' as const, error: 'proposal_stale' },
        'failed',
        'stale_authority',
      ],
      [
        true,
        { id: 'p', status: 'failed' as const, error: 'office_state_uncertain' },
        'applied_unverified',
        'office_state_uncertain',
      ],
    ] as const) {
      const subject = createOfficePowerPointVerification({
        authority: authority(),
        taskId: () => 'task-y',
      })
      const enrolled = await subject.enroll([textCall], undefined)
      if (enrolled.kind !== 'ready') throw new Error('not ready')
      subject.recordProposal({
        id: 'p',
        toolName: 'edit_slide_text',
        fingerprint: 'hash',
        targets: ['slide-1/shape-1'],
        verificationBinding: verificationBinding(textCall),
      })
      if (settlement) subject.recordSettlement(settlement)
      await expect(
        subject.complete({
          contract: enrolled.contract,
          mutated,
          cancelled: false,
          correctionPasses: 0,
        }),
      ).resolves.toMatchObject({
        kind: 'receipt',
        receipt: { status, ...(safeCode ? { safeCode } : {}) },
      })
    }
  })

  it.each([
    ['edit_slide_master', { program: {} }],
    ['edit_slide_master_xml', { program: {} }],
    ['edit_slide_chart', { slide_index: 0, program: {} }],
    ['edit_slide_xml', { slide_index: 0, program: {} }],
  ])(
    'leaves unsupported elevated tool %s on its existing confirmation path',
    async (name, input) => {
      const subject = createOfficePowerPointVerification({
        authority: authority(),
        taskId: () => 'task-3',
        platform: 'Mac',
      })
      await expect(subject.enroll([{ id: 'elevated', name, input }], undefined)).resolves.toEqual({
        kind: 'bypass',
      })
    },
  )
})
