import { describe, expect, it, vi } from 'vitest'
import {
  canonicalPowerPointVerificationBinding,
  createOfficePowerPointVerification,
} from '../src/skills/powerpoint/powerpoint-verification.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import type { PowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter.js'

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
  ...overrides,
})

const textCall = {
  id: 'c1',
  name: 'edit_slide_text',
  input: { slide_index: 0, shape_id: 'shape-1', text: 'After' },
}
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
