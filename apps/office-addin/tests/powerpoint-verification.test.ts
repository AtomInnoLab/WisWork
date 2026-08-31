import { describe, expect, it, vi } from 'vitest'
import { createOfficePowerPointVerification } from '../src/skills/powerpoint/powerpoint-verification.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'

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

describe('Office PowerPoint presentation verification', () => {
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
      fingerprint: 'proposal-hash',
      targets: ['slide-1/shape-1'],
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
        passedCheckIds: ['c1-text', 'c2-x', 'c2-y', 'c2-width', 'c2-height'],
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
    subject.recordProposal({ id: 'proposal-2', fingerprint: 'hash', targets: ['other-target'] })
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
    subject.recordProposal({ id: 'proposal-x', fingerprint: 'hash', ...proposalOverride })
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
      subject.recordProposal({ id: 'p', fingerprint: 'hash', targets: ['slide-1/shape-1'] })
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

  it('classifies style, solid background, master/chart/XML and Mac gates without autocorrection', async () => {
    const subject = createOfficePowerPointVerification({
      authority: authority(),
      taskId: () => 'task-3',
      platform: 'Mac',
    })
    const enrolled = await subject.enroll(
      [
        {
          id: 'style',
          name: 'set_shape_style',
          input: { slide_index: 0, shape_id: 'shape-1', color: '#112233' },
        },
        { id: 'bg', name: 'set_slide_background', input: { slide_index: 0, color: '#FFFFFF' } },
        { id: 'master', name: 'edit_slide_master', input: { program: {} } },
        { id: 'xml', name: 'edit_slide_master_xml', input: { program: {} } },
        { id: 'chart', name: 'edit_slide_chart', input: { slide_index: 0, program: {} } },
      ],
      undefined,
    )
    expect(enrolled).toMatchObject({
      kind: 'ready',
      requiresConfirmation: true,
      contract: { maxCorrectionPasses: 0 },
    })
    if (enrolled.kind !== 'ready') throw new Error('not ready')
    expect(enrolled.contract.checks).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ id: 'style-color', property: 'color' }),
        expect.objectContaining({ id: 'bg-background', property: 'background_color' }),
      ]),
    )
  })
})
