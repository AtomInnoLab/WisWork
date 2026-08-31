import { describe, expect, it, vi } from 'vitest'
import type { AgentToolCall } from '../src/shared/ipc'
import {
  compileCanonicalSlidesCalls,
  createSlidesTaskController,
  type SlidesTaskEnrollment,
} from '../src/renderer/ai/task-controller'
import { PRESENTATION_CONSISTENCY_GOLDEN } from '@wiswork/presentation-verification'
import type { SlidesTaskReviewAdapter } from '../src/renderer/ai/task-review'
import { createSlidesSkill, type DeckAccess } from '../src/renderer/ai/slides-skill'

const contract = {
  version: 1 as const,
  taskId: 'task-1',
  documentToken: 'doc-1',
  sessionToken: 'session-1',
  baseRevision: `sha256:${'a'.repeat(64)}`,
  affectedSlides: [1],
  referenceSlides: [],
  checks: [
    {
      id: 'check-001',
      kind: 'element_property' as const,
      slide: 1,
      roleOrTarget: { kind: 'target' as const, targetToken: 'target-1' },
      property: 'color' as const,
      expected: '#123456',
    },
  ],
  maxCorrectionPasses: 2 as const,
}

const call: AgentToolCall = {
  id: 'call-1',
  name: 'set_element_style',
  input: { slideIndex: 0, sourceId: 'runtime-1', color: '#123456' },
}

function adapter(overrides: Partial<SlidesTaskReviewAdapter> = {}): SlidesTaskReviewAdapter {
  const authority = {
    documentToken: 'doc-1',
    sessionToken: 'session-1',
    revision: `sha256:${'b'.repeat(64)}`,
    leaseToken: 'lease-1',
    baseRevision: contract.baseRevision,
    mutationReceiptIds: ['tx-1'],
  }
  return {
    refresh: async (lineage) => ({ ...authority, ...lineage }),
    verifyDeterministic: async () => [{ checkId: 'check-001', status: 'pass' }],
    capture: async ({ slide, role }) => ({
      slide,
      role,
      mediaToken: 'media-1',
      bytes: 100,
      revision: authority.revision,
      leaseToken: authority.leaseToken,
      sessionToken: authority.sessionToken,
    }),
    review: async () => ({ status: 'pass', failedCheckIds: [], observations: [], fixIntents: [] }),
    correct: async () => ({ mutationReceiptId: 'correction-1', correctedCheckIds: ['check-001'] }),
    isCurrent: () => true,
    ...overrides,
  }
}

function enrollment(): SlidesTaskEnrollment {
  return { contract, plannedMutationTargets: ['target-1'] }
}

describe('Slides verified task controller', () => {
  it('executes the shared pages 6-8 golden through production compiler, controller and render reviewer', async () => {
    const calls = PRESENTATION_CONSISTENCY_GOLDEN.operations.map((operation, index) => ({
      id: `golden-${index}`,
      name: operation.property === 'color' ? 'set_element_style' : 'set_element_transform',
      input: {
        slideIndex: operation.slide - 1,
        sourceId: operation.targetToken.split(':')[1],
        ...(operation.property === 'color'
          ? { color: operation.value }
          : {
              [operation.property === 'width'
                ? 'w'
                : operation.property === 'height'
                  ? 'h'
                  : operation.property]: operation.value,
            }),
      },
    })) as AgentToolCall[]
    const sourceTargetTokens = Object.fromEntries(
      PRESENTATION_CONSISTENCY_GOLDEN.pages.flatMap((page) =>
        page.shapes.map((shape) => [`${page.slide}:${shape.role}`, shape.targetToken]),
      ),
    )
    const authority = {
      documentToken: PRESENTATION_CONSISTENCY_GOLDEN.documentToken,
      sessionToken: PRESENTATION_CONSISTENCY_GOLDEN.sessionToken,
      revision: PRESENTATION_CONSISTENCY_GOLDEN.baseRevision,
      slides: PRESENTATION_CONSISTENCY_GOLDEN.pages.map((page) => ({
        number: page.slide,
        slideToken: `slide-${page.slide}`,
        elements: page.shapes.map((shape) => ({
          targetToken: shape.targetToken,
          role: shape.role as 'title' | 'body' | 'emphasis',
          locked: false,
          properties: { color: shape.color, ...(shape.x === undefined ? {} : { x: shape.x }) },
        })),
      })),
    }
    const compiled = compileCanonicalSlidesCalls({
      calls,
      authority,
      sourceTargetTokens,
      taskId: 'golden-pc',
    })
    if ('kind' in compiled) throw new Error('golden did not compile')
    expect(compiled.plan).toEqual([
      'presentation_apply_bounded_edits',
      'presentation_verify_postconditions',
    ])
    const outside = JSON.stringify(authority.slides[4])
    const renderReview = adapter({
      refresh: vi.fn().mockImplementation(async (lineage) => ({
        documentToken: authority.documentToken,
        sessionToken: authority.sessionToken,
        revision: `sha256:${'2'.repeat(64)}`,
        leaseToken: 'golden-lease',
        ...lineage,
      })),
      verifyDeterministic: vi
        .fn()
        .mockImplementation(async (contract: typeof compiled.contract) =>
          contract.checks.map(({ id }) => ({ checkId: id, status: 'pass' as const })),
        ),
      capture: vi.fn().mockImplementation(async ({ slide, role, authority }) => ({
        slide,
        role,
        mediaToken: `golden-${slide}`,
        bytes: 100,
        revision: authority.revision,
        leaseToken: authority.leaseToken,
        sessionToken: authority.sessionToken,
      })),
      review: vi.fn().mockResolvedValue({
        status: 'pass',
        failedCheckIds: [],
        observations: [],
        fixIntents: [],
      }),
    })
    const telemetry = vi.fn()
    const controller = createSlidesTaskController({
      enroll: async () => compiled,
      reviewAdapter: renderReview,
      telemetry,
    })
    const ready = await controller.hooks.enroll!(calls, undefined)
    if (ready.kind !== 'ready') throw new Error('not ready')
    for (const [index, target] of [...new Set(compiled.plannedMutationTargets)].entries())
      controller.recordMutation({
        transactionId: `golden-tx-${index}`,
        status: 'applied',
        mutatedTargetTokens: [target],
      })
    await expect(
      controller.hooks.complete({
        contract: ready.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'verified', passedCheckIds: { length: 17 } },
    })
    expect(renderReview.capture).toHaveBeenCalled()
    expect(telemetry).toHaveBeenCalledWith({
      host: 'pc',
      phase: 'complete',
      outcome: 'success',
      code: 'verified',
      count: 3,
      durationMs: expect.any(Number),
    })
    expect(JSON.stringify(telemetry.mock.calls)).not.toMatch(/slide-|golden-|title|body|emphasis/)
    expect(JSON.stringify(authority.slides[4])).toBe(outside)
  })
  it('mounts set_element_style through enrollment, one canonical transaction, screenshot review, and receipt', async () => {
    const execute = vi.fn(async (request: { transactionId: string }) => ({
      receipt: {
        status: 'applied' as const,
        transactionId: request.transactionId,
        beforeRevision: `sha256:${'a'.repeat(64)}`,
        afterRevision: `sha256:${'b'.repeat(64)}`,
        operationCount: 1,
        mutatedTargets: ['target-1'],
      },
      authoritativeState: 'fresh' as const,
    }))
    const slide = {
      widthPx: 1280,
      heightPx: 720,
      scale: 1,
      nodes: [
        {
          id: 'runtime-1',
          sourceId: 'runtime-1',
          type: 'text',
          box: {
            x: 10,
            y: 10,
            w: 200,
            h: 80,
            rotationDeg: 0,
            flipH: false,
            flipV: false,
            centerX: 110,
            centerY: 50,
          },
          fill: { kind: 'none' },
          text: {
            lines: [
              {
                runs: [
                  {
                    text: 'Title',
                    x: 0,
                    baselineY: 20,
                    fontFamily: 'Arial',
                    fontSizePx: 16,
                    color: '#000000',
                    bold: false,
                    italic: false,
                    underline: false,
                    widthPx: 50,
                  },
                ],
                top: 0,
                height: 20,
              },
            ],
            insets: { l: 0, t: 0, r: 0, b: 0 },
            anchor: 'top',
            fontScale: 1,
            wrap: true,
            contentHeight: 20,
          },
        },
      ],
    }
    const access: DeckAccess = {
      getSlides: () => [slide] as never,
      getCurrent: () => 0,
      getSelectedIds: () => [],
      getAcceptanceAuthorityLease: async () => ({
        documentToken: 'doc-1',
        sessionToken: 'session-1',
        revision: contract.baseRevision,
        slideCount: 1,
        leaseToken: 'lease-1',
      }),
      inspectAcceptanceAuthority: async () =>
        ({
          documentToken: 'doc-1',
          sessionToken: 'session-1',
          revision: contract.baseRevision,
          leaseToken: 'lease-1',
          sourceTargetTokens: { '1:runtime-1': 'target-1' },
          slides: [
            {
              number: 1,
              slideToken: 'slide-1',
              elements: [
                {
                  targetToken: 'target-1',
                  locked: false,
                  properties: { color: '#000000', x: 10, y: 10, width: 200, height: 80 },
                },
              ],
            },
          ],
        }) as never,
      applySlide: () => {},
      applyDeck: () => {},
      fitWidthPx: 1280,
      executePresentationOperation: execute as never,
      taskReviewAdapter: adapter(),
    }
    const skill = createSlidesSkill(access)
    const ready = await skill.presentation!.enroll!([call], undefined)
    expect(ready.kind).toBe('ready')
    if (ready.kind !== 'ready') throw new Error('expected ready')
    const execution = await skill.executeTool(call)
    expect(execution).toMatchObject({ mutated: true })
    expect(execute).toHaveBeenCalledOnce()
    await expect(
      skill.presentation!.complete({
        contract: ready.contract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({ kind: 'receipt', receipt: { status: 'verified' } })
  })

  it('enrolls one exact canonical call and closes its one applied transaction as verified', async () => {
    const enroll = vi.fn(async () => enrollment())
    const review = adapter()
    const controller = createSlidesTaskController({ enroll, reviewAdapter: review })

    await expect(controller.hooks.enroll?.([call], undefined)).resolves.toMatchObject({
      kind: 'ready',
      contract,
    })
    controller.recordMutation({
      transactionId: 'tx-1',
      status: 'applied',
      mutatedTargetTokens: ['target-1'],
      rollbackId: 'rollback-1',
    })
    await expect(
      controller.hooks.complete({ contract, mutated: true, cancelled: false, correctionPasses: 0 }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'verified', mutationReceiptIds: ['tx-1'], rollbackId: 'rollback-1' },
    })
    expect(enroll).toHaveBeenCalledWith([call], undefined, undefined)
  })

  it('reports applied_unverified on screenshot failure and never reexecutes the original mutation', async () => {
    const executeOriginal = vi.fn()
    const controller = createSlidesTaskController({
      enroll: async () => enrollment(),
      reviewAdapter: adapter({ capture: async () => null }),
      executeOriginal,
    })
    await controller.hooks.enroll?.([call], undefined)
    controller.recordMutation({
      transactionId: 'tx-1',
      status: 'applied',
      mutatedTargetTokens: ['target-1'],
    })
    const result = await controller.hooks.complete({
      contract,
      mutated: true,
      cancelled: false,
      correctionPasses: 0,
    })
    expect(result).toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode: 'screenshot_unavailable' },
    })
    expect(executeOriginal).not.toHaveBeenCalled()
  })

  it('does not execute an unsafe visual fix', async () => {
    const correct = vi.fn()
    const controller = createSlidesTaskController({
      enroll: async () => enrollment(),
      reviewAdapter: adapter({
        review: async () => ({
          status: 'needs_fix',
          failedCheckIds: ['check-001'],
          observations: [],
          fixIntents: [
            {
              checkId: 'check-001',
              kind: 'set_property',
              roleOrTarget: { kind: 'target', targetToken: 'other-target' },
              property: 'color',
              value: '#123456',
            },
          ],
        }),
        correct,
      }),
    })
    await controller.hooks.enroll?.([call], undefined)
    controller.recordMutation({
      transactionId: 'tx-1',
      status: 'applied',
      mutatedTargetTokens: ['target-1'],
    })
    await expect(
      controller.hooks.complete({ contract, mutated: true, cancelled: false, correctionPasses: 0 }),
    ).resolves.toMatchObject({ kind: 'receipt', receipt: { status: 'needs_user' } })
    expect(correct).not.toHaveBeenCalled()
  })

  it('reconciles cancellation/session replacement and preserves task-review-before-QC ordering', async () => {
    const order: string[] = []
    const controller = createSlidesTaskController({
      enroll: async () => enrollment(),
      reviewAdapter: adapter({
        refresh: async () => ({
          documentToken: 'doc-1',
          sessionToken: 'replacement',
          revision: `sha256:${'b'.repeat(64)}`,
          leaseToken: 'lease-2',
          baseRevision: contract.baseRevision,
          mutationReceiptIds: ['tx-1'],
        }),
      }),
      afterTaskReview: () => {
        order.push('generic-qc')
      },
    })
    await controller.hooks.enroll?.([call], undefined)
    controller.recordMutation({
      transactionId: 'tx-1',
      status: 'applied',
      mutatedTargetTokens: ['target-1'],
    })
    const result = await controller.hooks.complete({
      contract,
      mutated: true,
      cancelled: true,
      correctionPasses: 0,
    })
    order.unshift('task-review')
    expect(result).toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode: 'stale_authority' },
    })
    expect(order).toEqual(['task-review', 'generic-qc'])
  })

  it('isolates a deferred abandoned run from the newly enrolled run lifecycle', async () => {
    const newerContract = { ...contract, taskId: 'task-2' }
    let enrollmentCount = 0
    let resolveOldReview!: (value: Awaited<ReturnType<SlidesTaskReviewAdapter['review']>>) => void
    let reviewCount = 0
    const reviewAdapter = adapter({
      review: async () => {
        reviewCount += 1
        if (reviewCount === 1)
          return new Promise((resolve) => {
            resolveOldReview = resolve
          })
        return { status: 'pass', failedCheckIds: [], observations: [], fixIntents: [] }
      },
    })
    const afterTaskReview = vi.fn()
    const controller = createSlidesTaskController({
      enroll: async () => ({
        contract: enrollmentCount++ === 0 ? contract : newerContract,
        plannedMutationTargets: ['target-1'],
      }),
      reviewAdapter,
      afterTaskReview,
    })

    await controller.hooks.enroll?.([call], undefined)
    controller.recordMutation({
      transactionId: 'tx-old',
      status: 'applied',
      mutatedTargetTokens: ['target-1'],
    })
    const oldCompletion = controller.hooks.complete({
      contract,
      mutated: true,
      cancelled: false,
      correctionPasses: 0,
    })
    await vi.waitFor(() => expect(resolveOldReview).toBeTypeOf('function'))

    controller.reset()
    await controller.hooks.enroll?.([call], undefined)
    controller.recordMutation({
      transactionId: 'tx-new',
      status: 'applied',
      mutatedTargetTokens: ['target-1'],
    })
    resolveOldReview({ status: 'pass', failedCheckIds: [], observations: [], fixIntents: [] })
    await expect(oldCompletion).resolves.toMatchObject({ kind: 'receipt' })
    expect(afterTaskReview).not.toHaveBeenCalled()

    await expect(
      controller.hooks.complete({
        contract: newerContract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'verified', mutationReceiptIds: ['tx-new'] },
    })
    expect(afterTaskReview).toHaveBeenCalledOnce()
  })

  it('skips every refresh side effect when an old run is abandoned before review starts', async () => {
    const newerContract = { ...contract, taskId: 'task-2' }
    let enrollmentCount = 0
    const baseAdapter = adapter()
    const refresh = vi.fn(baseAdapter.refresh)
    const afterTaskReview = vi.fn()
    const controller = createSlidesTaskController({
      enroll: async () => ({
        contract: enrollmentCount++ === 0 ? contract : newerContract,
        plannedMutationTargets: ['target-1'],
      }),
      reviewAdapter: { ...baseAdapter, refresh },
      afterTaskReview,
    })
    await controller.hooks.enroll?.([call], undefined)
    controller.recordMutation({
      transactionId: 'tx-old',
      status: 'applied',
      mutatedTargetTokens: ['target-1'],
    })
    controller.reset()
    await controller.hooks.enroll?.([call], undefined)
    controller.recordMutation({
      transactionId: 'tx-new',
      status: 'applied',
      mutatedTargetTokens: ['target-1'],
    })

    await expect(
      controller.hooks.complete({
        contract,
        mutated: true,
        cancelled: true,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'applied_unverified', safeCode: 'cancelled_after_apply' },
    })
    expect(refresh).not.toHaveBeenCalled()
    expect(afterTaskReview).not.toHaveBeenCalled()

    await expect(
      controller.hooks.complete({
        contract: newerContract,
        mutated: true,
        cancelled: false,
        correctionPasses: 0,
      }),
    ).resolves.toMatchObject({
      kind: 'receipt',
      receipt: { status: 'verified', mutationReceiptIds: ['tx-new'] },
    })
    expect(refresh).toHaveBeenCalledOnce()
    expect(afterTaskReview).toHaveBeenCalledOnce()
  })
})
