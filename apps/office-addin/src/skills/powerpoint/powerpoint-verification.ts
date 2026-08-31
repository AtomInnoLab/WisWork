import type {
  AgentToolCall,
  PresentationTaskCompletion,
  PresentationTaskHooks,
  PresentationTaskPreparation,
} from '@wiswork/agent-core'
import {
  digestPresentationAcceptanceContract,
  parsePresentationAcceptanceContract,
  parsePresentationRenderingFacts,
  parseVisualReviewResult,
  type PresentationAcceptanceCheck,
  type PresentationAcceptanceContract,
  type PresentationCompletionReceipt,
  type SupportedProperty,
  type VisualReviewResult,
} from '@wiswork/presentation-verification'
import type { PowerPointAdapter } from './browser-powerpoint-adapter.js'

export interface OfficePowerPointAuthorityLease {
  documentToken: string
  sessionToken: string
  revision: string
}

export interface OfficePowerPointShapeState {
  slideId: string
  shapeId: string
  text?: string
  color?: string
  fillColor?: string
  strokeColor?: string
  fontSize?: number
  fontFamily?: string
  bold?: boolean
  italic?: boolean
  left: number
  top: number
  width: number
  height: number
}

export interface OfficePowerPointVerificationAuthority {
  acquire(signal?: AbortSignal): Promise<OfficePowerPointAuthorityLease>
  current(signal?: AbortSignal): Promise<OfficePowerPointAuthorityLease>
  readShape(
    slideIndex: number,
    shapeId: string,
    signal?: AbortSignal,
  ): Promise<OfficePowerPointShapeState>
  readSlide(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ slideId: string; backgroundColor?: string }>
  captureScreenshot(
    slideIndex: number,
    signal?: AbortSignal,
  ): Promise<{ base64: string; mime: 'image/png' }>
}

export interface OfficePowerPointVisualReviewer {
  review(request: {
    facts: ReturnType<typeof parsePresentationRenderingFacts>
    images: Array<{ mediaToken: string; base64: string; mime: 'image/png' }>
    isolation: { tools: []; maxTurns: 1 }
    signal?: AbortSignal
  }): Promise<VisualReviewResult>
}

export type SafeProposalRecord = {
  id: string
  toolName?: string
  fingerprint: string
  targets: string[]
  verificationBinding?: { callId: string; fingerprint: string; targets: string[] }
}
export type SafeProposalSettlement = {
  id: string
  status: 'confirmed' | 'rejected' | 'cancelled' | 'failed'
  error?: string
}

export interface OfficePowerPointVerificationHooks extends PresentationTaskHooks {
  enroll(
    calls: readonly AgentToolCall[],
    currentContract?: PresentationAcceptanceContract,
    signal?: AbortSignal,
  ): PresentationTaskPreparation | Promise<PresentationTaskPreparation>
  recordProposal(proposal: SafeProposalRecord): void
  recordSettlement(settlement: SafeProposalSettlement): void
  shouldSkip(call: AgentToolCall): boolean
  setReviewer(reviewer: OfficePowerPointVisualReviewer): void
}

async function sha256(value: string): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(value))
  return `sha256:${Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')}`
}

export function powerPointProposalFingerprint(value: string): string {
  let result = 0x811c9dc5
  for (let index = 0; index < value.length; index += 1) {
    result ^= value.charCodeAt(index)
    result = Math.imul(result, 0x01000193)
  }
  return `${value.length}:${(result >>> 0).toString(16).padStart(8, '0')}`
}

export function createBrowserPowerPointVerificationAuthority(
  adapter: PowerPointAdapter,
): OfficePowerPointVerificationAuthority {
  const sessionToken = `office-${crypto.randomUUID()}`
  const lease = async (signal?: AbortSignal): Promise<OfficePowerPointAuthorityLease> => {
    const state = await adapter.verifySlides(signal)
    let documentSeed = 'unsaved'
    try {
      documentSeed = String(Office.context.document.url || 'unsaved')
    } catch {
      /* bounded fallback */
    }
    return {
      documentToken: `doc-${(await sha256(documentSeed)).slice(7, 39)}`,
      sessionToken,
      revision: await sha256(JSON.stringify(state)),
    }
  }
  return {
    acquire: lease,
    current: lease,
    async readShape(slideIndex, shapeId, signal) {
      const [shapes, text] = await Promise.all([
        adapter.listSlideShapes(slideIndex, signal),
        adapter.readSlideText(slideIndex, shapeId, signal).catch(() => undefined),
      ])
      const shape = shapes.shapes.find((item) => item.id === shapeId)
      if (!shape) throw new Error('office_read_failed')
      return { slideId: shapes.slideId, shapeId, ...(text ? { text: text.text } : {}), ...shape }
    },
    async readSlide(slideIndex, signal) {
      const state = await adapter.listSlideShapes(slideIndex, signal)
      return { slideId: state.slideId }
    },
    captureScreenshot: (slideIndex, signal) => adapter.screenshotSlide(slideIndex, signal),
  }
}

const geometry = ['left', 'top', 'width', 'height'] as const
function program(input: unknown): Array<Record<string, unknown>> {
  if (!input || typeof input !== 'object') return []
  let value = (input as Record<string, unknown>).program
  if (typeof value === 'string') {
    try {
      value = JSON.parse(value)
    } catch {
      return []
    }
  }
  if (
    !value ||
    typeof value !== 'object' ||
    !Array.isArray((value as Record<string, unknown>).operations)
  )
    return []
  return (value as { operations: unknown[] }).operations.filter(
    (item): item is Record<string, unknown> =>
      Boolean(item && typeof item === 'object' && !Array.isArray(item)),
  )
}

function callOperations(call: AgentToolCall): Array<{
  slide: number
  target?: string
  property: SupportedProperty
  expected: string | number | boolean
  opIndex: number
}> {
  const input = call.input as Record<string, unknown>
  if (
    call.name === 'edit_slide_text' &&
    Number.isSafeInteger(input.slide_index) &&
    typeof input.shape_id === 'string' &&
    typeof input.text === 'string'
  )
    return [
      {
        slide: (input.slide_index as number) + 1,
        target: input.shape_id,
        property: 'text',
        expected: input.text,
        opIndex: 0,
      },
    ]
  if (call.name !== 'execute_office_js') return []
  const result: Array<{
    slide: number
    target?: string
    property: SupportedProperty
    expected: string | number | boolean
    opIndex: number
  }> = []
  for (const [opIndex, op] of program(input).entries()) {
    if (!Number.isSafeInteger(op.slide_index) || typeof op.shape_id !== 'string') continue
    const slide = (op.slide_index as number) + 1
    if (op.op === 'set_shape_text' && typeof op.text === 'string')
      result.push({ slide, target: op.shape_id, property: 'text', expected: op.text, opIndex })
    if (op.op === 'set_shape_geometry')
      for (const key of geometry)
        if (typeof op[key] === 'number')
          result.push({
            slide,
            target: op.shape_id,
            property: ({ left: 'x', top: 'y', width: 'width', height: 'height' } as const)[key],
            expected: op[key] as number,
            opIndex,
          })
  }
  return result
}

export function canonicalPowerPointVerificationBinding(
  call: AgentToolCall,
  normalizedTargets: readonly string[],
): { callId: string; fingerprint: string; targets: string[] } {
  const callId = call.invocationId ?? call.id
  const targets = [...normalizedTargets].sort()
  return {
    callId,
    fingerprint: powerPointProposalFingerprint(
      JSON.stringify({
        callId,
        toolName: call.name,
        operations: callOperations(call),
        targets,
      }),
    ),
    targets,
  }
}

const equal = (property: SupportedProperty, actual: unknown, expected: unknown) =>
  typeof expected === 'number' && typeof actual === 'number'
    ? Math.abs(actual - expected) <= (['x', 'y', 'width', 'height'].includes(property) ? 0.01 : 0)
    : typeof expected === 'string' && property.includes('color')
      ? String(actual).toUpperCase() === expected.toUpperCase()
      : actual === expected

const actualProperty = (shape: OfficePowerPointShapeState, property: SupportedProperty): unknown =>
  (
    ({
      text: shape.text,
      color: shape.color,
      fill_color: shape.fillColor,
      stroke_color: shape.strokeColor,
      font_size: shape.fontSize,
      font_family: shape.fontFamily,
      bold: shape.bold,
      italic: shape.italic,
      x: shape.left,
      y: shape.top,
      width: shape.width,
      height: shape.height,
    }) as Partial<Record<SupportedProperty, unknown>>
  )[property]

export function createOfficePowerPointVerification(options: {
  authority: OfficePowerPointVerificationAuthority
  taskId?: () => string
  platform?: string
  delay?: (milliseconds: number) => Promise<void>
  reviewer?: OfficePowerPointVisualReviewer
}): OfficePowerPointVerificationHooks {
  let reviewer = options.reviewer
  let enrolled:
    | {
        contract: PresentationAcceptanceContract
        targets: Set<string>
        locations: Map<string, { slideId: string; shapeId?: string }>
        plannedProposals: Array<{
          callId: string
          toolName: string
          fingerprint: string
          targets: string[]
        }>
        skipCallIds: Set<string>
      }
    | undefined
  let proposals: SafeProposalRecord[] = []
  let settlements: SafeProposalSettlement[] = []
  let activeReviewAbort: AbortController | undefined
  const delay = options.delay ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)))

  const enroll = async (
    calls: readonly AgentToolCall[],
    currentContract?: PresentationAcceptanceContract,
    signal?: AbortSignal,
  ): Promise<PresentationTaskPreparation> => {
    const fullySupported = (call: AgentToolCall) =>
      call.name === 'edit_slide_text' ||
      (call.name === 'execute_office_js' &&
        program(call.input).length > 0 &&
        program(call.input).every(
          (op) => op.op === 'set_shape_text' || op.op === 'set_shape_geometry',
        ))
    const mutationTools = new Set([
      'edit_slide_text',
      'execute_office_js',
      'duplicate_slide',
      'edit_slide_master',
      'edit_slide_master_xml',
      'edit_slide_chart',
      'edit_slide_xml',
    ])
    const hasUnsupportedMutation = calls.some(
      (call) => mutationTools.has(call.name) && !fullySupported(call),
    )
    if (currentContract && hasUnsupportedMutation)
      return { kind: 'clarify', question: 'presentation_scope_required' }
    if (hasUnsupportedMutation) return { kind: 'bypass' }
    const rawOperations = calls.flatMap((call) =>
      callOperations(call).map((operation) => ({
        ...operation,
        callId: call.invocationId ?? call.id,
        sourceCallId: call.id,
      })),
    )
    if (rawOperations.length === 0)
      return currentContract
        ? { kind: 'clarify', question: 'presentation_scope_required' }
        : { kind: 'bypass' }
    if (currentContract) {
      if (!enrolled || currentContract.taskId !== enrolled.contract.taskId)
        return { kind: 'clarify', question: 'presentation_scope_required' }
      const correctionTargets = new Map<string, string[]>()
      for (const operation of rawOperations) {
        const allowed = currentContract.checks.find((check) => {
          if (
            check.kind !== 'element_property' ||
            check.roleOrTarget.kind !== 'target' ||
            check.slide !== operation.slide ||
            check.property !== operation.property ||
            check.expected !== operation.expected
          )
            return false
          return (
            enrolled?.locations.get(check.roleOrTarget.targetToken)?.shapeId === operation.target
          )
        })
        if (!allowed) return { kind: 'clarify', question: 'presentation_scope_required' }
      }
      for (const call of calls) {
        const seen = new Set<number>()
        for (const operation of callOperations(call)) {
          if (!operation.target || seen.has(operation.opIndex)) continue
          seen.add(operation.opIndex)
          const state = await options.authority.readShape(
            operation.slide - 1,
            operation.target,
            signal,
          )
          const targets = correctionTargets.get(call.id) ?? []
          targets.push(`${state.slideId}/${state.shapeId}`)
          correctionTargets.set(call.id, targets)
        }
      }
      enrolled.plannedProposals.push(
        ...calls.map((call) => ({
          toolName: call.name,
          ...canonicalPowerPointVerificationBinding(call, correctionTargets.get(call.id) ?? []),
        })),
      )
      return { kind: 'ready', contract: currentContract, requiresConfirmation: false }
    }
    const collapsed = new Map<string, (typeof rawOperations)[number]>()
    for (const operation of rawOperations)
      collapsed.set(
        `${operation.slide}:${operation.target ?? 'background'}:${operation.property}`,
        operation,
      )
    const operations = [...collapsed.values()]
    const lease = await options.authority.acquire()
    const targets = new Set<string>()
    const locations = new Map<string, { slideId: string; shapeId?: string }>()
    const checks: PresentationAcceptanceCheck[] = []
    const callMatches = new Map<string, boolean[]>()
    const shapeStates = new Map<string, OfficePowerPointShapeState>()
    for (const [index, operation] of operations.entries()) {
      let targetToken: string | undefined
      if (operation.target) {
        const state = await options.authority.readShape(operation.slide - 1, operation.target)
        shapeStates.set(`${operation.slide}:${operation.target}`, state)
        targets.add(`${state.slideId}/${state.shapeId}`)
        targetToken = `target-${index}`
        locations.set(targetToken, { slideId: state.slideId, shapeId: state.shapeId })
        const values = callMatches.get(operation.sourceCallId) ?? []
        values.push(
          equal(operation.property, actualProperty(state, operation.property), operation.expected),
        )
        callMatches.set(operation.sourceCallId, values)
      } else {
        const state = await options.authority.readSlide(operation.slide - 1)
        targets.add(`${state.slideId}/background`)
        targetToken = `target-${index}`
        locations.set(targetToken, { slideId: state.slideId })
      }
      checks.push({
        id: `${operation.callId || `check-${index}`}-op${operation.opIndex}-${operation.property}`,
        kind: 'element_property',
        slide: operation.slide,
        roleOrTarget: { kind: 'target', targetToken },
        property: operation.property,
        expected: operation.expected,
      })
    }
    const targetsByCall = new Map<string, string[]>()
    const targetOperations = new Set<string>()
    for (const call of calls) {
      const matches: boolean[] = []
      for (const operation of callOperations(call)) {
        if (!operation.target) continue
        const key = `${operation.slide}:${operation.target}`
        const state =
          shapeStates.get(key) ??
          (await options.authority.readShape(operation.slide - 1, operation.target))
        shapeStates.set(key, state)
        const targetOperation = `${call.id}:${operation.opIndex}`
        if (!targetOperations.has(targetOperation)) {
          targetOperations.add(targetOperation)
          const callTargets = targetsByCall.get(call.id) ?? []
          callTargets.push(`${state.slideId}/${state.shapeId}`)
          targetsByCall.set(call.id, callTargets)
        }
        matches.push(
          equal(operation.property, actualProperty(state, operation.property), operation.expected),
        )
      }
      if (!callMatches.has(call.id) && matches.length) callMatches.set(call.id, matches)
    }
    if (checks.length === 0)
      checks.push({
        id: 'elevated-unavailable',
        kind: 'render_quality',
        slide: 1,
        rules: ['legible'],
      })
    const contract = parsePresentationAcceptanceContract({
      version: 1,
      taskId: options.taskId?.() ?? `task-${Date.now()}`,
      documentToken: lease.documentToken,
      sessionToken: lease.sessionToken,
      baseRevision: lease.revision,
      affectedSlides: [
        ...new Set(operations.map((item) => item.slide).concat(checks.length ? [] : [1])),
      ],
      referenceSlides: [],
      checks,
      maxCorrectionPasses: 2,
    })
    const skipCallIds = new Set(
      calls
        .filter((call) => fullySupported(call) && callMatches.get(call.id)?.every(Boolean))
        .map((call) => call.id),
    )
    const plannedCalls = calls.filter(
      (call) => callOperations(call).length && !skipCallIds.has(call.id),
    )
    const plannedProposals = plannedCalls.map((call) => ({
      toolName: call.name,
      ...canonicalPowerPointVerificationBinding(call, targetsByCall.get(call.id) ?? []),
    }))
    enrolled = {
      contract,
      targets,
      locations,
      plannedProposals,
      skipCallIds,
    }
    proposals = []
    settlements = []
    return {
      kind: 'ready',
      contract,
      requiresConfirmation: false,
      ...(calls.length > 1
        ? { plan: ['Apply approved PowerPoint edits', 'Verify exact post-write properties'] }
        : {}),
    }
  }

  const receipt = (
    contract: PresentationAcceptanceContract,
    partial: Omit<
      PresentationCompletionReceipt,
      'version' | 'taskId' | 'affectedSlides' | 'correctionPasses'
    > & { correctionPasses?: number },
  ): PresentationTaskCompletion => ({
    kind: 'receipt',
    receipt: {
      version: 1,
      taskId: contract.taskId,
      affectedSlides: contract.affectedSlides,
      correctionPasses: partial.correctionPasses ?? 0,
      ...partial,
    },
  })

  return {
    prepare: () => ({ kind: 'bypass' }),
    enroll: (calls, currentContract, signal) => enroll(calls, currentContract, signal),
    recordProposal(value) {
      if (proposals.length < 50)
        proposals.push({
          ...value,
          targets: [...value.targets],
          ...(value.verificationBinding
            ? {
                verificationBinding: {
                  ...value.verificationBinding,
                  targets: [...value.verificationBinding.targets],
                },
              }
            : {}),
        })
    },
    recordSettlement(value) {
      if (settlements.length < 50) settlements.push({ ...value })
    },
    shouldSkip(call) {
      return enrolled?.skipCallIds.has(call.id) ?? false
    },
    setReviewer(value) {
      reviewer = value
    },
    abandon() {
      activeReviewAbort?.abort()
    },
    async complete(context): Promise<PresentationTaskCompletion> {
      const contract = context.contract
      const allIds = contract.checks.map((check) => check.id)
      const confirmed = proposals.filter((item) =>
        settlements.some(
          (settlement) => settlement.id === item.id && settlement.status === 'confirmed',
        ),
      )
      const failedSettlement = settlements.find((item) => item.status === 'failed')
      const rejected = settlements.some((item) => item.status === 'rejected')
      if (!context.mutated) {
        if (failedSettlement)
          return receipt(contract, {
            status: 'failed',
            mutationReceiptIds: [],
            passedCheckIds: [],
            failedCheckIds: allIds,
            unavailableCheckIds: [],
            safeCode:
              failedSettlement.error === 'proposal_stale' ? 'stale_authority' : 'mutation_failed',
          })
        return receipt(contract, {
          status: rejected ? 'needs_user' : 'unchanged',
          mutationReceiptIds: [],
          passedCheckIds: [],
          failedCheckIds: [],
          unavailableCheckIds: allIds,
          ...(rejected ? { safeCode: 'confirmation_required' as const } : {}),
        })
      }
      if (confirmed.length === 0 || confirmed.length !== proposals.length)
        return receipt(contract, {
          status: 'applied_unverified',
          mutationReceiptIds: confirmed.map((item) => item.id),
          passedCheckIds: [],
          failedCheckIds: [],
          unavailableCheckIds: allIds,
          safeCode: 'office_state_uncertain',
        })
      if (context.cancelled)
        return receipt(contract, {
          status: 'applied_unverified',
          mutationReceiptIds: confirmed.map((item) => item.id),
          passedCheckIds: [],
          failedCheckIds: [],
          unavailableCheckIds: allIds,
          safeCode: 'cancelled_after_apply',
        })
      const current = await options.authority.current(context.signal)
      if (
        current.documentToken !== contract.documentToken ||
        current.sessionToken !== contract.sessionToken
      )
        return receipt(contract, {
          status: 'applied_unverified',
          mutationReceiptIds: confirmed.map((item) => item.id),
          passedCheckIds: [],
          failedCheckIds: [],
          unavailableCheckIds: allIds,
          safeCode: 'stale_authority',
        })
      const actualTargets = new Set(confirmed.flatMap((item) => item.targets))
      const enrollment = enrolled
      const actualBindings = confirmed
        .map((item) =>
          JSON.stringify({
            callId: item.verificationBinding?.callId ?? '',
            toolName: item.toolName ?? '',
            fingerprint: item.verificationBinding?.fingerprint ?? '',
            targets: [...(item.verificationBinding?.targets ?? [])].sort(),
          }),
        )
        .sort()
      const plannedBindings = (enrollment?.plannedProposals ?? [])
        .map((item) =>
          JSON.stringify({
            callId: item.callId,
            toolName: item.toolName,
            fingerprint: item.fingerprint,
            targets: [...item.targets].sort(),
          }),
        )
        .sort()
      const eachProposalTargetsBound = confirmed.every(
        (item) =>
          JSON.stringify([...item.targets].sort()) ===
          JSON.stringify([...(item.verificationBinding?.targets ?? [])].sort()),
      )
      const proposalBindingsValid =
        enrollment !== undefined &&
        eachProposalTargetsBound &&
        JSON.stringify(actualBindings) === JSON.stringify(plannedBindings)
      const targetCoverageValid =
        enrollment !== undefined &&
        actualTargets.size === enrollment.targets.size &&
        [...actualTargets].every((target) => enrollment.targets.has(target))
      if (!proposalBindingsValid || !targetCoverageValid)
        return receipt(contract, {
          status: 'applied_unverified',
          mutationReceiptIds: confirmed.map((item) => item.id),
          passedCheckIds: [],
          failedCheckIds: allIds,
          unavailableCheckIds: [],
          safeCode: 'verification_invalid',
        })
      const passed: string[] = []
      const failed: string[] = []
      const unavailable: string[] = []
      for (const check of contract.checks) {
        if (check.kind !== 'element_property' || check.roleOrTarget.kind !== 'target') {
          unavailable.push(check.id)
          continue
        }
        const location = enrolled?.locations.get(check.roleOrTarget.targetToken)
        if (!location) {
          unavailable.push(check.id)
          continue
        }
        const { slideId, shapeId } = location
        let matched = false
        try {
          for (let attempt = 0; attempt < 3; attempt++) {
            if (check.property === 'background_color') {
              const state = await options.authority.readSlide(check.slide - 1, context.signal)
              matched =
                state.slideId === slideId &&
                equal(check.property, state.backgroundColor, check.expected)
            } else {
              const state = await options.authority.readShape(
                check.slide - 1,
                shapeId!,
                context.signal,
              )
              matched =
                state.slideId === slideId &&
                state.shapeId === shapeId &&
                equal(check.property, actualProperty(state, check.property), check.expected)
            }
            if (matched) break
            if (attempt < 2) await delay(25)
          }
          ;(matched ? passed : failed).push(check.id)
        } catch {
          unavailable.push(check.id)
        }
      }
      const status = unavailable.length
        ? 'applied_unverified'
        : failed.length
          ? 'needs_user'
          : 'verified'
      if (status === 'verified' && reviewer) {
        const reviewAbort = new AbortController()
        activeReviewAbort = reviewAbort
        const abortReview = () => reviewAbort.abort()
        context.signal?.addEventListener('abort', abortReview, { once: true })
        try {
          const beforeCapture = await options.authority.current(reviewAbort.signal)
          if (
            beforeCapture.documentToken !== contract.documentToken ||
            beforeCapture.sessionToken !== contract.sessionToken
          )
            throw new Error('stale_authority')
          const slides = [...new Set([...contract.affectedSlides, ...contract.referenceSlides])]
          if (slides.length > 8) throw new Error('screenshot_unavailable')
          const images = await Promise.all(
            slides.map(async (slide) => {
              const image = await options.authority.captureScreenshot(slide - 1, reviewAbort.signal)
              if (image.mime !== 'image/png') throw new Error('screenshot_unavailable')
              return {
                slide,
                role: contract.affectedSlides.includes(slide)
                  ? ('affected' as const)
                  : ('reference' as const),
                mediaToken: `shot-${powerPointProposalFingerprint(image.base64).replace(':', '-')}`,
                bytes: Math.floor((image.base64.length * 3) / 4),
                base64: image.base64,
                mime: image.mime,
              }
            }),
          )
          const afterCapture = await options.authority.current(reviewAbort.signal)
          if (
            afterCapture.documentToken !== beforeCapture.documentToken ||
            afterCapture.sessionToken !== beforeCapture.sessionToken ||
            afterCapture.revision !== beforeCapture.revision
          )
            throw new Error('stale_authority')
          const facts = parsePresentationRenderingFacts({
            contractDigest: await digestPresentationAcceptanceContract(contract),
            revision: afterCapture.revision,
            screenshots: images.map(({ slide, role, mediaToken, bytes }) => ({
              slide,
              role,
              mediaToken,
              bytes: Math.max(1, bytes),
            })),
            deterministicResults: contract.checks.map((check) => ({
              checkId: check.id,
              status: passed.includes(check.id) ? 'pass' : 'unavailable',
              ...(passed.includes(check.id) ? {} : { code: 'unsupported_check' as const }),
            })),
          })
          let review: VisualReviewResult
          try {
            review = parseVisualReviewResult(
              await reviewer.review({
                facts,
                images: images.map(({ mediaToken, base64, mime }) => ({
                  mediaToken,
                  base64,
                  mime,
                })),
                isolation: { tools: [], maxTurns: 1 },
                signal: reviewAbort.signal,
              }),
            )
          } catch (error) {
            throw new Error('verification_invalid', { cause: error })
          }
          const failedIds = new Set(review.failedCheckIds)
          const fixIds = new Set(review.fixIntents.map((intent) => intent.checkId))
          const reviewBindingValid =
            review.failedCheckIds.every((id) => allIds.includes(id)) &&
            (review.status !== 'needs_fix' ||
              (failedIds.size === fixIds.size && [...failedIds].every((id) => fixIds.has(id))))
          if (!reviewBindingValid)
            return receipt(contract, {
              status: 'applied_unverified',
              mutationReceiptIds: confirmed.map((item) => item.id),
              passedCheckIds: [],
              failedCheckIds: [],
              unavailableCheckIds: allIds,
              correctionPasses: context.correctionPasses,
              safeCode: 'verification_invalid',
            })
          if (review.status === 'cannot_verify')
            return receipt(contract, {
              status: 'applied_unverified',
              mutationReceiptIds: confirmed.map((item) => item.id),
              passedCheckIds: [],
              failedCheckIds: [],
              unavailableCheckIds: allIds,
              correctionPasses: context.correctionPasses,
              safeCode: 'review_unavailable',
            })
          if (review.status === 'needs_fix') {
            const safe = review.fixIntents.every((intent) => {
              const check = contract.checks.find((item) => item.id === intent.checkId)
              return (
                check?.kind === 'element_property' &&
                check.roleOrTarget.kind === 'target' &&
                ['text', 'x', 'y', 'width', 'height'].includes(intent.property) &&
                intent.property === check.property &&
                JSON.stringify(intent.roleOrTarget) === JSON.stringify(check.roleOrTarget) &&
                intent.value === check.expected
              )
            })
            if (safe && context.correctionPasses < contract.maxCorrectionPasses)
              return {
                kind: 'correct',
                instruction:
                  'Apply only the failed approved PowerPoint acceptance checks using their already approved values; do not expand targets or use package, master, chart, table, or XML tools.',
              }
            return receipt(contract, {
              status: 'needs_user',
              mutationReceiptIds: confirmed.map((item) => item.id),
              passedCheckIds: passed.filter((id) => !review.failedCheckIds.includes(id)),
              failedCheckIds: review.failedCheckIds,
              unavailableCheckIds: [],
              correctionPasses: context.correctionPasses,
              safeCode: safe ? 'verification_invalid' : 'confirmation_required',
            })
          }
        } catch (error) {
          return receipt(contract, {
            status: 'applied_unverified',
            mutationReceiptIds: confirmed.map((item) => item.id),
            passedCheckIds: [],
            failedCheckIds: [],
            unavailableCheckIds: allIds,
            correctionPasses: context.correctionPasses,
            safeCode:
              error instanceof Error && error.message === 'stale_authority'
                ? 'stale_authority'
                : error instanceof Error && error.message === 'verification_invalid'
                  ? 'verification_invalid'
                  : 'screenshot_unavailable',
          })
        } finally {
          context.signal?.removeEventListener('abort', abortReview)
          if (activeReviewAbort === reviewAbort) activeReviewAbort = undefined
        }
      }
      return receipt(contract, {
        status,
        mutationReceiptIds: confirmed.map((item) => item.id),
        passedCheckIds: passed,
        failedCheckIds: failed,
        unavailableCheckIds: unavailable,
        correctionPasses: context.correctionPasses,
        ...(unavailable.length
          ? { safeCode: 'office_state_uncertain' as const }
          : failed.length
            ? { safeCode: 'verification_invalid' as const }
            : {}),
      })
    },
  }
}
