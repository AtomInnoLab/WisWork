import type {
  AgentToolCall,
  PresentationTaskCompletion,
  PresentationTaskHooks,
  PresentationTaskPreparation,
} from '@wiswork/agent-core'
import {
  parsePresentationAcceptanceContract,
  type PresentationAcceptanceCheck,
  type PresentationAcceptanceContract,
  type PresentationCompletionReceipt,
  type SupportedProperty,
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
}

export type SafeProposalRecord = { id: string; fingerprint: string; targets: string[] }
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
}

async function sha256(value: string): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(value))
  return `sha256:${Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')}`
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
  }
}

const elevated = new Set([
  'edit_slide_master',
  'edit_slide_master_xml',
  'edit_slide_chart',
  'edit_slide_xml',
])
const geometry = ['left', 'top', 'width', 'height'] as const
const propertyMap: Record<string, SupportedProperty> = {
  color: 'color',
  fill_color: 'fill_color',
  stroke_color: 'stroke_color',
  font_size: 'font_size',
  font_family: 'font_family',
  bold: 'bold',
  italic: 'italic',
}

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
      },
    ]
  if (
    call.name === 'set_shape_style' &&
    Number.isSafeInteger(input.slide_index) &&
    typeof input.shape_id === 'string'
  )
    return Object.entries(propertyMap).flatMap(([key, property]) =>
      typeof input[key] === 'string' ||
      typeof input[key] === 'number' ||
      typeof input[key] === 'boolean'
        ? [
            {
              slide: (input.slide_index as number) + 1,
              target: input.shape_id as string,
              property,
              expected: input[key] as string | number | boolean,
            },
          ]
        : [],
    )
  if (
    call.name === 'set_slide_background' &&
    Number.isSafeInteger(input.slide_index) &&
    typeof input.color === 'string'
  )
    return [
      {
        slide: (input.slide_index as number) + 1,
        property: 'background_color',
        expected: input.color,
      },
    ]
  if (call.name !== 'execute_office_js') return []
  const result: Array<{
    slide: number
    target?: string
    property: SupportedProperty
    expected: string | number | boolean
  }> = []
  for (const op of program(input)) {
    if (!Number.isSafeInteger(op.slide_index) || typeof op.shape_id !== 'string') continue
    const slide = (op.slide_index as number) + 1
    if (op.op === 'set_shape_text' && typeof op.text === 'string')
      result.push({ slide, target: op.shape_id, property: 'text', expected: op.text })
    if (op.op === 'set_shape_geometry')
      for (const key of geometry)
        if (typeof op[key] === 'number')
          result.push({
            slide,
            target: op.shape_id,
            property: ({ left: 'x', top: 'y', width: 'width', height: 'height' } as const)[key],
            expected: op[key] as number,
          })
  }
  return result
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
}): OfficePowerPointVerificationHooks {
  let enrolled:
    | {
        contract: PresentationAcceptanceContract
        targets: Set<string>
        locations: Map<string, { slideId: string; shapeId?: string }>
        elevated: boolean
      }
    | undefined
  let proposal: SafeProposalRecord | undefined
  let settlement: SafeProposalSettlement | undefined
  const delay = options.delay ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)))

  const enroll = async (calls: readonly AgentToolCall[]): Promise<PresentationTaskPreparation> => {
    const operations = calls.flatMap((call) =>
      callOperations(call).map((operation) => ({ ...operation, callId: call.id })),
    )
    const requiresConfirmation = calls.some((call) => elevated.has(call.name))
    if (operations.length === 0 && !requiresConfirmation) return { kind: 'bypass' }
    const lease = await options.authority.acquire()
    const targets = new Set<string>()
    const locations = new Map<string, { slideId: string; shapeId?: string }>()
    const checks: PresentationAcceptanceCheck[] = []
    for (const [index, operation] of operations.entries()) {
      let targetToken: string | undefined
      if (operation.target) {
        const state = await options.authority.readShape(operation.slide - 1, operation.target)
        targets.add(`${state.slideId}/${state.shapeId}`)
        targets.add(`${operation.slide - 1}/${state.shapeId}`)
        targetToken = `target-${index}`
        locations.set(targetToken, { slideId: state.slideId, shapeId: state.shapeId })
      } else {
        const state = await options.authority.readSlide(operation.slide - 1)
        targets.add(`${state.slideId}/background`)
        targetToken = `target-${index}`
        locations.set(targetToken, { slideId: state.slideId })
      }
      checks.push({
        id: `${operation.callId || `check-${index}`}-${operation.property === 'background_color' ? 'background' : operation.property}`,
        kind: 'element_property',
        slide: operation.slide,
        roleOrTarget: { kind: 'target', targetToken },
        property: operation.property,
        expected: operation.expected,
      })
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
      maxCorrectionPasses: requiresConfirmation ? 0 : 2,
    })
    enrolled = { contract, targets, locations, elevated: requiresConfirmation }
    proposal = undefined
    settlement = undefined
    return {
      kind: 'ready',
      contract,
      requiresConfirmation,
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
    enroll: (calls) => enroll(calls),
    recordProposal(value) {
      proposal = { ...value, targets: [...value.targets] }
    },
    recordSettlement(value) {
      settlement = { ...value }
    },
    abandon() {
      /* reconciliation deliberately retains safe state */
    },
    async complete(context): Promise<PresentationTaskCompletion> {
      const contract = context.contract
      const allIds = contract.checks.map((check) => check.id)
      if (!context.mutated) {
        if (settlement?.status === 'failed')
          return receipt(contract, {
            status: 'failed',
            mutationReceiptIds: [],
            passedCheckIds: [],
            failedCheckIds: allIds,
            unavailableCheckIds: [],
            safeCode: settlement.error === 'proposal_stale' ? 'stale_authority' : 'mutation_failed',
          })
        return receipt(contract, {
          status: settlement?.status === 'rejected' ? 'needs_user' : 'unchanged',
          mutationReceiptIds: [],
          passedCheckIds: [],
          failedCheckIds: [],
          unavailableCheckIds: allIds,
          ...(settlement?.status === 'rejected'
            ? { safeCode: 'confirmation_required' as const }
            : {}),
        })
      }
      if (!proposal || settlement?.status !== 'confirmed' || settlement.id !== proposal.id)
        return receipt(contract, {
          status: 'applied_unverified',
          mutationReceiptIds: proposal ? [proposal.id] : [],
          passedCheckIds: [],
          failedCheckIds: [],
          unavailableCheckIds: allIds,
          safeCode: 'office_state_uncertain',
        })
      if (context.cancelled)
        return receipt(contract, {
          status: 'applied_unverified',
          mutationReceiptIds: [proposal.id],
          passedCheckIds: [],
          failedCheckIds: [],
          unavailableCheckIds: allIds,
          safeCode: 'cancelled_after_apply',
        })
      if (enrolled?.elevated)
        return receipt(contract, {
          status: 'applied_unverified',
          mutationReceiptIds: [proposal.id],
          passedCheckIds: [],
          failedCheckIds: [],
          unavailableCheckIds: allIds,
          safeCode: 'unsupported_check',
        })
      const current = await options.authority.current(context.signal)
      if (
        current.documentToken !== contract.documentToken ||
        current.sessionToken !== contract.sessionToken
      )
        return receipt(contract, {
          status: 'applied_unverified',
          mutationReceiptIds: [proposal.id],
          passedCheckIds: [],
          failedCheckIds: [],
          unavailableCheckIds: allIds,
          safeCode: 'stale_authority',
        })
      if ([...proposal.targets].some((target) => !enrolled?.targets.has(target)))
        return receipt(contract, {
          status: 'applied_unverified',
          mutationReceiptIds: [proposal.id],
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
      return receipt(contract, {
        status,
        mutationReceiptIds: [proposal.id],
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
