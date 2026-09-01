import type { OfficeDocumentClient } from '../office-document.js'
import type { OfficeDiagnostics } from '../diagnostics/office-diagnostics.js'

export type ProposalOperation = 'replace' | 'append'
export const MAX_PROPOSAL_SELECTION_LENGTH = 12_000
export const MAX_PROPOSAL_PREVIEW_BYTES = 64 * 1024

export interface ProposalImpact {
  host: string
  targets: string[]
  count: number
}

export interface StructuredProposal {
  id: string
  operation: string
  toolName?: string
  title: string
  preview: Readonly<Record<string, unknown>>
  impact: Readonly<ProposalImpact>
  fingerprint: string
  before?: unknown
  after?: unknown
  code?: string
}

export interface StructuredProposalRequest extends Omit<StructuredProposal, 'id'> {
  verificationBinding?: {
    callId: string
    fingerprint: string
    targets: string[]
  }
  validate(signal?: AbortSignal): boolean | Promise<boolean>
  execute(signal?: AbortSignal): void | Promise<void>
  verify?(signal?: AbortSignal): void | Promise<void>
}

export type ProposalDecision =
  | { status: 'confirmed' }
  | { status: 'applied_unverified'; historyId?: string }
  | { status: 'rejected' | 'cancelled' }
  | { status: 'failed'; error: string }

interface ProposalDecisionLifecycle {
  id: string
  promise: Promise<ProposalDecision>
  resolve(value: ProposalDecision): void
  settled: boolean
}

export interface StructuredProposalController {
  pending(): StructuredProposal | undefined
  subscribe(listener: () => void): () => void
  waitForDecision(id: string): Promise<ProposalDecision>
  propose(request: StructuredProposalRequest): StructuredProposal
  confirm(id: string): Promise<void>
  reject(): void
  newTurn(): void
  logout(): void
  subscribeAudit?(listener: (event: StructuredProposalAuditEvent) => void): () => void
}

export type StructuredProposalAuditEvent =
  | {
      kind: 'proposed'
      id: string
      toolName?: string
      fingerprint: string
      targets: string[]
      verificationBinding?: { callId: string; fingerprint: string; targets: string[] }
    }
  | { kind: 'settled'; id: string; status: ProposalDecision['status']; error?: string }

export interface OfficeProposal {
  id: string
  operation: ProposalOperation
  before: string
  value: string
  fingerprint: string
}

export interface ProposalController {
  pending(): OfficeProposal | undefined
  subscribe(listener: () => void): () => void
  waitForDecision(id: string): Promise<ProposalDecision>
  propose(operation: ProposalOperation, value: string): Promise<OfficeProposal>
  confirm(id: string): Promise<void>
  reject(): void
  newTurn(): void
  logout(): void
}

function invalidProposal(): never {
  throw new Error('invalid_tool_input')
}

const PROPOSAL_ERROR_CODES = new Set([
  'proposal_missing',
  'proposal_stale',
  'office_overwrite_required',
  'office_write_failed',
  'office_verify_failed',
  'office_recovery_failed',
  'office_concurrent_change',
  'office_state_uncertain',
  'office_applied_unverified',
])

function stableProposalError(error: unknown): string {
  const code = error instanceof Error ? error.message : ''
  return PROPOSAL_ERROR_CODES.has(code) || /^office_recovery_failed:word_[a-z_]+$/.test(code)
    ? code
    : 'office_write_failed'
}

function boundedCopy<T>(value: T): T {
  let serialized: string | undefined
  try {
    serialized = JSON.stringify(value)
  } catch {
    invalidProposal()
  }
  if (
    serialized === undefined ||
    new TextEncoder().encode(serialized).byteLength > MAX_PROPOSAL_PREVIEW_BYTES
  ) {
    invalidProposal()
  }
  return JSON.parse(serialized) as T
}

function deepFreeze<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    Object.freeze(value)
    for (const item of Object.values(value)) deepFreeze(item)
  }
  return value
}

export function createStructuredProposalController(
  diagnostics?: Pick<OfficeDiagnostics, 'setTool' | 'record'>,
): StructuredProposalController {
  const diagnose = (action: () => void) => {
    try {
      action()
    } catch {
      /* diagnostics never changes document behavior */
    }
  }
  let current:
    | {
        snapshot: StructuredProposal
        request: StructuredProposalRequest
        decision: ProposalDecisionLifecycle
      }
    | undefined
  let confirming: AbortController | undefined
  const listeners = new Set<() => void>()
  const auditListeners = new Set<(event: StructuredProposalAuditEvent) => void>()
  const snapshot = (value: StructuredProposal) => deepFreeze(boundedCopy(value))
  const publish = () => listeners.forEach((listener) => listener())
  const settle = (decision: ProposalDecisionLifecycle, value: ProposalDecision) => {
    if (decision.settled) return
    decision.settled = true
    decision.resolve(value)
    auditListeners.forEach((listener) => listener({ kind: 'settled', id: decision.id, ...value }))
  }
  const invalidate = (status: 'rejected' | 'cancelled') => {
    if (current) settle(current.decision, { status })
    current = undefined
    confirming?.abort()
    publish()
  }
  return {
    pending: () => (current ? snapshot(current.snapshot) : undefined),
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    waitForDecision(id) {
      if (!current || current.snapshot.id !== id)
        return Promise.reject(new Error('proposal_missing'))
      return current.decision.promise
    },
    propose(request) {
      if (confirming) throw new Error('proposal_confirmation_in_progress')
      if (
        !request.operation ||
        request.operation.length > 128 ||
        !request.title ||
        request.title.length > 512 ||
        !request.fingerprint ||
        request.fingerprint.length > 512 ||
        (request.toolName !== undefined && (!request.toolName || request.toolName.length > 128)) ||
        (request.code !== undefined && request.code.length > MAX_PROPOSAL_PREVIEW_BYTES) ||
        !request.impact ||
        !request.impact.host ||
        request.impact.host.length > 32 ||
        !Number.isSafeInteger(request.impact.count) ||
        request.impact.count < 0 ||
        request.impact.targets.length > 256 ||
        request.impact.targets.some((target) => !target || target.length > 512) ||
        (request.verificationBinding !== undefined &&
          (!request.verificationBinding.callId ||
            request.verificationBinding.callId.length > 128 ||
            !request.verificationBinding.fingerprint ||
            request.verificationBinding.fingerprint.length > 128 ||
            request.verificationBinding.targets.length > 256 ||
            request.verificationBinding.targets.some((target) => !target || target.length > 512)))
      )
        invalidProposal()
      const publicValue = snapshot({
        id: crypto.randomUUID(),
        operation: request.operation,
        toolName: request.toolName,
        title: request.title,
        preview: request.preview,
        impact: request.impact,
        fingerprint: request.fingerprint,
        before: request.before,
        after: request.after,
        code: request.code,
      })
      diagnose(() => diagnostics?.setTool(request.toolName ?? request.operation))
      if (current) settle(current.decision, { status: 'cancelled' })
      let resolve!: (value: ProposalDecision) => void
      const promise = new Promise<ProposalDecision>((next) => {
        resolve = next
      })
      current = {
        snapshot: publicValue,
        request: { ...request },
        decision: { id: publicValue.id, promise, resolve, settled: false },
      }
      auditListeners.forEach((listener) =>
        listener({
          kind: 'proposed',
          id: publicValue.id,
          ...(publicValue.toolName ? { toolName: publicValue.toolName } : {}),
          fingerprint: publicValue.fingerprint,
          targets: [...publicValue.impact.targets],
          ...(request.verificationBinding
            ? {
                verificationBinding: {
                  ...request.verificationBinding,
                  targets: [...request.verificationBinding.targets],
                },
              }
            : {}),
        }),
      )
      publish()
      return snapshot(publicValue)
    },
    async confirm(id) {
      if (confirming) throw new Error('proposal_confirmation_in_progress')
      const proposal = current
      if (!proposal || proposal.snapshot.id !== id) throw new Error('proposal_missing')
      current = undefined
      publish()
      const controller = new AbortController()
      confirming = controller
      let phase: 'validate' | 'write' | 'verify' = 'validate'
      let phaseStartedAt = Date.now()
      try {
        if (!(await proposal.request.validate(controller.signal)) || controller.signal.aborted) {
          throw new Error('proposal_stale')
        }
        phase = 'write'
        phaseStartedAt = Date.now()
        await proposal.request.execute(controller.signal)
        phase = 'verify'
        phaseStartedAt = Date.now()
        // Once execute resolves, cancellation cannot truthfully imply that the write was not
        // applied. Always reconcile/verify; use an un-aborted signal when Stop raced the commit.
        await proposal.request.verify?.(controller.signal.aborted ? undefined : controller.signal)
        settle(proposal.decision, { status: 'confirmed' })
      } catch (error) {
        const code = stableProposalError(error)
        diagnose(() =>
          diagnostics?.record({
            phase: code.startsWith('office_recovery_failed') ? 'recovery' : phase,
            errorCode: code,
            error,
            durationMs: Math.max(0, Date.now() - phaseStartedAt),
          }),
        )
        if (code === 'office_applied_unverified') {
          settle(proposal.decision, { status: 'applied_unverified' })
        } else {
          settle(proposal.decision, { status: 'failed', error: code })
          throw error
        }
      } finally {
        if (confirming === controller) confirming = undefined
      }
    },
    reject: () => invalidate('rejected'),
    newTurn: () => invalidate('cancelled'),
    logout: () => invalidate('cancelled'),
    subscribeAudit(listener) {
      auditListeners.add(listener)
      return () => auditListeners.delete(listener)
    },
  }
}

export function selectionFingerprint(value: string): string {
  let hash = 0x811c9dc5
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index)
    hash = Math.imul(hash, 0x01000193)
  }
  return `${value.length}:${(hash >>> 0).toString(16).padStart(8, '0')}`
}

export function createProposalController(
  document: OfficeDocumentClient,
  diagnostics?: Pick<OfficeDiagnostics, 'setTool' | 'record'>,
): ProposalController {
  const structured = createStructuredProposalController(diagnostics)
  const legacy = (proposal: StructuredProposal): OfficeProposal => {
    const preview = proposal.preview as { value: string }
    return Object.freeze({
      ...proposal,
      operation: proposal.operation as ProposalOperation,
      before: proposal.before as string,
      value: preview.value,
    })
  }
  return {
    pending: () => {
      const proposal = structured.pending()
      return proposal ? legacy(proposal) : undefined
    },
    subscribe: structured.subscribe,
    waitForDecision: structured.waitForDecision,
    async propose(operation, value) {
      const before = await document.readSelection()
      if (before.length > MAX_PROPOSAL_SELECTION_LENGTH) throw new Error('selection_too_large')
      const fingerprint = selectionFingerprint(before)
      const expected = operation === 'replace' ? value : `${before}${value}`
      return legacy(
        structured.propose({
          operation,
          toolName: operation === 'replace' ? 'propose_replace_selection' : 'propose_append_text',
          title: operation === 'replace' ? 'Replace selection' : 'Append to selection',
          preview: { value },
          impact: { host: 'office', targets: ['selection'], count: 1 },
          fingerprint,
          before,
          after: expected,
          validate: async (signal) => {
            if (signal?.aborted) return false
            const current = await document.readSelection()
            return (
              !signal?.aborted &&
              current === before &&
              selectionFingerprint(current) === fingerprint
            )
          },
          execute: async (signal) => {
            if (signal?.aborted) throw new Error('proposal_stale')
            await (operation === 'replace'
              ? document.replaceSelection(value)
              : document.appendText(before, value))
            // Office callbacks cannot be revoked after dispatch. Always re-read before reporting
            // completion, including when logout/cancellation happened during the callback.
            const current = await document.readSelection()
            if (signal?.aborted) throw new Error('proposal_stale')
            if (current !== expected) throw new Error('office_verify_failed')
          },
        }),
      )
    },
    confirm: structured.confirm,
    reject: structured.reject,
    newTurn: structured.newTurn,
    logout: structured.logout,
  }
}
