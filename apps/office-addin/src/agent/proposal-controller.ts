import type { OfficeDocumentClient } from '../office-document.js'

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
  validate(signal?: AbortSignal): boolean | Promise<boolean>
  execute(signal?: AbortSignal): void | Promise<void>
  verify?(signal?: AbortSignal): void | Promise<void>
}

export interface StructuredProposalController {
  pending(): StructuredProposal | undefined
  propose(request: StructuredProposalRequest): StructuredProposal
  confirm(id: string): Promise<void>
  reject(): void
  newTurn(): void
  logout(): void
}

export interface OfficeProposal {
  id: string
  operation: ProposalOperation
  before: string
  value: string
  fingerprint: string
}

export interface ProposalController {
  pending(): OfficeProposal | undefined
  propose(operation: ProposalOperation, value: string): Promise<OfficeProposal>
  confirm(id: string): Promise<void>
  reject(): void
  newTurn(): void
  logout(): void
}

function invalidProposal(): never {
  throw new Error('invalid_tool_input')
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

export function createStructuredProposalController(): StructuredProposalController {
  let current: { snapshot: StructuredProposal; request: StructuredProposalRequest } | undefined
  let confirming: AbortController | undefined
  const snapshot = (value: StructuredProposal) => deepFreeze(boundedCopy(value))
  const invalidate = () => {
    current = undefined
    confirming?.abort()
  }
  return {
    pending: () => (current ? snapshot(current.snapshot) : undefined),
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
        request.impact.targets.some((target) => !target || target.length > 512)
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
      current = { snapshot: publicValue, request: { ...request } }
      return snapshot(publicValue)
    },
    async confirm(id) {
      if (confirming) throw new Error('proposal_confirmation_in_progress')
      const proposal = current
      if (!proposal || proposal.snapshot.id !== id) throw new Error('proposal_missing')
      current = undefined
      const controller = new AbortController()
      confirming = controller
      try {
        if (!(await proposal.request.validate(controller.signal)) || controller.signal.aborted) {
          throw new Error('proposal_stale')
        }
        await proposal.request.execute(controller.signal)
        if (controller.signal.aborted) throw new Error('proposal_stale')
        await proposal.request.verify?.(controller.signal)
      } finally {
        if (confirming === controller) confirming = undefined
      }
    },
    reject: invalidate,
    newTurn: invalidate,
    logout: invalidate,
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

export function createProposalController(document: OfficeDocumentClient): ProposalController {
  const structured = createStructuredProposalController()
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
