import { useEffect, useRef } from 'react'
import {
  type EnhancedMutationConfirmationKey,
  type Lang,
} from '@wiswork/i18n'

const MAX_ID = 160
const MAX_TOOL = 96

const OPERATIONS = ['insert', 'replace', 'delete', 'format', 'restructure', 'compile'] as const
const TARGETS = [
  'document',
  'selection',
  'blocks',
  'cells',
  'sheet',
  'slides',
  'elements',
  'project-files',
] as const
const SCOPES = ['single', 'selection', 'bounded-set', 'whole-document'] as const

export type EnhancedMutationOperation = (typeof OPERATIONS)[number]
export type EnhancedMutationTarget = (typeof TARGETS)[number]
export type EnhancedMutationScope = (typeof SCOPES)[number]

export interface EnhancedMutationSummary {
  readonly operation: EnhancedMutationOperation
  readonly target: EnhancedMutationTarget
  readonly scope: EnhancedMutationScope
  readonly count?: number
}

export interface EnhancedMutationProposal {
  readonly proposalId: string
  readonly documentId: string
  readonly generation: number
  readonly toolName: string
  readonly summary: EnhancedMutationSummary
  readonly expiresAt: number
}

export interface EnhancedMutationProposalApi {
  onProposal(listener: (proposal: EnhancedMutationProposal) => void): () => void
  confirmProposal(documentId: string, generation: number, proposalId: string): Promise<void>
  cancelProposal(documentId: string, generation: number, proposalId: string): Promise<void>
}

function validText(value: unknown, maximum: number): value is string {
  return typeof value === 'string' && value.length > 0 && value.length <= maximum
}

function hasValidEnvelope(value: unknown): value is Omit<EnhancedMutationProposal, 'summary'> & {
  summary: unknown
} {
  if (!value || typeof value !== 'object') return false
  const proposal = value as Record<string, unknown>
  return (
    Object.keys(proposal).every((key) =>
      ['proposalId', 'documentId', 'generation', 'toolName', 'summary', 'expiresAt'].includes(key),
    ) &&
    validText(proposal.proposalId, MAX_ID) &&
    validText(proposal.documentId, MAX_ID) &&
    Number.isSafeInteger(proposal.generation) &&
    (proposal.generation as number) >= 0 &&
    validText(proposal.toolName, MAX_TOOL) &&
    Number.isSafeInteger(proposal.expiresAt) &&
    (proposal.expiresAt as number) > 0
  )
}

function isEnhancedMutationSummary(value: unknown): value is EnhancedMutationSummary {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const summary = value as Record<string, unknown>
  return (
    Object.keys(summary).every((key) => ['operation', 'target', 'scope', 'count'].includes(key)) &&
    OPERATIONS.includes(summary.operation as EnhancedMutationOperation) &&
    TARGETS.includes(summary.target as EnhancedMutationTarget) &&
    SCOPES.includes(summary.scope as EnhancedMutationScope) &&
    (summary.count === undefined ||
      (Number.isSafeInteger(summary.count) &&
        (summary.count as number) >= 1 &&
        (summary.count as number) <= 10_000))
  )
}

export function isEnhancedMutationProposal(value: unknown): value is EnhancedMutationProposal {
  return hasValidEnvelope(value) && isEnhancedMutationSummary(value.summary)
}

export type EnhancedMutationConfirmationTranslator = (
  key: EnhancedMutationConfirmationKey,
) => string

export interface EnhancedMutationConfirmationProps {
  readonly api?: EnhancedMutationProposalApi
  /** The host must supply its resolved UI locale; document content language is not authority. */
  readonly locale: Lang
  readonly translate?: EnhancedMutationConfirmationTranslator
}

/**
 * Automatically submits valid bounded proposals to the privileged owner. The owner still checks
 * the exact document, generation, catalog digest and mutation capability before executing it.
 */
export function EnhancedMutationConfirmation({ api }: EnhancedMutationConfirmationProps) {
  const consumedRef = useRef(new Set<string>())

  useEffect(() => {
    if (!api) return
    const unsubscribe = api.onProposal((candidate) => {
      if (!hasValidEnvelope(candidate)) return
      const key = `${candidate.documentId}\0${candidate.generation}\0${candidate.proposalId}`
      if (consumedRef.current.has(key)) return
      consumedRef.current.add(key)
      if (!isEnhancedMutationSummary(candidate.summary)) {
        void api
          .cancelProposal(candidate.documentId, candidate.generation, candidate.proposalId)
          .catch(() => {})
        return
      }
      if (candidate.expiresAt <= Date.now()) {
        void api
          .cancelProposal(candidate.documentId, candidate.generation, candidate.proposalId)
          .catch(() => {})
        return
      }
      void api
        .confirmProposal(candidate.documentId, candidate.generation, candidate.proposalId)
        .catch(() => {})
    })
    return unsubscribe
  }, [api])

  return null
}
