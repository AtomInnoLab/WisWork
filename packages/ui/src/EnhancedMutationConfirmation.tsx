import { useCallback, useEffect, useRef, useState } from 'react'

const MAX_ID = 160
const MAX_TOOL = 96
const MAX_SUMMARY = 480

export interface EnhancedMutationProposal {
  readonly proposalId: string
  readonly documentId: string
  readonly generation: number
  readonly toolName: string
  readonly summary: string
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

export function isEnhancedMutationProposal(value: unknown): value is EnhancedMutationProposal {
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
    validText(proposal.summary, MAX_SUMMARY) &&
    Number.isSafeInteger(proposal.expiresAt) &&
    (proposal.expiresAt as number) > 0
  )
}

/**
 * Renderer-only consent surface. It never receives tool arguments and never executes a writer;
 * confirmation merely asks the privileged owner to claim its exact pending proposal.
 */
export function EnhancedMutationConfirmation({ api }: { api?: EnhancedMutationProposalApi }) {
  const [pending, setPending] = useState<EnhancedMutationProposal | null>(null)
  const [submitting, setSubmitting] = useState(false)
  const pendingRef = useRef<EnhancedMutationProposal | null>(null)
  const consumedRef = useRef(new Set<string>())
  const mountedRef = useRef(true)

  const consume = useCallback(
    async (action: 'confirm' | 'cancel', expected?: EnhancedMutationProposal) => {
      const proposal = pendingRef.current
      if (!api || !proposal || (expected && proposal !== expected)) return
      const key = `${proposal.documentId}\0${proposal.generation}\0${proposal.proposalId}`
      if (consumedRef.current.has(key)) return
      consumedRef.current.add(key)
      if (mountedRef.current) setSubmitting(true)
      pendingRef.current = null
      if (mountedRef.current) setPending(null)
      try {
        const method = action === 'confirm' ? api.confirmProposal : api.cancelProposal
        await method.call(api, proposal.documentId, proposal.generation, proposal.proposalId)
      } catch {
        // The privileged owner remains fail-closed. The renderer must not retry a consumed
        // consent token or infer whether the pending transaction still exists.
      } finally {
        if (mountedRef.current) setSubmitting(false)
      }
    },
    [api],
  )

  useEffect(() => {
    mountedRef.current = true
    if (!api) return
    const unsubscribe = api.onProposal((candidate) => {
      if (!isEnhancedMutationProposal(candidate) || candidate.expiresAt <= Date.now()) return
      const key = `${candidate.documentId}\0${candidate.generation}\0${candidate.proposalId}`
      if (consumedRef.current.has(key)) return
      const previous = pendingRef.current
      if (previous) {
        const previousKey = `${previous.documentId}\0${previous.generation}\0${previous.proposalId}`
        if (previousKey === key) return
        void consume('cancel', previous)
      }
      pendingRef.current = candidate
      setPending(candidate)
    })
    return () => {
      mountedRef.current = false
      unsubscribe()
      const current = pendingRef.current
      if (current) void consume('cancel', current)
    }
  }, [api, consume])

  useEffect(() => {
    if (!pending) return
    const remaining = pending.expiresAt - Date.now()
    if (remaining <= 0) {
      void consume('cancel', pending)
      return
    }
    const timer = window.setTimeout(() => void consume('cancel', pending), remaining)
    return () => window.clearTimeout(timer)
  }, [consume, pending])

  if (!pending) return null
  return (
    <div className="enhanced-confirm-backdrop" role="presentation">
      <section
        className="enhanced-confirm-dialog"
        role="alertdialog"
        aria-modal="true"
        aria-labelledby="enhanced-confirm-title"
        aria-describedby="enhanced-confirm-summary"
      >
        <h2 id="enhanced-confirm-title">Confirm document change</h2>
        <p className="enhanced-confirm-tool">{pending.toolName.replace(/[_-]+/g, ' ')}</p>
        <p id="enhanced-confirm-summary">{pending.summary}</p>
        <p className="enhanced-confirm-warning">
          Review this request carefully. WisWork will apply it as one bounded transaction only after
          you confirm.
        </p>
        <div className="enhanced-confirm-actions">
          <button
            type="button"
            data-action="cancel"
            disabled={submitting}
            onClick={() => void consume('cancel')}
          >
            Reject
          </button>
          <button
            type="button"
            data-action="confirm"
            disabled={submitting}
            onClick={() => void consume('confirm')}
          >
            Confirm change
          </button>
        </div>
      </section>
    </div>
  )
}
