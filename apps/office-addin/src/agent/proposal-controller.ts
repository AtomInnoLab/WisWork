import type { OfficeDocumentClient } from '../office-document.js'

export type ProposalOperation = 'replace' | 'append'
export const MAX_PROPOSAL_SELECTION_LENGTH = 12_000

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

export function selectionFingerprint(value: string): string {
  let hash = 0x811c9dc5
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index)
    hash = Math.imul(hash, 0x01000193)
  }
  return `${value.length}:${(hash >>> 0).toString(16).padStart(8, '0')}`
}

export function createProposalController(document: OfficeDocumentClient): ProposalController {
  let current: OfficeProposal | undefined

  const invalidate = () => {
    current = undefined
  }

  return {
    pending: () => current,

    async propose(operation, value) {
      const before = await document.readSelection()
      if (before.length > MAX_PROPOSAL_SELECTION_LENGTH) {
        throw new Error('selection_too_large')
      }
      const proposal = {
        id: crypto.randomUUID(),
        operation,
        before,
        value,
        fingerprint: selectionFingerprint(before),
      }
      current = proposal
      return proposal
    },

    async confirm(id) {
      const proposal = current
      if (!proposal || proposal.id !== id) throw new Error('proposal_missing')
      current = undefined
      const selection = await document.readSelection()
      if (
        selection !== proposal.before ||
        selectionFingerprint(selection) !== proposal.fingerprint
      ) {
        throw new Error('proposal_stale')
      }
      if (proposal.operation === 'replace') await document.replaceSelection(proposal.value)
      else await document.appendText(selection, proposal.value)
    },

    reject: invalidate,
    newTurn: invalidate,
    logout: invalidate,
  }
}
