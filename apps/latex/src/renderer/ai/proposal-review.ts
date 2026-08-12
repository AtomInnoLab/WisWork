import type {
  LatexIpcErrorCode,
  LatexIpcResult,
  ProposalVerificationDto,
} from '../../shared/ipc.js'

export interface ReviewProposal {
  id: string
  projectId: string
  expiresAt: number
  files: Array<{
    path: string
    beforeText: string | null
    beforeSha256: string | null
    afterText: string
  }>
}

type ProposalResult = { ok: true; value: unknown } | { ok: false; error: { message: string } }

export type ReviewVerification =
  | { state: 'verifying' }
  | {
      state: 'verified'
      evidence: Extract<ProposalVerificationDto, { state: 'verified' }>
    }
  | { state: 'failed'; evidence: Extract<ProposalVerificationDto, { state: 'failed' }> }
  | {
      state: 'unverifiable'
      evidence: Extract<ProposalVerificationDto, { state: 'unverifiable' }>
    }
  | { state: 'rejected'; code: LatexIpcErrorCode; message: string }

export interface ProposalVerificationScope {
  projectId: string
  proposalId: string
  generation: number
}

export class ProposalVerificationController {
  private generation = 0
  private current: ProposalVerificationScope | null = null

  begin(projectId: string, proposalId: string): ProposalVerificationScope {
    this.generation += 1
    this.current = { projectId, proposalId, generation: this.generation }
    return this.current
  }

  accepts(scope: ProposalVerificationScope): boolean {
    return Boolean(
      this.current &&
      this.current.projectId === scope.projectId &&
      this.current.proposalId === scope.proposalId &&
      this.current.generation === scope.generation,
    )
  }

  cancel(): void {
    this.generation += 1
    this.current = null
  }
}

export async function requestProposalVerification(
  controller: ProposalVerificationController,
  proposal: ReviewProposal,
  verify: (request: {
    projectId: string
    proposalId: string
  }) => Promise<LatexIpcResult<ProposalVerificationDto>>,
  commit: (state: ReviewVerification) => void,
): Promise<void> {
  const scope = controller.begin(proposal.projectId, proposal.id)
  commit({ state: 'verifying' })
  try {
    const result = await verify({ projectId: proposal.projectId, proposalId: proposal.id })
    if (!controller.accepts(scope)) return
    if (!result.ok) {
      commit({ state: 'rejected', code: result.error.code, message: result.error.message })
      return
    }
    if (result.value.proposalId !== proposal.id) {
      commit({
        state: 'rejected',
        code: 'LATEX_INVALID_PAYLOAD',
        message: 'Verification evidence did not match the current proposal.',
      })
      return
    }
    if (result.value.state === 'verified') commit({ state: 'verified', evidence: result.value })
    else if (result.value.state === 'failed') commit({ state: 'failed', evidence: result.value })
    else commit({ state: 'unverifiable', evidence: result.value })
  } catch (error) {
    if (!controller.accepts(scope)) return
    commit({
      state: 'rejected',
      code: 'LATEX_INTERNAL',
      message: error instanceof Error ? error.message : String(error),
    })
  }
}

export function verificationCompileComparison(
  verification: { state: 'verified' | 'failed' | 'unverifiable'; diagnosticCount: number },
  formalDiagnosticCount: number,
): string {
  if (verification.state === 'unverifiable') {
    return `Isolation comparison unavailable; formal compile reported ${formalDiagnosticCount} diagnostics.`
  }
  const relation =
    verification.diagnosticCount === formalDiagnosticCount ? 'consistent' : 'different'
  return `Isolation reported ${verification.diagnosticCount} diagnostics; formal compile reported ${formalDiagnosticCount} — ${relation}.`
}

export interface ReviewAction {
  kind: 'apply' | 'verify-selection'
  label: string
  disabled: boolean
}

export function reviewAction(
  proposal: ReviewProposal,
  selectedPaths: ReadonlySet<string>,
  verification: { state: ReviewVerification['state'] },
): ReviewAction {
  const fullSelection =
    selectedPaths.size === proposal.files.length &&
    proposal.files.every((file) => selectedPaths.has(file.path))
  if (!fullSelection) {
    return {
      kind: 'verify-selection',
      label: 'Verify selected',
      disabled: selectedPaths.size === 0,
    }
  }
  if (verification.state === 'verifying')
    return { kind: 'apply', label: 'Verifying…', disabled: true }
  if (verification.state === 'rejected')
    return { kind: 'apply', label: 'Regenerate proposal to apply', disabled: true }
  if (verification.state === 'verified')
    return { kind: 'apply', label: 'Apply verified changes', disabled: false }
  return {
    kind: 'apply',
    label: 'Apply without successful verification',
    disabled: false,
  }
}

export async function loadProposalForReview(
  toolOutput: string,
  projectId: string,
  getProposal: (request: { projectId: string; proposalId: string }) => Promise<ProposalResult>,
): Promise<ReviewProposal> {
  const summary = JSON.parse(toolOutput) as unknown
  if (!summary || typeof summary !== 'object' || Array.isArray(summary))
    throw new Error('The proposal summary was invalid.')
  const record = summary as Record<string, unknown>
  if (
    Object.keys(record).some((key) => !['proposalId', 'expiresAt', 'fileCount'].includes(key)) ||
    typeof record.proposalId !== 'string' ||
    !record.proposalId ||
    record.proposalId.length > 128
  ) {
    throw new Error('The proposal summary was invalid.')
  }
  const result = await getProposal({ projectId, proposalId: record.proposalId })
  if (!result.ok) throw new Error(result.error.message)
  return validateReviewProposal(result.value, projectId, record.proposalId)
}

function validateReviewProposal(
  value: unknown,
  projectId: string,
  proposalId: string,
): ReviewProposal {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('The proposal response was invalid.')
  const proposal = value as Record<string, unknown>
  if (
    proposal.id !== proposalId ||
    proposal.projectId !== projectId ||
    typeof proposal.expiresAt !== 'number' ||
    !Array.isArray(proposal.files) ||
    proposal.files.length < 1 ||
    proposal.files.length > 20
  ) {
    throw new Error('The proposal response was invalid.')
  }
  let totalBytes = 0
  for (const value of proposal.files) {
    if (!value || typeof value !== 'object' || Array.isArray(value))
      throw new Error('The proposal response was invalid.')
    const file = value as Record<string, unknown>
    if (
      typeof file.path !== 'string' ||
      (file.beforeText !== null && typeof file.beforeText !== 'string') ||
      (file.beforeSha256 !== null && typeof file.beforeSha256 !== 'string') ||
      typeof file.afterText !== 'string'
    ) {
      throw new Error('The proposal response was invalid.')
    }
    totalBytes += new TextEncoder().encode(file.beforeText ?? '').byteLength
    totalBytes += new TextEncoder().encode(file.afterText).byteLength
    if (totalBytes > 4 * 1024 * 1024) throw new Error('The proposal response was too large.')
  }
  return value as ReviewProposal
}

export async function proposalForSelection(
  proposal: ReviewProposal,
  selectedPaths: ReadonlySet<string>,
  create: (files: Array<{ path: string; afterText: string }>) => Promise<ReviewProposal>,
): Promise<ReviewProposal> {
  const selected = proposal.files.filter((file) => selectedPaths.has(file.path))
  if (selected.length === 0) throw new Error('Select at least one file')
  if (selected.length === proposal.files.length) return proposal
  return create(selected.map(({ path, afterText }) => ({ path, afterText })))
}
