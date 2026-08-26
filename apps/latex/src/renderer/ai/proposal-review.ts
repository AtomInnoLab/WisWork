import type { LatexIpcErrorCode, ProposalVerificationDto } from '../../shared/ipc.js'

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
  kind: 'apply' | 'verify-selection' | 'review-risk'
  label: string
  disabled: boolean
}

export function reviewAction(
  proposal: ReviewProposal,
  selectedPaths: ReadonlySet<string>,
  verification: { state: ReviewVerification['state'] },
  riskArmed = false,
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
  if (!riskArmed) return { kind: 'review-risk', label: 'Review risk', disabled: false }
  return {
    kind: 'apply',
    label: 'Apply unverified changes',
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

export function validateReviewProposal(
  value: unknown,
  projectId: string,
  proposalId?: string,
): ReviewProposal {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('The proposal response was invalid.')
  const proposal = value as Record<string, unknown>
  if (
    typeof proposal.id !== 'string' ||
    !proposal.id ||
    proposal.id.length > 128 ||
    (proposalId !== undefined && proposal.id !== proposalId) ||
    proposal.projectId !== projectId ||
    !Number.isSafeInteger(proposal.expiresAt) ||
    (proposal.expiresAt as number) < 0 ||
    Object.keys(proposal).some((key) => !['id', 'projectId', 'expiresAt', 'files'].includes(key)) ||
    !Array.isArray(proposal.files) ||
    proposal.files.length < 1 ||
    proposal.files.length > 20
  ) {
    throw new Error('The proposal response was invalid.')
  }
  let totalBytes = 0
  const seenPaths = new Set<string>()
  for (const value of proposal.files) {
    if (!value || typeof value !== 'object' || Array.isArray(value))
      throw new Error('The proposal response was invalid.')
    const file = value as Record<string, unknown>
    if (
      typeof file.path !== 'string' ||
      !file.path ||
      file.path.length > 1_024 ||
      file.path.includes('\0') ||
      file.path.includes('\\') ||
      file.path.startsWith('/') ||
      file.path.split('/').some((part) => !part || part === '.' || part === '..') ||
      (file.beforeText !== null && typeof file.beforeText !== 'string') ||
      (file.beforeText === null
        ? file.beforeSha256 !== null
        : typeof file.beforeSha256 !== 'string' || !/^[a-f0-9]{64}$/.test(file.beforeSha256)) ||
      typeof file.afterText !== 'string' ||
      Object.keys(file).some(
        (key) => !['path', 'beforeText', 'beforeSha256', 'afterText'].includes(key),
      )
    ) {
      throw new Error('The proposal response was invalid.')
    }
    if (seenPaths.has(file.path)) throw new Error('The proposal response was invalid.')
    seenPaths.add(file.path)
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
