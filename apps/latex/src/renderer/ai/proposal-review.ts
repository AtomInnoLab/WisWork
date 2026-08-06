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
