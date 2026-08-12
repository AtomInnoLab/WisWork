import type {
  LatexIpcErrorCode,
  LatexIpcResult,
  ProposalVerificationDiagnosticDto,
  ProposalVerificationDto,
} from '../../shared/ipc.js'
import {
  reviewAction,
  validateReviewProposal,
  verificationCompileComparison,
  type ReviewProposal,
  type ReviewVerification,
} from './proposal-review.js'
import {
  MAX_FORMAL_COMPILE_DIAGNOSTICS,
  MAX_PROPOSAL_DIAGNOSTICS,
  MAX_PROPOSAL_DIAGNOSTIC_MESSAGE_BYTES,
  MAX_PROPOSAL_DIAGNOSTIC_PATH_BYTES,
  MAX_PROPOSAL_DIAGNOSTIC_POSITION,
} from '../../shared/proposal-verification.js'

const MAX_DIAGNOSTICS = MAX_PROPOSAL_DIAGNOSTICS
const MAX_FORMAL_DIAGNOSTICS = MAX_FORMAL_COMPILE_DIAGNOSTICS
const MAX_LOG_BYTES = 16_000
const MAX_REASON_BYTES = 4_096
const MAX_MESSAGE_BYTES = MAX_PROPOSAL_DIAGNOSTIC_MESSAGE_BYTES
const MAX_FORMAL_LOG_BYTES = 2 * 1024 * 1024
const encoder = new TextEncoder()

type WorkflowResult<T> = LatexIpcResult<T>

interface ProposalWorkflowDependencies {
  verify(request: { projectId: string; proposalId: string }): Promise<WorkflowResult<unknown>>
  create(request: {
    projectId: string
    files: Array<{ path: string; afterText: string }>
  }): Promise<WorkflowResult<unknown>>
  apply(request: { projectId: string; proposalId: string }): Promise<WorkflowResult<unknown>>
  refresh(): void | Promise<void>
}

export interface ProposalWorkflowState {
  projectId: string
  proposal: ReviewProposal | null
  selection: ReadonlySet<string>
  verification: ReviewVerification | null
  riskArmed: boolean
  busy: boolean
  snapshotId: string | null
  status: string | null
}

interface AppliedProposalResult {
  proposalId: string
  snapshotId: string
  compile:
    | {
        ok: true
        result: {
          revision: number
          pdfUrl: string | null
          diagnostics: ProposalVerificationDiagnosticDto[]
          log: string
        }
      }
    | { ok: false; error: string }
}

export type UndoProposalResult =
  | { snapshotId: string; restored: false; compile: null }
  | { snapshotId: string; restored: true; compile: AppliedProposalResult['compile'] }

function exactKeys(record: Record<string, unknown>, expected: readonly string[]): boolean {
  return (
    Object.keys(record).length === expected.length &&
    Object.keys(record).every((key) => expected.includes(key))
  )
}

function boundedString(
  value: unknown,
  maxBytes: number,
  allowNull = false,
): value is string | null {
  return (
    (allowNull && value === null) ||
    (typeof value === 'string' && encoder.encode(value).byteLength <= maxBytes)
  )
}

function safePosition(value: unknown): value is number | null {
  return (
    value === null ||
    (Number.isSafeInteger(value) &&
      (value as number) >= 1 &&
      (value as number) <= MAX_PROPOSAL_DIAGNOSTIC_POSITION)
  )
}

function validateDiagnostic(value: unknown): ProposalVerificationDiagnosticDto {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('Verification diagnostic was invalid.')
  const record = value as Record<string, unknown>
  if (
    !exactKeys(record, ['path', 'line', 'column', 'severity', 'message']) ||
    !boundedString(record.path, MAX_PROPOSAL_DIAGNOSTIC_PATH_BYTES, true) ||
    (record.path !== null && (!record.path || record.path.includes('\0'))) ||
    !safePosition(record.line) ||
    !safePosition(record.column) ||
    (record.severity !== 'error' && record.severity !== 'warning') ||
    !boundedString(record.message, MAX_MESSAGE_BYTES) ||
    !record.message
  ) {
    throw new Error('Verification diagnostic was invalid.')
  }
  return record as unknown as ProposalVerificationDiagnosticDto
}

export function validateProposalVerification(
  value: unknown,
  proposalId: string,
): ProposalVerificationDto {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('Proposal verification response was invalid.')
  const record = value as Record<string, unknown>
  const state = record.state
  const expectedKeys =
    state === 'verified'
      ? ['proposalId', 'state', 'diagnostics', 'logSummary', 'verifiedAt']
      : ['proposalId', 'state', 'reason', 'diagnostics', 'logSummary', 'verifiedAt']
  if (
    (state !== 'verified' && state !== 'failed' && state !== 'unverifiable') ||
    !exactKeys(record, expectedKeys) ||
    record.proposalId !== proposalId ||
    !Number.isSafeInteger(record.verifiedAt) ||
    (record.verifiedAt as number) < 0 ||
    !Number.isFinite(new Date(record.verifiedAt as number).getTime()) ||
    !Array.isArray(record.diagnostics) ||
    record.diagnostics.length > MAX_DIAGNOSTICS ||
    !boundedString(record.logSummary, MAX_LOG_BYTES) ||
    (state !== 'verified' && (!boundedString(record.reason, MAX_REASON_BYTES) || !record.reason))
  ) {
    throw new Error('Proposal verification response was invalid.')
  }
  const diagnostics = record.diagnostics.map(validateDiagnostic)
  return { ...record, diagnostics } as ProposalVerificationDto
}

export function validateAppliedProposal(value: unknown, proposalId: string): AppliedProposalResult {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('Apply response was invalid.')
  const record = value as Record<string, unknown>
  if (
    !exactKeys(record, ['proposalId', 'snapshotId', 'compile']) ||
    record.proposalId !== proposalId ||
    typeof record.snapshotId !== 'string' ||
    !/^[a-f0-9]{32}$/.test(record.snapshotId) ||
    !record.compile ||
    typeof record.compile !== 'object' ||
    Array.isArray(record.compile)
  ) {
    throw new Error('Apply response was invalid.')
  }
  const compile = record.compile as Record<string, unknown>
  if (compile.ok === false) {
    if (
      !exactKeys(compile, ['ok', 'error']) ||
      !boundedString(compile.error, MAX_REASON_BYTES) ||
      !compile.error
    )
      throw new Error('Apply response was invalid.')
    return record as unknown as AppliedProposalResult
  }
  if (
    compile.ok !== true ||
    !exactKeys(compile, ['ok', 'result']) ||
    !compile.result ||
    typeof compile.result !== 'object' ||
    Array.isArray(compile.result)
  ) {
    throw new Error('Apply response was invalid.')
  }
  const result = compile.result as Record<string, unknown>
  if (
    !exactKeys(result, ['revision', 'pdfUrl', 'diagnostics', 'log']) ||
    !Number.isSafeInteger(result.revision) ||
    (result.revision as number) < 0 ||
    !boundedString(result.pdfUrl, 4_096, true) ||
    !Array.isArray(result.diagnostics) ||
    result.diagnostics.length > MAX_FORMAL_DIAGNOSTICS ||
    !boundedString(result.log, MAX_FORMAL_LOG_BYTES)
  ) {
    throw new Error('Apply response was invalid.')
  }
  result.diagnostics = result.diagnostics.map(validateDiagnostic)
  return record as unknown as AppliedProposalResult
}

export function validateUndoProposal(value: unknown, snapshotId: string): UndoProposalResult {
  if (!value || typeof value !== 'object' || Array.isArray(value))
    throw new Error('Undo response was invalid.')
  const record = value as Record<string, unknown>
  if (!exactKeys(record, ['snapshotId', 'restored', 'compile']) || record.snapshotId !== snapshotId)
    throw new Error('Undo response was invalid.')
  if (record.restored === false && record.compile === null) {
    return record as unknown as UndoProposalResult
  }
  if (record.restored !== true || record.compile === null) {
    throw new Error('Undo response was invalid.')
  }
  const validated = validateAppliedProposal(
    { proposalId: 'undo', snapshotId, compile: record.compile },
    'undo',
  )
  return { snapshotId, restored: true, compile: validated.compile }
}

function recoverSnapshotId(value: unknown, proposalId: string): string | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null
  const record = value as Record<string, unknown>
  return record.proposalId === proposalId &&
    typeof record.snapshotId === 'string' &&
    /^[a-f0-9]{32}$/.test(record.snapshotId)
    ? record.snapshotId
    : null
}

function verificationState(value: ProposalVerificationDto): ReviewVerification {
  if (value.state === 'verified') return { state: 'verified', evidence: value }
  if (value.state === 'failed') return { state: 'failed', evidence: value }
  return { state: 'unverifiable', evidence: value }
}

function rejection(error: unknown): Extract<ReviewVerification, { state: 'rejected' }> {
  if (error && typeof error === 'object' && !Array.isArray(error)) {
    const record = error as Record<string, unknown>
    const codes: LatexIpcErrorCode[] = [
      'LATEX_FORBIDDEN_SENDER',
      'LATEX_PROJECT_SESSION_MISMATCH',
      'LATEX_INVALID_PAYLOAD',
      'LATEX_CONFLICT',
      'LATEX_NOT_FOUND',
      'LATEX_VERIFICATION_REJECTED',
      'LATEX_INTERNAL',
    ]
    if (
      codes.includes(record.code as LatexIpcErrorCode) &&
      boundedString(record.message, MAX_MESSAGE_BYTES) &&
      record.message
    ) {
      return {
        state: 'rejected',
        code: record.code as LatexIpcErrorCode,
        message: record.message,
      }
    }
  }
  return {
    state: 'rejected',
    code: 'LATEX_INVALID_PAYLOAD',
    message: error instanceof Error ? error.message : 'The proposal response was invalid.',
  }
}

export class ProposalWorkflow {
  private generation = 0
  private listener: ((state: ProposalWorkflowState) => void) | null = null
  state: ProposalWorkflowState

  constructor(
    projectId: string,
    private readonly dependencies: ProposalWorkflowDependencies,
  ) {
    this.state = {
      projectId,
      proposal: null,
      selection: new Set(),
      verification: null,
      riskArmed: false,
      busy: false,
      snapshotId: null,
      status: null,
    }
  }

  subscribe(listener: (state: ProposalWorkflowState) => void): () => void {
    this.listener = listener
    listener(this.state)
    return () => {
      if (this.listener === listener) this.listener = null
    }
  }

  setProject(projectId: string): void {
    this.generation += 1
    this.update({
      projectId,
      proposal: null,
      selection: new Set(),
      verification: null,
      riskArmed: false,
      busy: false,
      snapshotId: null,
      status: null,
    })
  }

  setProposal(value: unknown): Promise<void> {
    let proposal: ReviewProposal
    try {
      proposal = validateReviewProposal(value, this.state.projectId)
    } catch (error) {
      this.update({ status: error instanceof Error ? error.message : String(error) })
      return Promise.resolve()
    }
    this.generation += 1
    const generation = this.generation
    this.update({
      proposal,
      selection: new Set(proposal.files.map((file) => file.path)),
      verification: { state: 'verifying' },
      riskArmed: false,
      status: null,
    })
    return this.verify(proposal, generation)
  }

  setSelection(selection: ReadonlySet<string>): void {
    if (!this.state.proposal) return
    const allowed = new Set(this.state.proposal.files.map((file) => file.path))
    const next = new Set([...selection].filter((path) => allowed.has(path)))
    const currentKey = [...this.state.selection].sort().join('\0')
    const nextKey = [...next].sort().join('\0')
    this.update({ selection: next, ...(currentKey === nextKey ? {} : { riskArmed: false }) })
  }

  async primaryAction(): Promise<void> {
    const { proposal, verification, selection, riskArmed } = this.state
    if (!proposal || !verification || this.state.busy) return
    const action = reviewAction(proposal, selection, verification, riskArmed)
    if (action.disabled) return
    if (action.kind === 'review-risk') {
      this.update({
        riskArmed: true,
        status: 'Review the verification warning, then confirm the unverified apply separately.',
      })
      return
    }
    if (action.kind === 'verify-selection') {
      await this.createSubset(proposal, selection)
      return
    }
    await this.apply(proposal, verification)
  }

  cancel(): void {
    this.generation += 1
    this.update({
      proposal: null,
      selection: new Set(),
      verification: null,
      riskArmed: false,
      busy: false,
    })
  }

  clearSnapshot(status: string): void {
    this.update({ snapshotId: null, status })
  }

  private async verify(proposal: ReviewProposal, generation: number): Promise<void> {
    try {
      const result = await this.dependencies.verify({
        projectId: proposal.projectId,
        proposalId: proposal.id,
      })
      if (!this.accepts(generation, proposal.id)) return
      if (!result.ok) {
        this.update({ verification: rejection(result.error), riskArmed: false })
        return
      }
      const value = validateProposalVerification(result.value, proposal.id)
      if (!this.accepts(generation, proposal.id)) return
      this.update({ verification: verificationState(value), riskArmed: false })
    } catch (error) {
      if (this.accepts(generation, proposal.id))
        this.update({ verification: rejection(error), riskArmed: false })
    }
  }

  private async createSubset(
    proposal: ReviewProposal,
    selection: ReadonlySet<string>,
  ): Promise<void> {
    const selected = proposal.files.filter((file) => selection.has(file.path))
    if (!selected.length || selected.length === proposal.files.length) return
    const generation = this.generation
    this.update({ busy: true, riskArmed: false, status: null })
    try {
      const result = await this.dependencies.create({
        projectId: proposal.projectId,
        files: selected.map(({ path, afterText }) => ({ path, afterText })),
      })
      if (!this.accepts(generation, proposal.id)) return
      if (!result.ok) throw new Error(result.error.message)
      const fresh = validateReviewProposal(result.value, proposal.projectId)
      this.update({ busy: false })
      await this.setProposal(fresh)
    } catch (error) {
      if (this.accepts(generation, proposal.id))
        this.update({
          busy: false,
          status: error instanceof Error ? error.message : String(error),
        })
    }
  }

  private async apply(proposal: ReviewProposal, verification: ReviewVerification): Promise<void> {
    if (verification.state === 'verifying' || verification.state === 'rejected') return
    const generation = this.generation
    this.update({ busy: true, status: null })
    let applied: AppliedProposalResult
    try {
      const result = await this.dependencies.apply({
        projectId: proposal.projectId,
        proposalId: proposal.id,
      })
      if (!this.accepts(generation, proposal.id)) return
      if (!result.ok) throw new Error(result.error.message)
      try {
        applied = validateAppliedProposal(result.value, proposal.id)
      } catch (error) {
        const snapshotId = recoverSnapshotId(result.value, proposal.id)
        this.generation += 1
        this.update({
          proposal: null,
          selection: new Set(),
          verification: null,
          riskArmed: false,
          busy: false,
          snapshotId,
          status: `${error instanceof Error ? error.message : String(error)} The project may have changed; reload before editing again.${snapshotId ? ' Undo remains available.' : ''}`,
        })
        return
      }
    } catch (error) {
      if (this.accepts(generation, proposal.id))
        this.update({
          busy: false,
          riskArmed: false,
          status: error instanceof Error ? error.message : String(error),
        })
      return
    }

    const comparison = applied.compile.ok
      ? verificationCompileComparison(
          {
            state: verification.state,
            diagnosticCount: verification.evidence.diagnostics.length,
          },
          applied.compile.result.diagnostics.length,
        )
      : `Formal compile failed: ${applied.compile.error}. Isolation reported ${verification.evidence.diagnostics.length} diagnostics; the results are different.`
    this.generation += 1
    const appliedProject = this.state.projectId
    this.update({
      proposal: null,
      selection: new Set(),
      verification: null,
      riskArmed: false,
      busy: false,
      snapshotId: applied.snapshotId,
      status: `Changes applied. ${comparison}`,
    })
    try {
      await this.dependencies.refresh()
    } catch (error) {
      if (this.state.projectId === appliedProject && this.state.snapshotId === applied.snapshotId) {
        this.update({
          status: `${this.state.status} File refresh failed: ${error instanceof Error ? error.message : String(error)}. Undo remains available.`,
        })
      }
    }
  }

  private accepts(generation: number, proposalId: string): boolean {
    return this.generation === generation && this.state.proposal?.id === proposalId
  }

  private update(patch: Partial<ProposalWorkflowState>): void {
    this.state = { ...this.state, ...patch }
    this.listener?.(this.state)
  }
}
