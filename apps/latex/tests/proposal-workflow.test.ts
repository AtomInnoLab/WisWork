import { describe, expect, it, vi } from 'vitest'
import {
  ProposalWorkflow,
  validateAppliedProposal,
  validateProposalVerification,
} from '../src/renderer/ai/proposal-workflow.js'

const hash = 'a'.repeat(64)

function proposal(id = 'proposal-1', paths = ['a.tex', 'b.tex']) {
  return {
    id,
    projectId: 'project-1',
    expiresAt: 10_000,
    files: paths.map((path) => ({
      path,
      beforeText: `before ${path}`,
      beforeSha256: hash,
      afterText: `after ${path}`,
    })),
  }
}

function verification(proposalId: string, state: 'verified' | 'failed' | 'unverifiable') {
  return {
    proposalId,
    state,
    ...((state === 'failed' || state === 'unverifiable') && { reason: `${state} reason` }),
    diagnostics: [],
    logSummary: 'bounded log',
    verifiedAt: 1,
  }
}

function deferred<T>() {
  let resolve!: (value: T) => void
  let reject!: (error: unknown) => void
  const promise = new Promise<T>((done, fail) => {
    resolve = done
    reject = fail
  })
  return { promise, resolve, reject }
}

function harness(overrides: Record<string, unknown> = {}) {
  const deps = {
    verify: vi.fn(async ({ proposalId }: { proposalId: string }) => ({
      ok: true as const,
      value: verification(proposalId, 'verified'),
    })),
    create: vi.fn(),
    apply: vi.fn(),
    refresh: vi.fn(async () => undefined),
    ...overrides,
  }
  return { deps, workflow: new ProposalWorkflow('project-1', deps as never) }
}

describe('proposal workflow wiring', () => {
  it('automatically verifies a proposal and ignores stale proposal/project results', async () => {
    const first = deferred<unknown>()
    const second = deferred<unknown>()
    const verify = vi.fn().mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise)
    const { workflow } = harness({ verify })
    const firstRun = workflow.setProposal(proposal('proposal-1'))
    const secondRun = workflow.setProposal(proposal('proposal-2'))
    first.resolve({ ok: true, value: verification('proposal-1', 'verified') })
    await firstRun
    expect(workflow.state.proposal?.id).toBe('proposal-2')
    expect(workflow.state.verification?.state).toBe('verifying')

    workflow.setProject('project-2')
    second.resolve({ ok: true, value: verification('proposal-2', 'verified') })
    await secondRun
    expect(workflow.state.projectId).toBe('project-2')
    expect(workflow.state.proposal).toBeNull()
  })

  it('discards a deferred verification result after review cancellation', async () => {
    const pending = deferred<unknown>()
    const { workflow } = harness({ verify: vi.fn(() => pending.promise) })
    const verifying = workflow.setProposal(proposal())
    workflow.cancel()
    pending.resolve({ ok: true, value: verification('proposal-1', 'verified') })
    await verifying
    expect(workflow.state.proposal).toBeNull()
    expect(workflow.state.verification).toBeNull()
  })

  it('creates and verifies a fresh subset before a separate apply action', async () => {
    const fresh = proposal('subset-1', ['a.tex'])
    const create = vi.fn(async () => ({ ok: true as const, value: fresh }))
    const apply = vi.fn(async () => ({
      ok: true as const,
      value: {
        proposalId: fresh.id,
        snapshotId: 'b'.repeat(32),
        compile: {
          ok: true,
          result: { revision: 2, pdfUrl: null, diagnostics: [], log: '' },
        },
      },
    }))
    const { workflow, deps } = harness({ create, apply })
    await workflow.setProposal(proposal())
    workflow.setSelection(new Set(['a.tex']))

    await workflow.primaryAction()
    expect(create).toHaveBeenCalledTimes(1)
    expect(deps.verify).toHaveBeenLastCalledWith({
      projectId: 'project-1',
      proposalId: 'subset-1',
    })
    expect(apply).not.toHaveBeenCalled()
    expect(workflow.state.proposal?.id).toBe('subset-1')

    await workflow.primaryAction()
    expect(apply).toHaveBeenCalledTimes(1)
    expect(workflow.state.snapshotId).toBe('b'.repeat(32))
  })

  it('requires two independent actions for failed verification and resets arming', async () => {
    const apply = vi.fn(async () => ({ ok: false as const, error: { message: 'unused' } }))
    const verify = vi.fn(async ({ proposalId }: { proposalId: string }) => ({
      ok: true as const,
      value: verification(proposalId, 'failed'),
    }))
    const { workflow } = harness({ verify, apply })
    await workflow.setProposal(proposal())

    await workflow.primaryAction()
    expect(workflow.state.riskArmed).toBe(true)
    expect(apply).not.toHaveBeenCalled()

    await workflow.primaryAction()
    expect(apply).toHaveBeenCalledTimes(1)
    expect(workflow.state.riskArmed).toBe(false)

    await workflow.primaryAction()
    expect(workflow.state.riskArmed).toBe(true)
    expect(apply).toHaveBeenCalledTimes(1)

    workflow.setSelection(new Set(['a.tex']))
    expect(workflow.state.riskArmed).toBe(false)
    workflow.setSelection(new Set(['a.tex', 'b.tex']))
    await workflow.primaryAction()
    expect(workflow.state.riskArmed).toBe(true)
    expect(apply).toHaveBeenCalledTimes(1)

    await workflow.setProposal(proposal('proposal-2'))
    expect(workflow.state.riskArmed).toBe(false)
  })

  it('blocks rejected and malformed verification responses', async () => {
    const apply = vi.fn()
    const verify = vi.fn(async () => ({
      ok: true as const,
      value: { ...verification('proposal-1', 'verified'), unexpected: true },
    }))
    const { workflow } = harness({ verify, apply })
    await workflow.setProposal(proposal())
    expect(workflow.state.verification).toMatchObject({
      state: 'rejected',
      code: 'LATEX_INVALID_PAYLOAD',
    })
    await workflow.primaryAction()
    expect(apply).not.toHaveBeenCalled()
  })

  it('validates a fresh subset response before replacing the owned proposal', async () => {
    const create = vi.fn(async () => ({
      ok: true as const,
      value: { ...proposal('subset-1', ['a.tex']), unexpected: true },
    }))
    const { workflow, deps } = harness({ create })
    await workflow.setProposal(proposal())
    workflow.setSelection(new Set(['a.tex']))
    await workflow.primaryAction()

    expect(workflow.state.proposal?.id).toBe('proposal-1')
    expect(workflow.state.status).toMatch(/invalid/i)
    expect(deps.verify).toHaveBeenCalledTimes(1)
  })

  it('treats malformed apply success as uncertain and prevents a duplicate apply', async () => {
    const apply = vi.fn(async () => ({
      ok: true as const,
      value: {
        proposalId: 'proposal-1',
        snapshotId: 'd'.repeat(32),
        compile: { ok: false, error: 'compile failed' },
        unexpected: true,
      },
    }))
    const { workflow } = harness({ apply })
    await workflow.setProposal(proposal())
    await workflow.primaryAction()
    expect(workflow.state.proposal).toBeNull()
    expect(workflow.state.snapshotId).toBe('d'.repeat(32))
    expect(workflow.state.status).toMatch(/may have changed/i)
    expect(workflow.state.status).toMatch(/undo remains available/i)
    await workflow.primaryAction()
    expect(apply).toHaveBeenCalledTimes(1)
  })

  it('retains undo and applied status when post-apply refresh rejects', async () => {
    const apply = vi.fn(async () => ({
      ok: true as const,
      value: {
        proposalId: 'proposal-1',
        snapshotId: 'c'.repeat(32),
        compile: {
          ok: true,
          result: { revision: 3, pdfUrl: null, diagnostics: [], log: '' },
        },
      },
    }))
    const refreshing = deferred<void>()
    const refresh = vi.fn(() => refreshing.promise)
    const { workflow } = harness({ apply, refresh })
    await workflow.setProposal(proposal())
    const applying = workflow.primaryAction()

    await vi.waitFor(() => expect(workflow.state.snapshotId).toBe('c'.repeat(32)))
    expect(workflow.state.proposal).toBeNull()
    expect(workflow.state.status).toMatch(/applied/i)
    refreshing.reject(new Error('refresh exploded'))
    await applying

    expect(workflow.state.proposal).toBeNull()
    expect(workflow.state.snapshotId).toBe('c'.repeat(32))
    expect(workflow.state.status).toMatch(/applied/i)
    expect(workflow.state.status).toMatch(/refresh exploded/i)
  })
})

describe('verification DTO validation', () => {
  it('accepts exact bounded state shapes and rejects malformed evidence', () => {
    expect(
      validateProposalVerification(verification('proposal-1', 'verified'), 'proposal-1'),
    ).toMatchObject({
      state: 'verified',
    })
    const invalid = [
      { ...verification('proposal-1', 'verified'), reason: 'not allowed' },
      { ...verification('other', 'verified') },
      { ...verification('proposal-1', 'verified'), verifiedAt: Number.NaN },
      { ...verification('proposal-1', 'verified'), verifiedAt: Number.MAX_SAFE_INTEGER },
      { ...verification('proposal-1', 'failed'), reason: 'x'.repeat(5_000) },
      { ...verification('proposal-1', 'verified'), logSummary: 'x'.repeat(16_001) },
      {
        ...verification('proposal-1', 'verified'),
        diagnostics: Array.from({ length: 101 }, () => ({
          path: null,
          line: null,
          column: null,
          severity: 'error',
          message: 'error',
        })),
      },
      {
        ...verification('proposal-1', 'verified'),
        diagnostics: [{ path: null, line: -1, column: null, severity: 'fatal', message: 'x' }],
      },
    ]
    for (const value of invalid)
      expect(() => validateProposalVerification(value, 'proposal-1')).toThrow()
  })
})

describe('apply DTO validation', () => {
  const valid = {
    proposalId: 'proposal-1',
    snapshotId: 'e'.repeat(32),
    compile: {
      ok: true,
      result: { revision: 4, pdfUrl: null, diagnostics: [], log: '' },
    },
  }

  it('accepts the exact compile union and rejects unsafe snapshot or diagnostic shapes', () => {
    expect(validateAppliedProposal(valid, 'proposal-1')).toEqual(valid)
    const invalid = [
      { ...valid, extra: true },
      { ...valid, snapshotId: '../snapshot' },
      { ...valid, proposalId: 'other' },
      { ...valid, compile: { ok: false, error: 'failed', result: {} } },
      {
        ...valid,
        compile: {
          ok: true,
          result: {
            ...valid.compile.result,
            diagnostics: Array.from({ length: 1_001 }, () => ({
              path: null,
              line: null,
              column: null,
              severity: 'error',
              message: 'x',
            })),
          },
        },
      },
      {
        ...valid,
        compile: {
          ok: true,
          result: {
            ...valid.compile.result,
            diagnostics: [{ path: null, line: 1, column: null, severity: 'fatal', message: 'x' }],
          },
        },
      },
    ]
    for (const value of invalid)
      expect(() => validateAppliedProposal(value, 'proposal-1')).toThrow()
  })
})
