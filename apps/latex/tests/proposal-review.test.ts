import { describe, expect, it, vi } from 'vitest'
import {
  ProposalVerificationController,
  proposalForSelection,
  requestProposalVerification,
  reviewAction,
  verificationCompileComparison,
} from '../src/renderer/ai/proposal-review.js'

const proposal = {
  id: 'original',
  projectId: 'project-1',
  expiresAt: 1_000,
  files: [
    { path: 'a.tex', beforeText: 'a0', beforeSha256: 'a-hash', afterText: 'a1' },
    { path: 'b.tex', beforeText: 'b0', beforeSha256: 'b-hash', afterText: 'b1' },
  ],
}

describe('proposal review selection', () => {
  it('keeps the original id only for the unchanged full proposal', async () => {
    const create = vi.fn()
    await expect(proposalForSelection(proposal, new Set(['a.tex', 'b.tex']), create)).resolves.toBe(
      proposal,
    )
    expect(create).not.toHaveBeenCalled()
  })

  it('creates a fresh one-time proposal every time the selected content changes', async () => {
    const create = vi
      .fn()
      .mockResolvedValueOnce({ ...proposal, id: 'subset-1', files: [proposal.files[0]] })
      .mockResolvedValueOnce({ ...proposal, id: 'subset-2', files: [proposal.files[0]] })
    await expect(proposalForSelection(proposal, new Set(['a.tex']), create)).resolves.toMatchObject(
      {
        id: 'subset-1',
      },
    )
    await expect(proposalForSelection(proposal, new Set(['a.tex']), create)).resolves.toMatchObject(
      {
        id: 'subset-2',
      },
    )
    expect(create).toHaveBeenCalledTimes(2)
    expect(create).toHaveBeenLastCalledWith([{ path: 'a.tex', afterText: 'a1' }])
  })

  it('does not authorize an empty selection', async () => {
    await expect(proposalForSelection(proposal, new Set(), vi.fn())).rejects.toThrow(/select/i)
  })

  it('requires subset verification before apply and permits explicit risky apply only after settling', () => {
    expect(reviewAction(proposal, new Set(['a.tex']), { state: 'verified' })).toEqual({
      kind: 'verify-selection',
      label: 'Verify selected',
      disabled: false,
    })
    expect(reviewAction(proposal, new Set(['a.tex', 'b.tex']), { state: 'verifying' })).toEqual({
      kind: 'apply',
      label: 'Verifying…',
      disabled: true,
    })
    expect(reviewAction(proposal, new Set(['a.tex', 'b.tex']), { state: 'failed' })).toEqual({
      kind: 'apply',
      label: 'Apply without successful verification',
      disabled: false,
    })
    expect(reviewAction(proposal, new Set(['a.tex', 'b.tex']), { state: 'rejected' })).toEqual({
      kind: 'apply',
      label: 'Regenerate proposal to apply',
      disabled: true,
    })
  })
})

describe('proposal verification generation isolation', () => {
  it('rejects late verification after a replacement proposal, project switch, or cancel', () => {
    const controller = new ProposalVerificationController()
    const first = controller.begin('project-a', 'proposal-1')
    const second = controller.begin('project-a', 'proposal-2')
    expect(controller.accepts(first)).toBe(false)
    expect(controller.accepts(second)).toBe(true)

    const switched = controller.begin('project-b', 'proposal-3')
    expect(controller.accepts(second)).toBe(false)
    expect(controller.accepts(switched)).toBe(true)

    controller.cancel()
    expect(controller.accepts(switched)).toBe(false)
  })

  it('commits only the latest deferred verification and maps IPC refusal to rejected', async () => {
    const controller = new ProposalVerificationController()
    let resolveFirst!: (value: unknown) => void
    const firstResult = new Promise((resolve) => {
      resolveFirst = resolve
    })
    const states: string[] = []
    const first = requestProposalVerification(
      controller,
      proposal,
      () => firstResult as never,
      (state) => states.push(`${proposal.id}:${state.state}`),
    )
    const replacement = { ...proposal, id: 'replacement' }
    await requestProposalVerification(
      controller,
      replacement,
      async () => ({
        ok: false,
        error: { code: 'LATEX_CONFLICT', message: 'Proposal baseline changed on disk' },
      }),
      (state) => states.push(`${replacement.id}:${state.state}`),
    )
    resolveFirst({
      ok: true,
      value: {
        proposalId: proposal.id,
        state: 'verified',
        diagnostics: [],
        logSummary: '',
        verifiedAt: 1,
      },
    })
    await first

    expect(states).toEqual(['original:verifying', 'replacement:verifying', 'replacement:rejected'])
  })

  it('compares isolated and formal compile diagnostic counts', () => {
    expect(verificationCompileComparison({ state: 'verified', diagnosticCount: 2 }, 2)).toContain(
      'consistent',
    )
    expect(verificationCompileComparison({ state: 'failed', diagnosticCount: 3 }, 1)).toContain(
      'different',
    )
    expect(
      verificationCompileComparison({ state: 'unverifiable', diagnosticCount: 0 }, 1),
    ).toContain('unavailable')
  })
})
