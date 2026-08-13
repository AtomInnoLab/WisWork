import { describe, expect, it } from 'vitest'
import { reviewAction, verificationCompileComparison } from '../src/renderer/ai/proposal-review.js'

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
      kind: 'review-risk',
      label: 'Review risk',
      disabled: false,
    })
    expect(reviewAction(proposal, new Set(['a.tex', 'b.tex']), { state: 'failed' }, true)).toEqual({
      kind: 'apply',
      label: 'Apply unverified changes',
      disabled: false,
    })
    expect(reviewAction(proposal, new Set(['a.tex', 'b.tex']), { state: 'rejected' })).toEqual({
      kind: 'apply',
      label: 'Regenerate proposal to apply',
      disabled: true,
    })
  })
})

describe('proposal verification comparison', () => {
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
