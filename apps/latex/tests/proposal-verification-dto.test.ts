import { describe, expect, it } from 'vitest'
import { normalizeProposalDiagnostics } from '../src/shared/proposal-verification.js'
import { validateProposalVerification } from '../src/renderer/ai/proposal-workflow.js'

describe('proposal verification diagnostic DTO', () => {
  it('bounds paths/messages/count and clamps unsafe numeric positions for renderer acceptance', () => {
    const diagnostics = normalizeProposalDiagnostics([
      {
        path: `${'路'.repeat(800)}.tex`,
        line: Number.MAX_VALUE,
        column: -42,
        severity: 'warning',
        message: '😀'.repeat(2_000),
      },
      ...Array.from({ length: 120 }, () => ({
        path: null,
        line: 1,
        column: null,
        severity: 'error' as const,
        message: 'error',
      })),
      { path: 'bad.tex', line: 1, column: 1, severity: 'fatal', message: 'bad' },
    ])
    expect(diagnostics).toHaveLength(100)
    expect(new TextEncoder().encode(diagnostics[0]!.path!).byteLength).toBeLessThanOrEqual(1_024)
    expect(new TextEncoder().encode(diagnostics[0]!.message).byteLength).toBeLessThanOrEqual(4_096)
    expect(diagnostics[0]!.line).toBe(10_000_000)
    expect(diagnostics[0]!.column).toBe(1)
    expect(() =>
      validateProposalVerification(
        {
          proposalId: 'p',
          state: 'verified',
          diagnostics,
          logSummary: '',
          verifiedAt: Date.now(),
        },
        'p',
      ),
    ).not.toThrow()
  })
})
