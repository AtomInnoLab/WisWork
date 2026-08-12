import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { ProposalReview } from '../src/renderer/ai/ProposalReview.js'

const proposal = {
  id: 'proposal-1',
  projectId: 'project-1',
  expiresAt: 1_000,
  files: [
    {
      path: 'main.tex',
      beforeText: 'before\nshared',
      beforeSha256: 'hash',
      afterText: 'after\nshared',
    },
  ],
}

describe('verified proposal review UI', () => {
  it('renders line hunks and successful verification evidence instead of whole-file panes', () => {
    const html = renderToStaticMarkup(
      createElement(ProposalReview, {
        proposal,
        busy: false,
        verification: {
          state: 'verified',
          evidence: {
            proposalId: 'proposal-1',
            state: 'verified',
            diagnostics: [],
            logSummary: 'Compilation completed',
            verifiedAt: Date.UTC(2026, 7, 12),
          },
        },
        onConfirm: () => undefined,
        onVerifySelection: () => undefined,
        onCancel: () => undefined,
      }),
    )

    expect(html).toContain('Verified')
    expect(html).toContain('Compilation completed')
    expect(html).toContain('1 addition')
    expect(html).toContain('1 removal')
    expect(html).toContain('diff-line-remove')
    expect(html).toContain('diff-line-add')
    expect(html).not.toContain('proposal-before')
    expect(html).toContain('Apply verified changes')
  })

  it('explains failed, unverifiable, and rejected verification actions', () => {
    const render = (verification: unknown) =>
      renderToStaticMarkup(
        createElement(ProposalReview, {
          proposal,
          busy: false,
          verification,
          onConfirm: () => undefined,
          onVerifySelection: () => undefined,
          onCancel: () => undefined,
        } as never),
      )

    expect(
      render({
        state: 'failed',
        evidence: {
          proposalId: 'proposal-1',
          state: 'failed',
          diagnostics: [
            {
              path: 'main.tex',
              line: 4,
              column: 2,
              severity: 'error',
              message: 'Undefined control sequence',
            },
          ],
          logSummary: 'tectonic failed',
          reason: 'Compilation failed',
          verifiedAt: 1,
        },
      }),
    ).toContain('Apply without successful verification')
    expect(
      render({
        state: 'unverifiable',
        evidence: {
          proposalId: 'proposal-1',
          state: 'unverifiable',
          diagnostics: [],
          logSummary: '',
          reason: 'Proposals that create new files cannot be verified in isolation',
          verifiedAt: 1,
        },
      }),
    ).toContain('cannot be verified in isolation')
    const rejected = render({
      state: 'rejected',
      code: 'LATEX_CONFLICT',
      message: 'Proposal baseline changed on disk',
    })
    expect(rejected).toContain('Verification rejected')
    expect(rejected).toContain('Save your work or regenerate')
    expect(rejected).toContain('disabled=""')
  })

  it('visually distinguishes verified, failed, and neutral diff evidence', () => {
    const styles = readFileSync(new URL('../src/renderer/styles.css', import.meta.url), 'utf8')
    expect(styles).toMatch(/\.verification-verified\s*{[^}]*#[0-9a-f]{6}/s)
    expect(styles).toMatch(/\.verification-failed\s*{[^}]*#[0-9a-f]{6}/s)
    expect(styles).toMatch(/\.diff-line-add\s*{[^}]*background:/s)
    expect(styles).toMatch(/\.diff-line-remove\s*{[^}]*background:/s)
  })
})
import { readFileSync } from 'node:fs'
