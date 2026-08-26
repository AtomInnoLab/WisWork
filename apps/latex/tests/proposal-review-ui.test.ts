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
        selection: new Set(['main.tex']),
        busy: false,
        riskArmed: false,
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
        onSelectionChange: () => undefined,
        onPrimaryAction: () => undefined,
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
    const render = (verification: unknown, riskArmed = false) =>
      renderToStaticMarkup(
        createElement(ProposalReview, {
          proposal,
          selection: new Set(['main.tex']),
          busy: false,
          riskArmed,
          verification,
          onSelectionChange: () => undefined,
          onPrimaryAction: () => undefined,
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
    ).toContain('Review risk')
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

    const armed = render(
      {
        state: 'failed',
        evidence: {
          proposalId: 'proposal-1',
          state: 'failed',
          diagnostics: [],
          logSummary: '',
          reason: 'Compilation failed',
          verifiedAt: 1,
        },
      },
      true,
    )
    expect(armed).toContain('Apply unverified changes')
    expect(armed).toContain('aria-describedby="proposal-risk-warning"')
  })

  it('visually distinguishes verified, failed, and neutral diff evidence in both themes', () => {
    const styles = readFileSync(new URL('../src/renderer/styles.css', import.meta.url), 'utf8')
    expect(styles).toMatch(/\.verification-verified\s*{[^}]*var\(--latex-proposal-success-bg\)/s)
    expect(styles).toMatch(/\.verification-failed\s*{[^}]*var\(--latex-proposal-danger-bg\)/s)
    expect(styles).toMatch(/\.diff-line-add\s*{[^}]*var\(--latex-proposal-success-bg\)/s)
    expect(styles).toMatch(/\.diff-line-remove\s*{[^}]*var\(--latex-proposal-danger-bg\)/s)

    const lightSuccess = styles.match(
      /:root\s*{[^}]*--latex-proposal-success-bg:\s*(#[0-9a-f]{6})/s,
    )?.[1]
    const darkSuccess = styles.match(
      /\[data-theme='dark'\]\s*{[^}]*--latex-proposal-success-bg:\s*(#[0-9a-f]{6})/s,
    )?.[1]
    expect(lightSuccess).toBeTruthy()
    expect(darkSuccess).toBeTruthy()
    expect(darkSuccess).not.toBe(lightSuccess)
  })

  it('announces verification busy state and disables apply until it settles', () => {
    const html = renderToStaticMarkup(
      createElement(ProposalReview, {
        proposal,
        selection: new Set(['main.tex']),
        busy: false,
        verification: { state: 'verifying' },
        riskArmed: false,
        onSelectionChange: () => undefined,
        onPrimaryAction: () => undefined,
        onCancel: () => undefined,
      }),
    )
    expect(html).toContain('aria-busy="true"')
    expect(html).toContain('role="status"')
    expect(html).toContain('Verifying…')
    expect(html).toContain('disabled=""')
  })

  it('explains when a change is beyond the bounded diff scan budget', () => {
    const longProposal = {
      ...proposal,
      files: [
        {
          ...proposal.files[0],
          beforeText: `${'same\n'.repeat(120_000)}before`,
          afterText: `${'same\n'.repeat(120_000)}after`,
        },
      ],
    }
    const html = renderToStaticMarkup(
      createElement(ProposalReview, {
        proposal: longProposal,
        selection: new Set(['main.tex']),
        busy: false,
        verification: { state: 'verifying' },
        riskArmed: false,
        onSelectionChange: () => undefined,
        onPrimaryAction: () => undefined,
        onCancel: () => undefined,
      }),
    )
    expect(html).toContain('Change location is beyond the preview budget')
  })
})
import { readFileSync } from 'node:fs'
