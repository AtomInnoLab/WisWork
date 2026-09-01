import { describe, expect, it } from 'vitest'
import type { ToolExecution } from '@wiswork/agent-core'
import { runEnhancedGolden } from '../../../packages/agent-runtime/src/production-golden'
import { createHostGoldenBridge } from '../../../packages/agent-runtime/tests/host-golden-bridge'
import { createLatexSkill } from '../src/renderer/ai/latex-skill'
import { ProposalWorkflow, validateUndoProposal } from '../src/renderer/ai/proposal-workflow'
import { loadProposalForReview } from '../src/renderer/ai/proposal-review'

describe('LaTeX production Enhanced golden', () => {
  it('runs proposal review, apply compile/readback and undo through the production skill', async () => {
    let source = 'before'
    const proposal = {
      id: 'latex-golden-proposal',
      projectId: 'latex-golden-project',
      expiresAt: Date.now() + 60_000,
      files: [
        {
          path: 'main.tex',
          beforeText: 'before',
          beforeSha256: 'b'.repeat(64),
          afterText: 'after',
        },
      ],
    }
    const verification = {
      proposalId: proposal.id,
      state: 'verified' as const,
      diagnostics: [],
      logSummary: 'clean',
      verifiedAt: Date.now(),
    }
    const workflow = new ProposalWorkflow(proposal.projectId, {
      verify: async () => ({ ok: true, value: verification }),
      create: async () => ({ ok: true, value: proposal }),
      apply: async () => {
        source = 'after'
        return {
          ok: true,
          value: {
            proposalId: proposal.id,
            snapshotId: 'a'.repeat(32),
            compile: {
              ok: true,
              result: { revision: 2, pdfUrl: null, diagnostics: [], log: 'ok' },
            },
          },
        }
      },
      refresh: () => undefined,
    })
    const skill = createLatexSkill(
      {
        listProjectFiles: async () => ({ ok: true, value: ['main.tex'] }),
        searchProjectText: async () => ({ ok: true, value: [] }),
        readProjectText: async () => ({ ok: true, value: source }),
        getCompileDiagnostics: async () => ({ ok: true, value: [] }),
        compileProjectForAi: async () => ({ ok: true, value: { diagnostics: [] } }),
        proposeProjectEdits: async () => ({
          ok: true,
          value: { proposalId: proposal.id, expiresAt: proposal.expiresAt, fileCount: 1 },
        }),
      },
      () => proposal.projectId,
    )
    const call = {
      id: 'latex-golden-call',
      name: 'propose_project_edits',
      input: { files: [{ path: 'main.tex', afterText: 'after' }] },
    }
    const result = await runEnhancedGolden('latex', {
      documentId: 'latex-golden-document',
      generation: 1,
      instruction: 'Update main.tex',
      bridge: createHostGoldenBridge({ documentId: 'latex-golden-document', generation: 1, call }),
      skill,
      confirm: async (execution: ToolExecution) => {
        const reviewed = await loadProposalForReview(
          execution.output,
          proposal.projectId,
          async () => ({ ok: true, value: proposal }),
        )
        await workflow.setProposal(reviewed)
        await workflow.primaryAction()
        return { mutationReceiptId: workflow.state.snapshotId! }
      },
      readback: async () => ({
        status:
          source === 'after' && workflow.state.status?.startsWith('Changes applied.')
            ? 'verified'
            : 'failed',
      }),
      rollback: async (confirmation) => {
        const undo = validateUndoProposal(
          {
            snapshotId: confirmation.mutationReceiptId,
            restored: true,
            compile: {
              ok: true,
              result: { revision: 3, pdfUrl: null, diagnostics: [], log: 'undo ok' },
            },
          },
          confirmation.mutationReceiptId,
        )
        source = 'before'
        return { status: undo.restored ? ('restored' as const) : ('failed' as const) }
      },
    })
    expect(result.verification).toEqual({ status: 'verified' })
    expect(result.rollback).toEqual({ status: 'restored' })
    expect(source).toBe('before')
    console.log(
      'ENHANCED_GOLDEN_REPORT',
      JSON.stringify({
        host: 'latex',
        verification: result.verification.status,
        rollback: result.rollback.status,
      }),
    )
  })
})
