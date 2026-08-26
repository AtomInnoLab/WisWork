import { describe, expect, it, vi } from 'vitest'
import type {
  CodexToolCancel,
  CodexToolRegistrationRequest,
  CodexToolRequest,
  CodexToolResponse,
} from '../../shell/src/shared/codex-api.js'
import { createLatexSkill } from '../src/renderer/ai/latex-skill.js'
import {
  LATEX_CODEX_TOOL_POLICY,
  LatexCodexToolSession,
} from '../src/renderer/ai/codex-tool-session.js'

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((done) => (resolve = done))
  return { promise, resolve }
}

function fixture() {
  let onRequest = (_request: CodexToolRequest): void => undefined
  let onCancel = (_cancel: CodexToolCancel): void => undefined
  const responses: CodexToolResponse[] = []
  const review = {
    id: 'ai-proposal-1',
    projectId: 'project-1',
    expiresAt: Date.now() + 60_000,
    files: [
      {
        path: 'main.tex',
        beforeText: 'before',
        beforeSha256: 'before-hash',
        afterText: 'after',
      },
    ],
  }
  const latexApi = {
    listProjectFiles: vi.fn(async () => ({ ok: true as const, value: { files: ['main.tex'] } })),
    searchProjectText: vi.fn(),
    readProjectText: vi.fn(),
    getCompileDiagnostics: vi.fn(),
    compileProjectForAi: vi.fn(),
    proposeProjectEdits: vi.fn(async () => ({ ok: true as const, value: review })),
    getProposal: vi.fn(async () => ({ ok: true as const, value: review })),
    discardProposal: vi.fn(async () => ({ ok: true as const, value: undefined })),
    getCodexMutationRevision: vi.fn(async () => ({
      ok: true as const,
      value: { revision: 'revision-1' },
    })),
    prepareCodexProposalMutation: vi.fn(async () => ({
      ok: true as const,
      value: { preparationId: 'preparation-1', snapshotId: 'snapshot-1' },
    })),
    executeCodexProposalMutation: vi.fn(async () => ({
      ok: true as const,
      value: {
        proposalId: 'ai-proposal-1',
        snapshotId: 'snapshot-1',
        compile: { ok: true, result: { diagnostics: [] } },
      },
    })),
    discardCodexProposalMutation: vi.fn(async () => ({
      ok: true as const,
      value: undefined,
    })),
  }
  const tools = {
    register: vi.fn(async (_request: CodexToolRegistrationRequest) => ({
      registered: true as const,
    })),
    unregister: vi.fn(async () => undefined),
    respond: vi.fn(async (response: CodexToolResponse) => {
      responses.push(response)
      return true
    }),
    onRequest: vi.fn((handler: typeof onRequest) => {
      onRequest = handler
      return vi.fn()
    }),
    onCancel: vi.fn((handler: typeof onCancel) => {
      onCancel = handler
      return vi.fn()
    }),
  }
  const requestReview = vi.fn(
    async (_proposal: unknown, _signal: AbortSignal): Promise<ReadonlySet<string> | null> =>
      new Set(['main.tex']),
  )
  const onApplied = vi.fn()
  const session = new LatexCodexToolSession({
    documentId: 't7',
    projectId: 'project-1',
    skill: createLatexSkill(latexApi, () => 'project-1'),
    tools,
    domain: latexApi,
    requestReview,
    onApplied,
  })
  return {
    session,
    latexApi,
    tools,
    responses,
    requestReview,
    onApplied,
    request: (request: CodexToolRequest) => onRequest(request),
    cancel: (cancel: CodexToolCancel) => onCancel(cancel),
  }
}

async function waitForResponse(f: ReturnType<typeof fixture>, requestId: string) {
  await vi.waitFor(() =>
    expect(f.responses.some((response) => response.requestId === requestId)).toBe(true),
  )
  return f.responses.find((response) => response.requestId === requestId)!
}

describe('LaTeX Codex renderer tool session', () => {
  it('registers the authoritative document with an explicit mutate-only proposal policy', async () => {
    const f = fixture()
    await f.session.start()

    expect(f.tools.register).toHaveBeenCalledWith({
      documentId: 't7',
      skill: expect.objectContaining({ id: 'latex-project' }),
      policy: LATEX_CODEX_TOOL_POLICY,
    })
    expect(LATEX_CODEX_TOOL_POLICY).toEqual({
      list_project_files: 'read',
      search_project_text: 'read',
      read_project_text: 'read',
      get_compile_diagnostics: 'read',
      compile_project: 'read',
      propose_project_edits: 'mutate',
    })
  })

  it('executes read tools automatically with the separate LaTeX project identity', async () => {
    const f = fixture()
    await f.session.start()
    f.request({
      type: 'execute',
      requestId: 'read-1',
      documentId: 't7',
      call: { id: 'call-read', name: 'list_project_files', input: {} },
    })

    await expect(waitForResponse(f, 'read-1')).resolves.toMatchObject({
      ok: true,
      type: 'execution',
      execution: { mutated: false },
    })
    expect(f.latexApi.listProjectFiles).toHaveBeenCalledWith({ projectId: 'project-1' })
  })

  it('waits for proposal review, captures one real snapshot, then executes the bound mutation', async () => {
    const f = fixture()
    await f.session.start()
    const call = {
      id: 'call-mutate',
      name: 'propose_project_edits',
      input: { files: [{ path: 'main.tex', afterText: 'after' }] },
    }
    f.request({
      type: 'approval',
      requestId: 'approval-1',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    await expect(waitForResponse(f, 'approval-1')).resolves.toMatchObject({
      ok: true,
      type: 'approval',
      approved: true,
    })
    expect(f.requestReview).toHaveBeenCalledOnce()

    f.request({
      type: 'snapshot',
      requestId: 'snapshot-request',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    await expect(waitForResponse(f, 'snapshot-request')).resolves.toMatchObject({
      ok: true,
      type: 'snapshot',
      snapshotId: 'snapshot-1',
    })
    expect(f.latexApi.prepareCodexProposalMutation).toHaveBeenCalledWith({
      projectId: 'project-1',
      documentId: 't7',
      callId: 'call-mutate',
      proposalId: 'ai-proposal-1',
      expectedRevision: 'revision-1',
    })

    f.request({
      type: 'executeMutation',
      requestId: 'execute-1',
      documentId: 't7',
      call,
      guard: { expectedRevision: 'revision-1', snapshotId: 'snapshot-1' },
    })
    await expect(waitForResponse(f, 'execute-1')).resolves.toMatchObject({
      ok: true,
      type: 'execution',
      execution: { mutated: true },
    })
    expect(f.latexApi.executeCodexProposalMutation).toHaveBeenCalledWith({
      projectId: 'project-1',
      documentId: 't7',
      callId: 'call-mutate',
      preparationId: 'preparation-1',
      snapshotId: 'snapshot-1',
      expectedRevision: 'revision-1',
    })
    expect(f.onApplied).toHaveBeenCalledWith(expect.objectContaining({ snapshotId: 'snapshot-1' }))
  })

  it('denies safely and aborts review without preparing or executing a mutation', async () => {
    const f = fixture()
    const review = deferred<ReadonlySet<string> | null>()
    f.requestReview.mockReturnValueOnce(review.promise)
    await f.session.start()
    const call = {
      id: 'denied-call',
      name: 'propose_project_edits',
      input: { files: [{ path: 'main.tex', afterText: 'after' }] },
    }
    f.request({
      type: 'approval',
      requestId: 'approval-denied',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    review.resolve(null)

    await expect(waitForResponse(f, 'approval-denied')).resolves.toMatchObject({
      ok: true,
      type: 'approval',
      approved: false,
    })
    expect(f.latexApi.prepareCodexProposalMutation).not.toHaveBeenCalled()
    expect(f.latexApi.executeCodexProposalMutation).not.toHaveBeenCalled()
  })

  it('discards a tentative snapshot exactly once on router cleanup and session close', async () => {
    const f = fixture()
    await f.session.start()
    const call = {
      id: 'cleanup-call',
      name: 'propose_project_edits',
      input: { files: [{ path: 'main.tex', afterText: 'after' }] },
    }
    f.request({
      type: 'approval',
      requestId: 'approval-cleanup',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    await waitForResponse(f, 'approval-cleanup')
    f.request({
      type: 'snapshot',
      requestId: 'snapshot-cleanup',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    await waitForResponse(f, 'snapshot-cleanup')
    f.request({
      type: 'discardSnapshot',
      requestId: 'discard-1',
      documentId: 't7',
      call,
      snapshotId: 'snapshot-1',
    })
    await expect(waitForResponse(f, 'discard-1')).resolves.toMatchObject({
      ok: true,
      type: 'discard',
    })
    expect(f.latexApi.discardCodexProposalMutation).toHaveBeenCalledTimes(1)

    await f.session.close()
    await f.session.close()
    expect(f.latexApi.discardCodexProposalMutation).toHaveBeenCalledTimes(1)
    expect(f.tools.unregister).toHaveBeenCalledTimes(1)
  })

  it('keeps a rejected guarded preparation available for the router cleanup request', async () => {
    const f = fixture()
    await f.session.start()
    const call = {
      id: 'failed-execute-call',
      name: 'propose_project_edits',
      input: { files: [{ path: 'main.tex', afterText: 'after' }] },
    }
    f.request({
      type: 'approval',
      requestId: 'failed-approval',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    await waitForResponse(f, 'failed-approval')
    f.request({
      type: 'snapshot',
      requestId: 'failed-snapshot',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    await waitForResponse(f, 'failed-snapshot')
    f.latexApi.executeCodexProposalMutation.mockResolvedValueOnce({
      ok: false as const,
      error: { message: 'bounded execute failure' },
    } as never)

    f.request({
      type: 'executeMutation',
      requestId: 'failed-execute',
      documentId: 't7',
      call,
      guard: { expectedRevision: 'revision-1', snapshotId: 'snapshot-1' },
    })
    await expect(waitForResponse(f, 'failed-execute')).resolves.toMatchObject({ ok: false })
    f.request({
      type: 'discardSnapshot',
      requestId: 'failed-discard',
      documentId: 't7',
      call,
      snapshotId: 'snapshot-1',
    })

    await expect(waitForResponse(f, 'failed-discard')).resolves.toMatchObject({
      ok: true,
      type: 'discard',
    })
    expect(f.latexApi.discardCodexProposalMutation).toHaveBeenCalledTimes(1)
  })

  it('discards a snapshot that finishes materializing while the renderer session closes', async () => {
    const f = fixture()
    const preparing = deferred<{
      ok: true
      value: { preparationId: string; snapshotId: string }
    }>()
    f.latexApi.prepareCodexProposalMutation.mockReturnValueOnce(preparing.promise)
    await f.session.start()
    const call = {
      id: 'closing-capture-call',
      name: 'propose_project_edits',
      input: { files: [{ path: 'main.tex', afterText: 'after' }] },
    }
    f.request({
      type: 'approval',
      requestId: 'closing-approval',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    await waitForResponse(f, 'closing-approval')
    f.request({
      type: 'snapshot',
      requestId: 'closing-snapshot',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    const closing = f.session.close()
    preparing.resolve({
      ok: true,
      value: { preparationId: 'preparation-1', snapshotId: 'snapshot-1' },
    })

    await closing
    await vi.waitFor(() => expect(f.latexApi.discardCodexProposalMutation).toHaveBeenCalledTimes(1))
    await expect(waitForResponse(f, 'closing-snapshot')).resolves.toMatchObject({ ok: false })
  })

  it('rejects cross-document requests and forwards cancellation into pending review', async () => {
    const f = fixture()
    const review = deferred<ReadonlySet<string> | null>()
    let reviewSignal: AbortSignal | undefined
    f.requestReview.mockImplementationOnce((_proposal, signal) => {
      reviewSignal = signal
      return review.promise
    })
    await f.session.start()
    const call = {
      id: 'cancel-call',
      name: 'propose_project_edits',
      input: { files: [{ path: 'main.tex', afterText: 'after' }] },
    }
    f.request({
      type: 'approval',
      requestId: 'approval-cancel',
      documentId: 't7',
      call,
      expectedRevision: 'revision-1',
    })
    await vi.waitFor(() => expect(reviewSignal).toBeDefined())
    f.cancel({ requestId: 'approval-cancel', documentId: 't7' })
    expect(reviewSignal?.aborted).toBe(true)
    review.resolve(null)
    await expect(waitForResponse(f, 'approval-cancel')).resolves.toMatchObject({
      ok: false,
      code: 'tool_cancelled',
    })

    f.request({
      type: 'revision',
      requestId: 'forged',
      documentId: 'other-document',
    })
    await new Promise((resolve) => setImmediate(resolve))
    expect(f.responses.some((response) => response.requestId === 'forged')).toBe(false)
  })
})
