import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AgentSkill, AgentToolCall, ToolExecution } from '@wiswork/agent-core'
import {
  DocumentToolRouter,
  type ToolSessionCredentials,
  type ToolSessionRegistration,
} from '@wiswork/codex-bridge'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type {
  CodexToolApi,
  CodexToolCancel,
  CodexToolRegistrationRequest,
  CodexToolRequest,
  CodexToolResponse,
} from '../../shell/src/shared/codex-api.js'
import { ProjectSessionRegistry } from '../src/main/project-session.js'
import { LatexCodexToolSession } from '../src/renderer/ai/codex-tool-session.js'
import { createLatexSkill } from '../src/renderer/ai/latex-skill.js'

type SuccessResponse = Extract<CodexToolResponse, { ok: true }>
type WithoutTransport<T> = T extends unknown ? Omit<T, 'requestId' | 'documentId'> : never
type OutboundRequest = WithoutTransport<CodexToolRequest>

class DriverToolApi implements CodexToolApi {
  readonly router = new DocumentToolRouter()
  directCalls = 0
  #credentials: ToolSessionCredentials | undefined
  #requestHandler = (_request: CodexToolRequest): void => undefined
  #cancelHandler = (_cancel: CodexToolCancel): void => undefined
  #pending = new Map<
    string,
    { resolve(value: SuccessResponse): void; reject(error: Error): void }
  >()
  #nextRequest = 1

  async register(request: CodexToolRegistrationRequest): Promise<{ registered: true }> {
    const skill: AgentSkill = {
      ...request.skill,
      tools: [...request.skill.tools],
      executeTool: (call, signal) =>
        this.#request(request.documentId, { type: 'execute', call }, 'execution', signal).then(
          (response) => response.execution,
        ),
    }
    const registration: ToolSessionRegistration = {
      skill,
      policy: request.policy,
      isOpen: () => this.#credentials !== undefined,
      getRevision: (signal) =>
        this.#request(request.documentId, { type: 'revision' }, 'revision', signal).then(
          (response) => response.revision,
        ),
      requestApproval: (call, expectedRevision, signal) =>
        this.#request(
          request.documentId,
          { type: 'approval', call, expectedRevision },
          'approval',
          signal,
        ).then((response) => response.approved),
      captureSnapshot: (call, expectedRevision, signal) =>
        this.#request(
          request.documentId,
          { type: 'snapshot', call, expectedRevision },
          'snapshot',
          signal,
        ).then((response) => response.snapshotId),
      discardSnapshot: (call, snapshotId) =>
        this.#request(
          request.documentId,
          { type: 'discardSnapshot', call, snapshotId },
          'discard',
        ).then(() => undefined),
      executeMutation: (call, guard, signal) =>
        this.#request(
          request.documentId,
          { type: 'executeMutation', call, guard },
          'execution',
          signal,
        ).then((response) => response.execution),
      validateMutation: async () => undefined,
    }
    this.#credentials = this.router.register(registration)
    return { registered: true }
  }

  async unregister(): Promise<void> {
    if (this.#credentials) this.router.close(this.#credentials)
    this.#credentials = undefined
  }

  async respond(response: CodexToolResponse): Promise<boolean> {
    const pending = this.#pending.get(response.requestId)
    if (!pending) return false
    this.#pending.delete(response.requestId)
    if (!response.ok) pending.reject(new Error(response.code))
    else pending.resolve(response)
    return true
  }

  onRequest(handler: (request: CodexToolRequest) => void): () => void {
    this.#requestHandler = handler
    return () => (this.#requestHandler = () => undefined)
  }

  onCancel(handler: (cancel: CodexToolCancel) => void): () => void {
    this.#cancelHandler = handler
    return () => (this.#cancelHandler = () => undefined)
  }

  async call(call: AgentToolCall): Promise<ToolExecution> {
    if (!this.#credentials) throw new Error('driver not registered')
    this.directCalls += 1
    return this.router.callTool(this.#credentials, call)
  }

  #request<TType extends SuccessResponse['type']>(
    documentId: string,
    request: OutboundRequest,
    expectedType: TType,
    signal?: AbortSignal,
  ): Promise<Extract<SuccessResponse, { type: TType }>> {
    const requestId = `driver-${this.#nextRequest++}`
    return new Promise((resolve, reject) => {
      const onAbort = () => {
        this.#cancelHandler({ requestId, documentId })
        reject(new Error('tool_cancelled'))
      }
      signal?.addEventListener('abort', onAbort, { once: true })
      this.#pending.set(requestId, {
        resolve: (response) => {
          signal?.removeEventListener('abort', onAbort)
          if (response.type !== expectedType) reject(new Error('unexpected response'))
          else resolve(response as never)
        },
        reject,
      })
      this.#requestHandler({ ...request, requestId, documentId } as CodexToolRequest)
    })
  }
}

describe('LaTeX Codex one-call workflow driver', () => {
  const roots: string[] = []

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })))
  })

  it('reads a multi-file project, waits for review, applies once, compiles, and undoes', async () => {
    const root = await mkdtemp(join(tmpdir(), 'latex-codex-driver-'))
    roots.push(root)
    const projectRoot = join(root, 'project')
    await mkdir(projectRoot)
    await writeFile(join(projectRoot, 'main.tex'), 'before main')
    await writeFile(join(projectRoot, 'chapter.tex'), 'before chapter')
    const session = await new ProjectSessionRegistry({
      watch: () => ({ close() {} }),
      compilerRuntime: { tectonicPath: '/fixed/tectonic', userDataPath: root },
    }).attach(51, projectRoot)
    await session.readText('main.tex')
    await session.readText('chapter.tex')
    const compile = vi.spyOn(session, 'compile').mockImplementation(async (revision) => ({
      revision,
      pdfUrl: null,
      diagnostics: [],
      log: '',
    }))
    let approve!: (selection: ReadonlySet<string> | null) => void
    const approvalPromise = new Promise<ReadonlySet<string> | null>(
      (resolve) => (approve = resolve),
    )
    let markReviewStarted!: () => void
    const reviewStarted = new Promise<void>((resolve) => (markReviewStarted = resolve))

    const domain = {
      listProjectFiles: async () => ({
        ok: true as const,
        value: await session.listProjectFilesForAi(),
      }),
      searchProjectText: async (request: { query: string; maxResults: number }) => ({
        ok: true as const,
        value: await session.searchProjectTextForAi(request.query, request.maxResults),
      }),
      readProjectText: async (request: { path: string; offset: number; maxChars: number }) => ({
        ok: true as const,
        value: await session.readProjectTextForAi(request.path, request.offset, request.maxChars),
      }),
      getCompileDiagnostics: async () => ({
        ok: true as const,
        value: session.getCompileDiagnosticsForAi(),
      }),
      compileProjectForAi: async () => ({ ok: true as const, value: await session.compileForAi() }),
      proposeProjectEdits: async (request: {
        files: Array<{ path: string; afterText: string }>
      }) => ({ ok: true as const, value: await session.createEditProposal(request.files) }),
      getProposal: async (request: { proposalId: string }) => ({
        ok: true as const,
        value: session.getProposal(request.proposalId),
      }),
      discardProposal: async (request: { proposalId: string }) => ({
        ok: true as const,
        value: await session.discardEditProposal(request.proposalId),
      }),
      getCodexMutationRevision: async () => ({
        ok: true as const,
        value: { revision: session.getCodexMutationRevision() },
      }),
      prepareCodexProposalMutation: async (
        request: Parameters<typeof session.prepareCodexProposalMutation>[0],
      ) => ({
        ok: true as const,
        value: await session.prepareCodexProposalMutation(request),
      }),
      executeCodexProposalMutation: async (
        request: Parameters<typeof session.executeCodexProposalMutation>[0],
      ) => ({
        ok: true as const,
        value: await session.executeCodexProposalMutation(request),
      }),
      discardCodexProposalMutation: async (
        request: Parameters<typeof session.discardCodexProposalMutation>[0],
      ) => ({
        ok: true as const,
        value: await session.discardCodexProposalMutation(request),
      }),
    }
    const transport = new DriverToolApi()
    const adapter = new LatexCodexToolSession({
      documentId: 't7',
      projectId: session.projectId,
      skill: createLatexSkill(domain, () => session.projectId),
      tools: transport,
      domain,
      requestReview: async () => {
        markReviewStarted()
        return approvalPromise
      },
      onApplied: () => undefined,
    })
    await adapter.start()
    const call = transport.call({
      id: 'one-direct-call',
      name: 'propose_project_edits',
      input: {
        files: [
          { path: 'main.tex', afterText: 'after main' },
          { path: 'chapter.tex', afterText: 'after chapter' },
        ],
      },
    })

    await reviewStarted
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('before main')
    expect(await readFile(join(projectRoot, 'chapter.tex'), 'utf8')).toBe('before chapter')
    approve(new Set(['main.tex', 'chapter.tex']))
    const execution = await call
    expect(execution).toMatchObject({ mutated: true })
    expect(execution.isError).not.toBe(true)
    expect(transport.directCalls).toBe(1)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('after main')
    expect(await readFile(join(projectRoot, 'chapter.tex'), 'utf8')).toBe('after chapter')
    expect(compile).toHaveBeenCalledTimes(1)

    const result = JSON.parse(execution.output) as { snapshotId: string }
    await session.undoConfirmedProposal(result.snapshotId)
    expect(await readFile(join(projectRoot, 'main.tex'), 'utf8')).toBe('before main')
    expect(await readFile(join(projectRoot, 'chapter.tex'), 'utf8')).toBe('before chapter')
    expect(compile).toHaveBeenCalledTimes(2)
    await adapter.close()
  })
})
