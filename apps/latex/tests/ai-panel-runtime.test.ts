// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { AiPanel } from '../src/renderer/ai/AiPanel.js'

function mount(): { container: HTMLElement; root: Root } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(createElement(AiPanel, { projectId: 'project-1' })))
  return { container, root }
}

function typeAndSend(container: HTMLElement, text: string): void {
  const textarea = container.querySelector<HTMLTextAreaElement>('textarea')!
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!
  act(() => {
    setter.call(textarea, text)
    textarea.dispatchEvent(new Event('input', { bubbles: true }))
  })
  act(() => container.querySelector<HTMLButtonElement>('.ai-send-btn')!.click())
}

function installApis(runtime: 'legacy' | 'codex') {
  let runtimeEvent = (_event: unknown): void => undefined
  let toolRequest = (_request: unknown): void => undefined
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
    resolveDirectoryChat: vi.fn(async () => ({
      ok: true as const,
      value: { projectId: 'stored-project', chatId: 'chat-1' },
    })),
    loadDirectoryChat: vi.fn(async () => ({ ok: true as const, value: [] })),
    appendDirectoryChat: vi.fn(async () => ({ ok: true as const, value: undefined })),
    onAiStream: vi.fn(() => vi.fn()),
    aiStream: vi.fn(async () => undefined),
    aiStreamCancel: vi.fn(async () => undefined),
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
    applyProposal: vi.fn(),
    undoProposal: vi.fn(async () => ({ ok: true as const, value: {} })),
  }
  const codexRuntime = {
    status: vi.fn(async () => ({ runtime, documentId: runtime === 'codex' ? 't7' : 't7' })),
    startTurn: vi.fn(async () => undefined),
    cancelTurn: vi.fn(async () => undefined),
    onEvent: vi.fn((handler: typeof runtimeEvent) => {
      runtimeEvent = handler
      return vi.fn()
    }),
  }
  const codexTools = {
    register: vi.fn(async () => ({ registered: true as const })),
    unregister: vi.fn(async () => undefined),
    respond: vi.fn(async () => true),
    onRequest: vi.fn((handler: typeof toolRequest) => {
      toolRequest = handler
      return vi.fn()
    }),
    onCancel: vi.fn((_handler: (cancel: unknown) => void) => vi.fn()),
  }
  Object.assign(window, {
    latexApi,
    wisworkCodexRuntime: codexRuntime,
    wisworkCodexTools: codexTools,
  })
  return {
    latexApi,
    codexRuntime,
    codexTools,
    emitRuntime: (event: unknown) => runtimeEvent(event),
    requestTool: (request: unknown) => toolRequest(request),
  }
}

afterEach(() => {
  document.body.replaceChildren()
  vi.restoreAllMocks()
})

describe('LaTeX AI panel runtime selection', () => {
  it('uses Shell Codex start/events in codex mode and never invokes the legacy model transport', async () => {
    const f = installApis('codex')
    const { container, root } = mount()
    await vi.waitFor(() => expect(f.codexTools.register).toHaveBeenCalledOnce())

    typeAndSend(container, 'Improve the introduction')
    await vi.waitFor(() =>
      expect(f.codexRuntime.startTurn).toHaveBeenCalledWith({
        documentId: 't7',
        text: 'Improve the introduction',
      }),
    )
    expect(f.latexApi.aiStream).not.toHaveBeenCalled()

    act(() => f.emitRuntime({ documentId: 't7', event: { type: 'text', text: 'Working answer' } }))
    expect(container.textContent).toContain('Working answer')
    act(() =>
      f.emitRuntime({
        documentId: 't7',
        event: {
          type: 'done',
          result: { text: 'Final answer', toolExecutions: [], cancelled: false },
        },
      }),
    )
    expect(container.textContent).toContain('Final answer')
    expect(container.textContent).not.toContain('Working answer')

    act(() => root.unmount())
  })

  it('keeps the characterized AgentLoop transport as the authoritative legacy path', async () => {
    const f = installApis('legacy')
    const { container, root } = mount()
    await vi.waitFor(() => expect(f.codexRuntime.status).toHaveBeenCalledOnce())

    typeAndSend(container, 'Explain this project')
    await vi.waitFor(() => expect(f.latexApi.aiStream).toHaveBeenCalledOnce())
    expect(f.codexRuntime.startTurn).not.toHaveBeenCalled()
    expect(f.codexTools.register).not.toHaveBeenCalled()

    act(() => root.unmount())
  })

  it('shows an install-required Enhanced mode message before any turn can start', async () => {
    const f = installApis('codex')
    f.codexTools.register.mockRejectedValueOnce(new Error('enhanced_mode_install_required'))
    const { container, root } = mount()

    await vi.waitFor(() =>
      expect(container.textContent).toContain('Install Enhanced mode before use.'),
    )
    expect(f.codexRuntime.startTurn).not.toHaveBeenCalled()
    expect(f.latexApi.aiStream).not.toHaveBeenCalled()

    act(() => root.unmount())
  })

  it('uses the existing proposal review click before guarded Codex apply and exposes undo', async () => {
    const f = installApis('codex')
    const { container, root } = mount()
    await vi.waitFor(() => expect(f.codexTools.register).toHaveBeenCalledOnce())
    const call = {
      id: 'call-1',
      name: 'propose_project_edits',
      input: { files: [{ path: 'main.tex', afterText: 'after' }] },
    }
    act(() =>
      f.requestTool({
        type: 'approval',
        requestId: 'approval-1',
        documentId: 't7',
        call,
        expectedRevision: 'revision-1',
      }),
    )
    await vi.waitFor(() => expect(container.textContent).toContain('Review AI changes'))
    expect(f.latexApi.executeCodexProposalMutation).not.toHaveBeenCalled()
    act(() =>
      Array.from(container.querySelectorAll('button'))
        .find((button) => button.textContent === 'Confirm selected changes')!
        .click(),
    )
    await vi.waitFor(() =>
      expect(f.codexTools.respond).toHaveBeenCalledWith(
        expect.objectContaining({ requestId: 'approval-1', approved: true }),
      ),
    )

    act(() =>
      f.requestTool({
        type: 'snapshot',
        requestId: 'snapshot-1',
        documentId: 't7',
        call,
        expectedRevision: 'revision-1',
      }),
    )
    await vi.waitFor(() => expect(f.latexApi.prepareCodexProposalMutation).toHaveBeenCalledOnce())
    act(() =>
      f.requestTool({
        type: 'executeMutation',
        requestId: 'execute-1',
        documentId: 't7',
        call,
        guard: { expectedRevision: 'revision-1', snapshotId: 'snapshot-1' },
      }),
    )
    await vi.waitFor(() => expect(f.latexApi.executeCodexProposalMutation).toHaveBeenCalledOnce())
    await vi.waitFor(() => expect(container.textContent).toContain('Changes applied and compiled'))
    expect(container.textContent).toContain('Undo AI changes')

    act(() => root.unmount())
  })

  it('keeps stop, crash errors, and teardown bounded without falling back mid-turn', async () => {
    const f = installApis('codex')
    const { container, root } = mount()
    await vi.waitFor(() => expect(f.codexTools.register).toHaveBeenCalledOnce())
    typeAndSend(container, 'Long task')
    await vi.waitFor(() => expect(f.codexRuntime.startTurn).toHaveBeenCalledOnce())

    act(() =>
      Array.from(container.querySelectorAll<HTMLButtonElement>('button'))
        .find((button) => button.textContent === 'Stop')!
        .click(),
    )
    expect(f.codexRuntime.cancelTurn).toHaveBeenCalledWith('t7')
    expect(container.textContent).toContain('Stopped.')
    expect(f.latexApi.aiStream).not.toHaveBeenCalled()

    act(() =>
      f.emitRuntime({
        documentId: 't7',
        event: { type: 'error', code: 'enhanced_mode_stopped', message: 'Enhanced mode stopped.' },
      }),
    )
    expect(container.textContent).toContain('Enhanced mode stopped.')
    expect(f.latexApi.aiStream).not.toHaveBeenCalled()
    await vi.waitFor(() => expect(f.codexTools.unregister).toHaveBeenCalledTimes(1))

    act(() => root.unmount())
    await vi.waitFor(() => expect(f.codexTools.unregister).toHaveBeenCalledTimes(1))
  })
})
