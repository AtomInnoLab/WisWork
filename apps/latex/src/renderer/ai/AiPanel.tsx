import { useEffect, useRef, useState } from 'react'
import { AgentLoop, type ToolExecutedEvent } from '@wiswork/agent-core'
import type { AgentEvent } from '@wiswork/agent-runtime'
import { AiComposer, AiTypingIndicator, Markdown, WisWorkAppMark } from '@wiswork/ui'
import { createLatexSkill } from './latex-skill.js'
import {
  loadProposalForReview,
  proposalForSelection,
  type ReviewProposal,
} from './proposal-review.js'
import { ProposalReview } from './ProposalReview.js'
import { LatexCodexToolSession, type LatexCodexMutationResult } from './codex-tool-session.js'
import { createLatexTransport } from './transport.js'

function enhancedModeInstallRequired(error: unknown): boolean {
  if (error instanceof Error && error.message.includes('enhanced_mode_install_required')) {
    return true
  }
  return (
    typeof error === 'object' &&
    error !== null &&
    'code' in error &&
    error.code === 'enhanced_mode_install_required'
  )
}

const E2E_PROPOSAL_TEXT = String.raw`\documentclass{article}
\begin{document}
AI-confirmed WisWork
\end{document}`

interface ChatEntry {
  role: 'user' | 'assistant'
  text: string
}

interface PanelRuntime {
  run(text: string): void | Promise<void>
  cancel(): void | Promise<void>
}

const PANEL_WIDTH_KEY = 'latex-ai-panel-width'
const PANEL_WIDTH_DEFAULT = 360
const PANEL_WIDTH_MIN = 300

function clampPanelWidth(width: number): number {
  const viewportWidth = typeof window === 'undefined' ? 1440 : window.innerWidth
  return Math.min(Math.max(width, PANEL_WIDTH_MIN), Math.min(620, viewportWidth * 0.55))
}

function loadPanelWidth(): number {
  if (typeof localStorage === 'undefined' || typeof localStorage.getItem !== 'function') {
    return PANEL_WIDTH_DEFAULT
  }
  const saved = Number(localStorage.getItem(PANEL_WIDTH_KEY))
  return Number.isFinite(saved) && saved > 0 ? clampPanelWidth(saved) : PANEL_WIDTH_DEFAULT
}

export function AiPanel({
  projectId,
  disabled = false,
  onProjectFilesChanged,
  open = true,
  onExpand,
  onCollapse,
}: {
  projectId: string
  disabled?: boolean
  onProjectFilesChanged?: () => void | Promise<void>
  open?: boolean
  onExpand?: () => void
  onCollapse?: () => void
}) {
  const [input, setInput] = useState('')
  const [text, setText] = useState('')
  const [busy, setBusy] = useState(false)
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [proposal, setProposal] = useState<ReviewProposal | null>(null)
  const [snapshotId, setSnapshotId] = useState<string | null>(null)
  const [status, setStatus] = useState<string | null>(null)
  const [activity, setActivity] = useState<string | null>(null)
  const [ready, setReady] = useState(false)
  const [reviewAwaiting, setReviewAwaiting] = useState(false)
  const [panelWidth, setPanelWidth] = useState(loadPanelWidth)
  const [resizing, setResizing] = useState(false)
  const projectRef = useRef(projectId)
  projectRef.current = projectId
  const loopRef = useRef<AgentLoop | null>(null)
  const runtimeRef = useRef<PanelRuntime | null>(null)
  const restoredChatRef = useRef<ChatEntry[] | null>(null)
  const approvalRef = useRef<
    | {
        proposalId: string
        resolve(selected: ReadonlySet<string> | null): void
      }
    | undefined
  >(undefined)
  const chatIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)
  const e2eProposalLoaded = useRef<string | null>(null)

  useEffect(() => {
    let disposed = false
    let removeRuntimeEvents: (() => void) | undefined
    let codexTools: LatexCodexToolSession | undefined

    const onDone = (result: { text: string }) => {
      setBusy(false)
      setActivity(null)
      setText('')
      if (result.text) setChat((current) => [...current, { role: 'assistant', text: result.text }])
      const ids = chatIdsRef.current
      if (ids)
        void window.latexApi.appendDirectoryChat({
          projectId: projectRef.current,
          storeProjectId: ids.projectId,
          chatId: ids.chatId,
          role: 'assistant',
          text: result.text,
        })
    }
    const onError = (error: string) => {
      setStatus(error)
      setActivity(null)
      setBusy(false)
    }
    const onLegacyToolExecuted = ({ call, execution }: ToolExecutedEvent<unknown>) => {
      if (call.name !== 'propose_project_edits' || execution.isError) return
      const projectId = projectRef.current
      void loadProposalForReview(execution.output, projectId, (request) =>
        window.latexApi.getProposal(request),
      )
        .then(setProposal)
        .catch((error: unknown) =>
          setStatus(error instanceof Error ? error.message : 'Unable to load proposal.'),
        )
    }
    const requestReview = (
      nextProposal: ReviewProposal,
      signal: AbortSignal,
    ): Promise<ReadonlySet<string> | null> => {
      approvalRef.current?.resolve(null)
      return new Promise((resolve) => {
        let settled = false
        const finish = (selected: ReadonlySet<string> | null) => {
          if (settled) return
          settled = true
          signal.removeEventListener('abort', onAbort)
          if (approvalRef.current?.proposalId === nextProposal.id) approvalRef.current = undefined
          setReviewAwaiting(false)
          if (selected === null) setProposal(null)
          resolve(selected)
        }
        const onAbort = () => finish(null)
        approvalRef.current = { proposalId: nextProposal.id, resolve: finish }
        setReviewAwaiting(true)
        setProposal(nextProposal)
        signal.addEventListener('abort', onAbort, { once: true })
        if (signal.aborted) onAbort()
      })
    }
    const onCodexApplied = async (result: LatexCodexMutationResult) => {
      await onProjectFilesChanged?.()
      setSnapshotId(result.snapshotId)
      setProposal(null)
      setStatus(
        result.compile.ok
          ? `Changes applied and compiled (${result.compile.result?.diagnostics.length ?? 0} diagnostics).`
          : `Changes applied; compile failed: ${result.compile.error ?? 'unknown error'}`,
      )
    }
    const onCodexEvent = (event: AgentEvent<unknown>) => {
      switch (event.type) {
        case 'text':
          setText(event.text)
          break
        case 'tool-start':
          setActivity(`Using ${event.call.name.slice(0, 80)}…`)
          break
        case 'tool-executed':
        case 'turn-end':
          setActivity(null)
          break
        case 'done':
          onDone(event.result)
          break
        case 'error':
          onError(event.message)
          if (event.code === 'enhanced_mode_stopped') {
            setReady(false)
            runtimeRef.current = null
            void codexTools?.close()
          }
          break
      }
    }

    const initialize = async () => {
      try {
        const runtimeStatus = await window.wisworkCodexRuntime.status()
        if (disposed) return
        if (runtimeStatus.runtime === 'legacy') {
          const loop = new AgentLoop({
            transport: createLatexTransport(),
            skill: createLatexSkill(window.latexApi, () => projectRef.current),
            events: {
              onText: setText,
              onToolExecuted: onLegacyToolExecuted,
              onDone,
              onError,
            },
          })
          if (restoredChatRef.current) loop.restore(restoredChatRef.current)
          loopRef.current = loop
          runtimeRef.current = {
            run: (instruction) => loop.run(instruction),
            cancel: () => loop.cancel(),
          }
          setReady(true)
          return
        }
        if (!runtimeStatus.documentId) throw new Error('codex_document_unavailable')
        const documentId = runtimeStatus.documentId
        removeRuntimeEvents = window.wisworkCodexRuntime.onEvent((message) => {
          if (message.documentId === documentId && !disposed) onCodexEvent(message.event)
        })
        codexTools = new LatexCodexToolSession({
          documentId,
          projectId: projectRef.current,
          skill: createLatexSkill(window.latexApi, () => projectRef.current),
          tools: window.wisworkCodexTools,
          domain: window.latexApi,
          requestReview,
          onApplied: onCodexApplied,
        })
        await codexTools.start()
        if (disposed) {
          await codexTools.close()
          return
        }
        runtimeRef.current = {
          run: (instruction) =>
            window.wisworkCodexRuntime.startTurn({ documentId, text: instruction }),
          cancel: () => window.wisworkCodexRuntime.cancelTurn(documentId),
        }
        setReady(true)
      } catch (error) {
        if (!disposed) {
          onError(
            enhancedModeInstallRequired(error)
              ? 'Install Enhanced mode before use.'
              : 'The selected AI runtime is unavailable.',
          )
        }
      }
    }
    void initialize()
    return () => {
      disposed = true
      approvalRef.current?.resolve(null)
      approvalRef.current = undefined
      void runtimeRef.current?.cancel()
      runtimeRef.current = null
      loopRef.current?.cancel()
      loopRef.current = null
      removeRuntimeEvents?.()
      void codexTools?.close()
    }
  }, [])

  useEffect(() => {
    void window.latexApi.resolveDirectoryChat({ projectId }).then(async (result) => {
      if (!result.ok) return
      chatIdsRef.current = result.value
      const loaded = await window.latexApi.loadDirectoryChat({
        projectId,
        storeProjectId: result.value.projectId,
        chatId: result.value.chatId,
        limit: 200,
      })
      if (loaded.ok) {
        const restored = loaded.value.map((message) => ({
          role: message.role,
          text: message.text,
        }))
        setChat(restored)
        restoredChatRef.current = restored
        loopRef.current?.restore(restored)
      }
    })
  }, [projectId])

  useEffect(() => {
    if (!resizing) return
    const move = (event: PointerEvent) => {
      const width = clampPanelWidth(window.innerWidth - event.clientX)
      setPanelWidth(width)
      localStorage.setItem(PANEL_WIDTH_KEY, String(Math.round(width)))
    }
    const stop = () => setResizing(false)
    window.addEventListener('pointermove', move)
    window.addEventListener('pointerup', stop, { once: true })
    return () => {
      window.removeEventListener('pointermove', move)
      window.removeEventListener('pointerup', stop)
    }
  }, [resizing])

  useEffect(() => {
    if (
      new URLSearchParams(window.location.search).get('e2eProposal') !== '1' ||
      e2eProposalLoaded.current === projectId
    ) {
      return
    }
    e2eProposalLoaded.current = projectId
    void window.latexApi
      .proposeProjectEdits({
        projectId,
        files: [{ path: 'main.tex', afterText: E2E_PROPOSAL_TEXT }],
      })
      .then((result) => {
        if (result.ok) setProposal(result.value)
        else setStatus(result.error.message)
      })
  }, [projectId])

  const send = () => {
    const instruction = input.trim()
    const runtime = runtimeRef.current
    if (!instruction || busy || disabled || !ready || !runtime) return
    setBusy(true)
    setStatus(null)
    setText('')
    setInput('')
    setChat((current) => [...current, { role: 'user', text: instruction }])
    const ids = chatIdsRef.current
    if (ids)
      void window.latexApi.appendDirectoryChat({
        projectId,
        storeProjectId: ids.projectId,
        chatId: ids.chatId,
        role: 'user',
        text: instruction,
      })
    try {
      void Promise.resolve(runtime.run(instruction)).catch(() => {
        setStatus('The selected AI runtime is unavailable.')
        setBusy(false)
      })
    } catch {
      setStatus('The selected AI runtime is unavailable.')
      setBusy(false)
    }
  }

  const cancel = () => {
    approvalRef.current?.resolve(null)
    void runtimeRef.current?.cancel()
    setBusy(false)
    setText('')
    setStatus('Stopped.')
  }

  const confirm = async (selected: ReadonlySet<string>) => {
    if (!proposal || disabled) return
    if (approvalRef.current?.proposalId === proposal.id) {
      approvalRef.current.resolve(selected)
      setProposal(null)
      return
    }
    setBusy(true)
    try {
      const owned = await proposalForSelection(proposal, selected, async (files) => {
        const result = await window.latexApi.proposeProjectEdits({ projectId, files })
        if (!result.ok) throw new Error(result.error.message)
        return result.value
      })
      const result = await window.latexApi.applyProposal({ projectId, proposalId: owned.id })
      if (!result.ok) throw new Error(result.error.message)
      const value = result.value as {
        snapshotId: string
        compile: { ok: boolean; error?: string; result?: { diagnostics: unknown[] } }
      }
      await onProjectFilesChanged?.()
      setSnapshotId(value.snapshotId)
      setProposal(null)
      setStatus(
        value.compile.ok
          ? `Changes applied and compiled (${value.compile.result?.diagnostics.length ?? 0} diagnostics).`
          : `Changes applied; compile failed: ${value.compile.error ?? 'unknown error'}`,
      )
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error))
    } finally {
      setBusy(false)
    }
  }

  const undo = async () => {
    if (!snapshotId || disabled) return
    setBusy(true)
    const result = await window.latexApi.undoProposal({ projectId, snapshotId })
    if (result.ok) {
      await onProjectFilesChanged?.()
      setSnapshotId(null)
      setStatus('AI changes were undone and the project was compiled again.')
    } else setStatus(result.error.message)
    setBusy(false)
  }

  // Keep the component mounted while collapsed so its transcript, draft and active run survive.
  if (!open) {
    return (
      <button className="latex-ai-rail" onClick={onExpand} aria-label="Open AI panel">
        <WisWorkAppMark className="ai-brand-icon" size={25} />
        <span>AI</span>
      </button>
    )
  }

  return (
    <div
      className={`latex-ai-dock${resizing ? ' ai-panel-resizing' : ''}`}
      style={{ width: panelWidth }}
    >
      <aside className="latex-ai-panel">
        <div
          className="ai-panel-resizer"
          role="separator"
          aria-label="Resize AI panel"
          onPointerDown={() => setResizing(true)}
        />
        <header className="ai-panel-header">
          <div className="ai-panel-title">
            <WisWorkAppMark className="ai-brand-icon" size={25} />
            <span>WisWork AI</span>
          </div>
          {onCollapse && (
            <button className="ai-header-btn" onClick={onCollapse} aria-label="Collapse AI panel">
              ›
            </button>
          )}
        </header>
        <div className="ai-chat" aria-live="polite">
          {chat.length === 0 && !text && !proposal && (
            <div className="ai-chat-empty">
              <div className="ai-chat-empty-title">Edit LaTeX with WisWork AI</div>
              <div className="ai-chat-empty-body">
                Ask questions, fix compilation issues, or propose project-wide changes.
              </div>
              <div className="ai-starter-list">
                {['Explain this document', 'Fix LaTeX errors', 'Improve the current section'].map(
                  (starter) => (
                    <button key={starter} className="ai-starter" onClick={() => setInput(starter)}>
                      {starter}
                    </button>
                  ),
                )}
              </div>
            </div>
          )}
          {chat.map((entry, index) => (
            <div key={`${entry.role}-${index}`} className={`ai-msg ai-msg-${entry.role}`}>
              {entry.role === 'assistant' ? (
                <div className="ai-md">
                  <Markdown text={entry.text} />
                </div>
              ) : (
                entry.text
              )}
            </div>
          ))}
          {text && (
            <div className="ai-msg ai-msg-assistant">
              <div className="ai-md">
                <Markdown text={text} />
              </div>
            </div>
          )}
          {busy && !text && <AiTypingIndicator label="WisWork is working" />}
          {activity && <div className="ai-status">{activity}</div>}
          {proposal && (
            <ProposalReview
              proposal={proposal}
              busy={(busy && !reviewAwaiting) || disabled}
              onConfirm={(selected) => void confirm(selected)}
              onCancel={() => {
                if (approvalRef.current?.proposalId === proposal.id)
                  approvalRef.current.resolve(null)
                setProposal(null)
              }}
            />
          )}
          {status && (
            <div className="ai-status" role="status">
              {status}
            </div>
          )}
          {snapshotId && (
            <button
              className="ai-undo-button"
              disabled={busy || disabled}
              onClick={() => void undo()}
            >
              Undo AI changes
            </button>
          )}
        </div>
        <div className="ai-composer-wrap">
          <AiComposer
            value={input}
            busy={busy || !ready}
            placeholder="Ask WisWork AI about this LaTeX project"
            hintIdle="Enter to send · Shift+Enter for new line"
            hintBusy="Working…"
            hintIdleTitle="Enter to send"
            sendLabel="Send"
            stopLabel="Stop"
            ariaLabel="Ask WisWork AI"
            onChange={setInput}
            onSend={send}
            onStop={cancel}
          />
        </div>
      </aside>
    </div>
  )
}
