import { useEffect, useRef, useState } from 'react'
import { AgentLoop } from '@wiswork/agent-core'
import { AiComposer, AiTypingIndicator, Markdown, WisWorkAppMark } from '@wiswork/ui'
import { createLatexSkill } from './latex-skill.js'
import {
  loadProposalForReview,
  proposalForSelection,
  type ReviewProposal,
} from './proposal-review.js'
import { ProposalReview } from './ProposalReview.js'
import { createLatexTransport } from './transport.js'

const E2E_PROPOSAL_TEXT = String.raw`\documentclass{article}
\begin{document}
AI-confirmed WisWork
\end{document}`

interface ChatEntry {
  role: 'user' | 'assistant'
  text: string
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
  onCollapse,
}: {
  projectId: string
  disabled?: boolean
  onProjectFilesChanged?: () => void | Promise<void>
  onCollapse?: () => void
}) {
  const [input, setInput] = useState('')
  const [text, setText] = useState('')
  const [busy, setBusy] = useState(false)
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [proposal, setProposal] = useState<ReviewProposal | null>(null)
  const [snapshotId, setSnapshotId] = useState<string | null>(null)
  const [status, setStatus] = useState<string | null>(null)
  const [panelWidth, setPanelWidth] = useState(loadPanelWidth)
  const [resizing, setResizing] = useState(false)
  const projectRef = useRef(projectId)
  projectRef.current = projectId
  const loopRef = useRef<AgentLoop | null>(null)
  const chatIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)
  const e2eProposalLoaded = useRef<string | null>(null)

  useEffect(() => {
    const loop = new AgentLoop({
      transport: createLatexTransport(),
      skill: createLatexSkill(window.latexApi, () => projectRef.current),
      events: {
        onText: setText,
        onToolExecuted: ({ call, execution }) => {
          if (call.name !== 'propose_project_edits' || execution.isError) return
          const projectId = projectRef.current
          void loadProposalForReview(execution.output, projectId, (request) =>
            window.latexApi.getProposal(request),
          )
            .then(setProposal)
            .catch((error: unknown) =>
              setStatus(error instanceof Error ? error.message : String(error)),
            )
        },
        onDone: (result) => {
          setBusy(false)
          setText('')
          if (result.text)
            setChat((current) => [...current, { role: 'assistant', text: result.text }])
          const ids = chatIdsRef.current
          if (ids)
            void window.latexApi.appendDirectoryChat({
              projectId: projectRef.current,
              storeProjectId: ids.projectId,
              chatId: ids.chatId,
              role: 'assistant',
              text: result.text,
            })
        },
        onError: (error) => {
          setStatus(error)
          setBusy(false)
        },
      },
    })
    loopRef.current = loop
    return () => {
      loop.cancel()
      loopRef.current = null
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
    if (!instruction || busy || disabled) return
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
    loopRef.current?.run(instruction)
  }

  const cancel = () => {
    loopRef.current?.cancel()
    setBusy(false)
    setText('')
    setStatus('Stopped.')
  }

  const confirm = async (selected: ReadonlySet<string>) => {
    if (!proposal || disabled) return
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
          {proposal && (
            <ProposalReview
              proposal={proposal}
              busy={busy || disabled}
              onConfirm={(selected) => void confirm(selected)}
              onCancel={() => setProposal(null)}
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
            busy={busy}
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
