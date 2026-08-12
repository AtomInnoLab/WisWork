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
import {
  normalizeAgentContext,
  serializeAgentPrompt,
  type AgentContext,
  type AgentContextKey,
} from './agent-context.js'
import {
  cancelRunningTimelineEntries,
  completeTimelineEntry,
  failRunningTimelineEntries,
  startTimelineEntry,
  type TaskTimelineEntry,
} from './task-timeline.js'
import {
  AgentPanelSession,
  type AgentRunScope,
  type AgentProjectScope,
} from './agent-panel-session.js'

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

function contextChips(context: AgentContext): Array<{ key: AgentContextKey; label: string }> {
  const normalized = normalizeAgentContext(context)
  return [
    ...(normalized.activeFile
      ? [
          {
            key: 'activeFile' as const,
            label: `${normalized.activeFile}${normalized.cursorLine ? `:${normalized.cursorLine}` : ''}`,
          },
        ]
      : []),
    ...(normalized.selection
      ? [
          {
            key: 'selection' as const,
            label: `Selection lines ${normalized.selection.startLine}–${normalized.selection.endLine}`,
          },
        ]
      : []),
    ...(normalized.diagnostic
      ? [
          {
            key: 'diagnostic' as const,
            label: `${normalized.diagnostic.severity === 'error' ? 'Error' : 'Warning'} at ${normalized.diagnostic.path}:${normalized.diagnostic.line}`,
          },
        ]
      : []),
  ]
}

function TaskTimeline({ entries }: { entries: readonly TaskTimelineEntry[] }) {
  if (!entries.length) return null
  return (
    <section className="agent-task-timeline" aria-label="Agent task timeline">
      <div className="agent-task-timeline-title">Task activity</div>
      <ol>
        {entries.map((entry) => (
          <li key={entry.id} className={`timeline-${entry.state}`}>
            <span className="timeline-state" aria-label={entry.state} />
            <div>
              <div className="timeline-label">{entry.label}</div>
              {entry.detail && (
                <details>
                  <summary>Details</summary>
                  <pre>{entry.detail}</pre>
                </details>
              )}
            </div>
          </li>
        ))}
      </ol>
    </section>
  )
}

export function AiPanel({
  projectId,
  disabled = false,
  onProjectFilesChanged,
  open = true,
  onExpand,
  onCollapse,
  context = {},
  onRemoveContext,
}: {
  projectId: string
  disabled?: boolean
  onProjectFilesChanged?: () => void | Promise<void>
  open?: boolean
  onExpand?: () => void
  onCollapse?: () => void
  context?: AgentContext
  onRemoveContext?: (key: AgentContextKey) => void
}) {
  const [input, setInput] = useState('')
  const [text, setText] = useState('')
  const [busy, setBusy] = useState(false)
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [proposal, setProposal] = useState<ReviewProposal | null>(null)
  const [snapshotId, setSnapshotId] = useState<string | null>(null)
  const [status, setStatus] = useState<string | null>(null)
  const [timeline, setTimeline] = useState<TaskTimelineEntry[]>([])
  const [panelWidth, setPanelWidth] = useState(loadPanelWidth)
  const [resizing, setResizing] = useState(false)
  const sessionRef = useRef<AgentPanelSession | null>(null)
  if (!sessionRef.current) sessionRef.current = new AgentPanelSession(projectId)
  const loopRef = useRef<AgentLoop | null>(null)
  const activeRunRef = useRef<AgentRunScope | null>(null)
  const chatIdsRef = useRef<
    (AgentProjectScope & { storeProjectId: string; chatId: string }) | null
  >(null)
  const e2eProposalLoaded = useRef<string | null>(null)

  useEffect(() => {
    const session = sessionRef.current!
    const projectScope = session.switchProject(projectId)
    activeRunRef.current = null
    chatIdsRef.current = null
    setInput('')
    setText('')
    setBusy(false)
    setChat([])
    setProposal(null)
    setSnapshotId(null)
    setStatus(null)
    setTimeline([])
    const loop = new AgentLoop({
      transport: createLatexTransport(),
      skill: createLatexSkill(window.latexApi, () => projectId),
      events: {
        onText: (value) => {
          const scope = activeRunRef.current
          if (scope && session.acceptsRun(scope)) setText(value)
        },
        onToolStart: (call) => {
          const scope = activeRunRef.current
          if (!scope || !session.acceptsRun(scope)) return
          const runId = `${scope.generation}.${scope.run}`
          setTimeline((current) => startTimelineEntry(current, call, runId))
        },
        onToolExecuted: ({ call, execution }) => {
          const scope = activeRunRef.current
          if (!scope || !session.acceptsRun(scope)) return
          const timelineId = session.timelineId(scope, call.id)
          const runId = `${scope.generation}.${scope.run}`
          setTimeline((current) => {
            const started = current.some((entry) => entry.id === timelineId)
              ? current
              : startTimelineEntry(current, call, runId)
            return completeTimelineEntry(started, timelineId, execution)
          })
          if (call.name !== 'propose_project_edits' || execution.isError) return
          void loadProposalForReview(execution.output, projectId, (request) =>
            window.latexApi.getProposal(request),
          )
            .then((value) => {
              if (session.acceptsRunResult(scope)) setProposal(value)
            })
            .catch((error: unknown) => {
              if (session.acceptsRunResult(scope))
                setStatus(error instanceof Error ? error.message : String(error))
            })
        },
        onDone: (result) => {
          const scope = activeRunRef.current
          if (!scope || !session.acceptsCompletion(scope)) return
          const acceptResult = session.acceptsRun(scope)
          if (result.cancelled) setTimeline(cancelRunningTimelineEntries)
          setBusy(false)
          setText('')
          if (acceptResult && result.text)
            setChat((current) => [...current, { role: 'assistant', text: result.text }])
          const ids = chatIdsRef.current
          if (acceptResult && ids && session.acceptsProject(ids))
            void window.latexApi.appendDirectoryChat({
              projectId,
              storeProjectId: ids.storeProjectId,
              chatId: ids.chatId,
              role: 'assistant',
              text: result.text,
            })
          session.finishRun(scope)
          activeRunRef.current = null
        },
        onError: (error) => {
          const scope = activeRunRef.current
          if (!scope || !session.acceptsCompletion(scope)) return
          if (session.acceptsRun(scope)) {
            setTimeline((current) => failRunningTimelineEntries(current, error))
            setStatus(error)
          }
          setBusy(false)
          setText('')
          session.finishRun(scope)
          activeRunRef.current = null
        },
      },
    })
    loopRef.current = loop
    void window.latexApi.resolveDirectoryChat({ projectId }).then(async (result) => {
      if (!result.ok || !session.acceptsProject(projectScope)) return
      chatIdsRef.current = {
        ...projectScope,
        storeProjectId: result.value.projectId,
        chatId: result.value.chatId,
      }
      const loaded = await window.latexApi.loadDirectoryChat({
        projectId,
        storeProjectId: result.value.projectId,
        chatId: result.value.chatId,
        limit: 200,
      })
      if (loaded.ok && session.acceptsProject(projectScope) && loopRef.current === loop) {
        const restored = loaded.value.map((message) => ({
          role: message.role,
          text: message.text,
        }))
        setChat(restored)
        loopRef.current?.restore(restored)
      }
    })
    return () => {
      session.switchProject(projectId)
      loop.cancel()
      if (loopRef.current === loop) loopRef.current = null
      activeRunRef.current = null
    }
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
    const session = sessionRef.current!
    const projectScope = session.captureProject()
    void window.latexApi
      .proposeProjectEdits({
        projectId,
        files: [{ path: 'main.tex', afterText: E2E_PROPOSAL_TEXT }],
      })
      .then((result) => {
        if (!session.acceptsProject(projectScope)) return
        if (result.ok) setProposal(result.value)
        else setStatus(result.error.message)
      })
  }, [projectId])

  const send = () => {
    const instruction = input.trim()
    const loop = loopRef.current
    if (!instruction || busy || disabled || !loop || loop.busy) return
    const session = sessionRef.current!
    const runScope = session.beginRun()
    activeRunRef.current = runScope
    setBusy(true)
    setStatus(null)
    setText('')
    setInput('')
    setChat((current) => [...current, { role: 'user', text: instruction }])
    const ids = chatIdsRef.current
    if (ids && session.acceptsProject(ids))
      void window.latexApi.appendDirectoryChat({
        projectId,
        storeProjectId: ids.storeProjectId,
        chatId: ids.chatId,
        role: 'user',
        text: instruction,
      })
    loop.run(serializeAgentPrompt(instruction, context))
  }

  const cancel = () => {
    const scope = activeRunRef.current
    if (!scope) return
    sessionRef.current!.cancelRun(scope)
    loopRef.current?.cancel()
    setText('')
    setStatus('Stopped.')
    setTimeline(cancelRunningTimelineEntries)
  }

  const confirm = async (selected: ReadonlySet<string>) => {
    if (!proposal || disabled) return
    const session = sessionRef.current!
    const projectScope = session.captureProject()
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
      if (!session.acceptsProject(projectScope)) return
      await onProjectFilesChanged?.()
      if (!session.acceptsProject(projectScope)) return
      setSnapshotId(value.snapshotId)
      setProposal(null)
      setStatus(
        value.compile.ok
          ? `Changes applied and compiled (${value.compile.result?.diagnostics.length ?? 0} diagnostics).`
          : `Changes applied; compile failed: ${value.compile.error ?? 'unknown error'}`,
      )
    } catch (error) {
      if (session.acceptsProject(projectScope))
        setStatus(error instanceof Error ? error.message : String(error))
    } finally {
      if (session.acceptsProject(projectScope)) setBusy(false)
    }
  }

  const undo = async () => {
    if (!snapshotId || disabled) return
    const session = sessionRef.current!
    const projectScope = session.captureProject()
    setBusy(true)
    const result = await window.latexApi.undoProposal({ projectId, snapshotId })
    if (!session.acceptsProject(projectScope)) return
    if (result.ok) {
      await onProjectFilesChanged?.()
      if (!session.acceptsProject(projectScope)) return
      setSnapshotId(null)
      setStatus('AI changes were undone and the project was compiled again.')
    } else setStatus(result.error.message)
    if (session.acceptsProject(projectScope)) setBusy(false)
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
          <TaskTimeline entries={timeline} />
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
          {contextChips(context).length > 0 && (
            <div className="ai-context-chips" aria-label="Attached context">
              {contextChips(context).map((chip) => (
                <span className="ai-context-chip" key={chip.key}>
                  <span>{chip.label}</span>
                  <button
                    type="button"
                    aria-label={`Remove ${chip.key} context`}
                    onClick={() => onRemoveContext?.(chip.key)}
                  >
                    ×
                  </button>
                </span>
              ))}
            </div>
          )}
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
