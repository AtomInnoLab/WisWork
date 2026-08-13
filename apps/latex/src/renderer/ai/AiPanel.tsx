import { useEffect, useRef, useState, type ReactNode } from 'react'
import { AgentLoop } from '@wiswork/agent-core'
import { AiComposer, AiTypingIndicator, Markdown, WisWorkAppMark } from '@wiswork/ui'
import type { LatexAccountStatus } from '../../shared/ipc.js'
import { createLatexSkill } from './latex-skill.js'
import { loadProposalForReview } from './proposal-review.js'
import { ProposalWorkflow, validateUndoProposal } from './proposal-workflow.js'
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
  type ChatLoadState,
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
  sensitiveContextBlocked = false,
  onRemoveContext,
  activeTab = 'ai',
  onTabChange = () => undefined,
  compilePanel = null,
  account = { loggedIn: false },
}: {
  projectId: string
  disabled?: boolean
  onProjectFilesChanged?: () => void | Promise<void>
  open?: boolean
  onExpand?: () => void
  onCollapse?: () => void
  context?: AgentContext
  sensitiveContextBlocked?: boolean
  onRemoveContext?: (key: AgentContextKey) => void
  activeTab?: 'ai' | 'compile'
  onTabChange?: (tab: 'ai' | 'compile') => void
  compilePanel?: ReactNode
  account?: LatexAccountStatus
}) {
  const [input, setInput] = useState('')
  const [text, setText] = useState('')
  const [busy, setBusy] = useState(false)
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [status, setStatus] = useState<string | null>(null)
  const [timeline, setTimeline] = useState<TaskTimelineEntry[]>([])
  const [chatLoadState, setChatLoadState] = useState<ChatLoadState>('loading')
  const [panelWidth, setPanelWidth] = useState(loadPanelWidth)
  const [resizing, setResizing] = useState(false)
  const sessionRef = useRef<AgentPanelSession | null>(null)
  if (!sessionRef.current) sessionRef.current = new AgentPanelSession(projectId)
  const loopRef = useRef<AgentLoop | null>(null)
  const loopBindingRef = useRef<{
    loop: AgentLoop
    getRun: () => AgentRunScope | null
    setRun: (scope: AgentRunScope | null) => void
  } | null>(null)
  const chatLoadStateRef = useRef<ChatLoadState>('loading')
  const chatIdsRef = useRef<
    (AgentProjectScope & { storeProjectId: string; chatId: string }) | null
  >(null)
  const e2eProposalLoaded = useRef<string | null>(null)
  const projectFilesChangedRef = useRef(onProjectFilesChanged)
  projectFilesChangedRef.current = onProjectFilesChanged
  const workflowRef = useRef<ProposalWorkflow | null>(null)
  if (!workflowRef.current) {
    workflowRef.current = new ProposalWorkflow(projectId, {
      verify: (request) => window.latexApi.verifyProposal(request),
      create: (request) => window.latexApi.proposeProjectEdits(request),
      apply: (request) => window.latexApi.applyProposal(request),
      refresh: () => projectFilesChangedRef.current?.(),
    })
  }
  const workflow = workflowRef.current
  const [proposalWorkflow, setProposalWorkflow] = useState(workflow.state)

  useEffect(() => workflow.subscribe(setProposalWorkflow), [workflow])

  useEffect(() => {
    const session = sessionRef.current!
    let loopRun: AgentRunScope | null = null
    chatIdsRef.current = null
    chatLoadStateRef.current = 'loading'
    setChatLoadState('loading')
    setInput('')
    setText('')
    setBusy(false)
    setChat([])
    workflow.setProject(projectId)
    setStatus(null)
    setTimeline([])
    const loop = new AgentLoop({
      transport: createLatexTransport(),
      skill: createLatexSkill(window.latexApi, () => projectId),
      events: {
        onText: (value) => {
          if (loopRef.current !== loop) return
          const scope = loopRun
          if (scope && session.acceptsRun(loop, scope)) setText(value)
        },
        onToolStart: (call) => {
          if (loopRef.current !== loop) return
          const scope = loopRun
          if (!scope || !session.acceptsRun(loop, scope)) return
          const runId = `${scope.generation}.${scope.run}`
          setTimeline((current) => startTimelineEntry(current, call, runId))
        },
        onToolExecuted: ({ call, execution }) => {
          if (loopRef.current !== loop) return
          const scope = loopRun
          if (!scope || !session.acceptsRun(loop, scope)) return
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
              if (loopRef.current === loop && session.acceptsRunResult(loop, scope))
                void workflow.setProposal(value)
            })
            .catch((error: unknown) => {
              if (loopRef.current === loop && session.acceptsRunResult(loop, scope))
                setStatus(error instanceof Error ? error.message : String(error))
            })
        },
        onDone: (result) => {
          if (loopRef.current !== loop) return
          const scope = loopRun
          if (!scope || !session.acceptsCompletion(loop, scope)) return
          const acceptResult = session.acceptsRun(loop, scope)
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
          session.finishRun(loop, scope)
          loopRun = null
        },
        onError: (error) => {
          if (loopRef.current !== loop) return
          const scope = loopRun
          if (!scope || !session.acceptsCompletion(loop, scope)) return
          if (session.acceptsRun(loop, scope)) {
            setTimeline((current) => failRunningTimelineEntries(current, error))
            setStatus(error)
          }
          setBusy(false)
          setText('')
          session.finishRun(loop, scope)
          loopRun = null
        },
      },
    })
    const projectScope = session.attachLoop(loop, projectId)
    loopRef.current = loop
    loopBindingRef.current = {
      loop,
      getRun: () => loopRun,
      setRun: (scope) => {
        loopRun = scope
      },
    }
    const finishChatLoad = (state: Exclude<ChatLoadState, 'loading'>, message?: string) => {
      if (loopRef.current !== loop || !session.acceptsLoopProject(loop, projectScope)) return
      chatLoadStateRef.current = state
      setChatLoadState(state)
      if (message) setStatus(message)
    }
    void (async () => {
      try {
        const result = await window.latexApi.resolveDirectoryChat({ projectId })
        if (loopRef.current !== loop || !session.acceptsLoopProject(loop, projectScope)) return
        if (!result.ok) {
          finishChatLoad('error', `Chat history unavailable: ${result.error.message}`)
          return
        }
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
        if (loopRef.current !== loop || !session.acceptsLoopProject(loop, projectScope)) return
        if (!loaded.ok) {
          chatIdsRef.current = null
          finishChatLoad('error', `Chat history unavailable: ${loaded.error.message}`)
          return
        }
        if (session.canRestoreChat(loop, projectScope)) {
          const restored = loaded.value.map((message) => ({
            role: message.role,
            text: message.text,
          }))
          setChat(restored)
          loop.restore(restored)
        }
        finishChatLoad('ready')
      } catch (error) {
        chatIdsRef.current = null
        finishChatLoad(
          'error',
          `Chat history unavailable: ${error instanceof Error ? error.message : String(error)}`,
        )
      }
    })()
    return () => {
      workflow.cancel()
      session.detachLoop(loop)
      loop.cancel()
      if (loopRef.current === loop) loopRef.current = null
      if (loopBindingRef.current?.loop === loop) loopBindingRef.current = null
      loopRun = null
    }
  }, [projectId, workflow])

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
        if (result.ok) void workflow.setProposal(result.value)
        else setStatus(result.error.message)
      })
  }, [projectId, workflow])

  const send = () => {
    const instruction = input.trim()
    const binding = loopBindingRef.current
    if (
      !instruction ||
      busy ||
      disabled ||
      !binding ||
      binding.loop.busy ||
      !sessionRef.current!.canSend(binding.loop, chatLoadStateRef.current)
    )
      return
    const session = sessionRef.current!
    const runScope = session.beginRun(binding.loop)
    binding.setRun(runScope)
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
    binding.loop.run(serializeAgentPrompt(instruction, context))
  }

  const cancel = () => {
    const binding = loopBindingRef.current
    const scope = binding?.getRun()
    if (!binding || !scope) return
    sessionRef.current!.cancelRun(binding.loop, scope)
    binding.loop.cancel()
    setText('')
    setStatus('Stopped.')
    setTimeline(cancelRunningTimelineEntries)
  }

  const undo = async () => {
    const snapshotId = proposalWorkflow.snapshotId
    if (!snapshotId || disabled) return
    const session = sessionRef.current!
    const projectScope = session.captureProject()
    setBusy(true)
    try {
      const result = await window.latexApi.undoProposal({ projectId, snapshotId })
      if (!session.acceptsProject(projectScope)) return
      if (!result.ok) {
        setStatus(result.error.message)
        return
      }
      const undone = validateUndoProposal(result.value, snapshotId)
      if (!undone.restored) {
        workflow.clearSnapshot('AI changes were already undone; no compile was needed.')
        return
      }
      workflow.clearSnapshot(
        undone.compile.ok
          ? 'AI changes were undone and the project was compiled again.'
          : `AI changes were undone, but compile failed: ${undone.compile.error}`,
      )
      try {
        await projectFilesChangedRef.current?.()
      } catch (error) {
        if (session.acceptsProject(projectScope)) {
          workflow.clearSnapshot(
            `AI changes were undone, but file refresh failed: ${error instanceof Error ? error.message : String(error)}.`,
          )
        }
      }
    } catch (error) {
      if (session.acceptsProject(projectScope)) {
        setStatus(error instanceof Error ? error.message : 'Undo response was invalid.')
      }
    } finally {
      if (session.acceptsProject(projectScope)) setBusy(false)
    }
  }

  // Keep the component mounted while collapsed so its transcript, draft and active run survive.
  if (!open) {
    return (
      <aside className="latex-ai-rail" role="tablist" aria-label="Workspace panel">
        <button
          type="button"
          role="tab"
          aria-selected={activeTab === 'ai'}
          className={activeTab === 'ai' ? 'active' : undefined}
          onClick={() => {
            onTabChange('ai')
            onExpand?.()
          }}
        >
          <WisWorkAppMark className="ai-brand-icon" size={22} />
          <span>WisWork AI</span>
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={activeTab === 'compile'}
          className={activeTab === 'compile' ? 'active' : undefined}
          onClick={() => {
            onTabChange('compile')
            onExpand?.()
          }}
        >
          <span>编译</span>
        </button>
      </aside>
    )
  }

  const accountLabel = account.loggedIn
    ? (account.email ?? account.userId ?? 'Signed in')
    : 'Not signed in'
  const accountInitial = account.loggedIn
    ? (account.email ?? account.userId ?? 'W').trim().charAt(0).toUpperCase() || 'W'
    : '?'

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
          <div className="workspace-tabs" role="tablist" aria-label="Workspace panel">
            <button
              type="button"
              role="tab"
              aria-selected={activeTab === 'ai'}
              className={activeTab === 'ai' ? 'active' : undefined}
              onClick={() => onTabChange('ai')}
            >
              <WisWorkAppMark className="ai-brand-icon" size={25} />
              <span>WisWork AI</span>
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={activeTab === 'compile'}
              className={activeTab === 'compile' ? 'active' : undefined}
              onClick={() => onTabChange('compile')}
            >
              编译
            </button>
          </div>
          <div className="ai-panel-actions">
            <div className="latex-account" title={accountLabel} aria-label={accountLabel}>
              <span className="latex-account-avatar">{accountInitial}</span>
              <span className="latex-account-label">{accountLabel}</span>
            </div>
            {onCollapse && (
              <button className="ai-header-btn" onClick={onCollapse} aria-label="Collapse AI panel">
                ›
              </button>
            )}
          </div>
        </header>
        {activeTab === 'ai' ? (
          <>
            <div className="ai-chat" aria-live="polite" role="tabpanel">
              {chat.length === 0 && !text && !proposalWorkflow.proposal && (
                <div className="ai-chat-empty">
                  <div className="ai-chat-empty-title">Edit LaTeX with WisWork AI</div>
                  <div className="ai-chat-empty-body">
                    Ask questions, fix compilation issues, or propose project-wide changes.
                  </div>
                  <div className="ai-starter-list">
                    {[
                      'Explain this document',
                      'Fix LaTeX errors',
                      'Improve the current section',
                    ].map((starter) => (
                      <button
                        key={starter}
                        className="ai-starter"
                        onClick={() => setInput(starter)}
                      >
                        {starter}
                      </button>
                    ))}
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
              {proposalWorkflow.proposal && proposalWorkflow.verification && (
                <ProposalReview
                  proposal={proposalWorkflow.proposal}
                  selection={proposalWorkflow.selection}
                  busy={busy || proposalWorkflow.busy || disabled}
                  verification={proposalWorkflow.verification}
                  riskArmed={proposalWorkflow.riskArmed}
                  onSelectionChange={(selection) => workflow.setSelection(selection)}
                  onPrimaryAction={() => void workflow.primaryAction()}
                  onCancel={() => workflow.cancel()}
                />
              )}
              {status && (
                <div className="ai-status" role="status">
                  {status}
                </div>
              )}
              {proposalWorkflow.status && (
                <div className="ai-status" role="status">
                  {proposalWorkflow.status}
                </div>
              )}
              {proposalWorkflow.snapshotId && (
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
              {sensitiveContextBlocked && (
                <div className="ai-status" role="status">
                  Sensitive files remain editable, but cannot be attached as AI context.
                </div>
              )}
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
                placeholder={
                  chatLoadState === 'loading'
                    ? 'Loading project chat…'
                    : 'Ask WisWork AI about this LaTeX project'
                }
                hintIdle={
                  chatLoadState === 'error'
                    ? 'Chat history unavailable · messages will not be saved'
                    : 'Enter to send · Shift+Enter for new line'
                }
                hintBusy="Working…"
                hintIdleTitle="Enter to send"
                sendLabel="Send"
                stopLabel="Stop"
                ariaLabel="Ask WisWork AI"
                onChange={(value) => {
                  if (chatLoadState !== 'loading') setInput(value)
                }}
                onSend={send}
                onStop={cancel}
              />
            </div>
          </>
        ) : (
          <div className="dock-compile-content" role="tabpanel">
            {compilePanel}
          </div>
        )}
      </aside>
    </div>
  )
}
