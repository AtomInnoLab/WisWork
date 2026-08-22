import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
import { createOfficeHostRuntime, type OfficeHostRuntime } from './agent/host-runtime.js'
import { officeCapabilityFlags, officeWorkspaceMode } from '../build-config.js'
import type { OfficeProposal, StructuredProposal } from './agent/proposal-controller.js'
import { MAX_SKILL_BYTES } from './skills/shared/skill-registry.js'
import { MAX_VFS_FILE_BYTES } from './skills/shared/vfs.js'
import { createPcBridgeAgentTransport } from './agent/transport.js'
import {
  createOfficeAgentSession,
  useOfficeAgent,
  type OfficeAgentSession,
} from './agent/use-office-agent.js'
import type {
  OfficePresentationEvent,
  ProposalPresentationEvent,
} from './agent/presentation-state.js'
import { createPcBridgeSession } from './pc-bridge/session.js'
import { createOfficeRelaySession, officeTransportMode } from './relay/session.js'
import {
  createBrowserOfficeRuntime,
  createOfficeDocumentClient,
  type OfficeHost,
} from './office-document.js'

const hostLabels: Record<OfficeHost, string> = {
  word: 'Microsoft Word',
  excel: 'Microsoft Excel',
  powerpoint: 'Microsoft PowerPoint',
  unknown: 'Office',
}

const agentProductLabels: Record<OfficeHost, string> = {
  word: 'AI Word',
  excel: 'AI Sheets',
  powerpoint: 'AI Slides',
  unknown: 'WisWork AI',
}

type DisplayProposal = OfficeProposal | StructuredProposal

const proposalHostLabels: Record<StructuredProposal['impact']['host'], string> = {
  word: 'Word',
  excel: 'Excel',
  powerpoint: 'PowerPoint',
}

function isLegacyProposal(proposal: DisplayProposal): proposal is OfficeProposal {
  return 'value' in proposal
}

function previewText(value: unknown, depth = 0): string {
  if (value === undefined) return ''
  if (typeof value === 'string') return value
  if (typeof value === 'number' || typeof value === 'boolean') return String(value)
  if (value === null) return 'None'
  if (Array.isArray(value)) {
    return value
      .map((item, index) => {
        const rendered = previewText(item, depth + 1)
        return `${index + 1}. ${rendered.replaceAll('\n', ' · ')}`
      })
      .join('\n')
  }
  if (typeof value === 'object') {
    return Object.entries(value as Record<string, unknown>)
      .map(([key, item]) => {
        const label = key.replaceAll('_', ' ').replace(/^./, (character) => character.toUpperCase())
        const rendered = previewText(item, depth + 1)
        return depth > 0 ? `${label}: ${rendered.replaceAll('\n', ' · ')}` : `${label}: ${rendered}`
      })
      .join('\n')
  }
  return String(value)
}

function proposalTarget(target: string): string {
  if (target === 'document:end') return 'End of document'
  if (target === 'document:start') return 'Start of document'
  if (target === 'document') return 'Document'
  if (target === 'selection') return 'Current selection'
  const slide = /^slide[-:](\d+)$/i.exec(target)
  if (slide) return `Slide ${slide[1]}`
  return target
    .replaceAll(':', ' · ')
    .replaceAll('_', ' ')
    .replaceAll('-', ' ')
    .replace(/^./, (character) => character.toUpperCase())
}

export function proposalPresentation(proposal: DisplayProposal) {
  const legacy = isLegacyProposal(proposal)
  const hasComparison = legacy || (proposal.before !== undefined && proposal.after !== undefined)
  const before = previewText(proposal.before)
  const after = previewText(
    legacy
      ? proposal.operation === 'replace'
        ? proposal.value
        : `${proposal.before}${proposal.value}`
      : proposal.after,
  )
  return {
    title: legacy
      ? proposal.operation === 'replace'
        ? 'Replace selection'
        : 'Append to selection'
      : proposal.title,
    host: legacy ? undefined : proposalHostLabels[proposal.impact.host],
    count: legacy ? undefined : proposal.impact.count,
    targets: legacy ? [] : proposal.impact.targets.map(proposalTarget),
    before,
    after,
    preview: hasComparison ? '' : previewText(proposal.preview),
    // Declarative code is an internal safety protocol, not user-facing review content.
    code: undefined,
  }
}

export function safeUploadError(error: unknown): string {
  const code = error instanceof Error ? error.message : ''
  return [
    'upload_cancelled',
    'vfs_limit',
    'vfs_path_denied',
    'invalid_skill_package',
    'skill_already_installed',
    'skill_not_installed',
    'skill_package_limit',
    'skill_package_timeout',
    'office_capability_disabled',
  ].includes(code)
    ? code
    : 'upload_failed'
}

interface SessionFile {
  name: string
  size: number
  arrayBuffer(): Promise<ArrayBuffer>
  text(): Promise<string>
}

export interface OfficeWorkspaceUi {
  readonly attachments: () => readonly string[]
  readonly skills: () => readonly string[]
  readonly skillPackagesEnabled: boolean
  readonly upload: (file: SessionFile) => Promise<void>
  readonly uninstallSkill?: (name: string) => void
  readonly clear: () => void
}

export function createOfficeWorkspaceUi(runtime: OfficeHostRuntime): OfficeWorkspaceUi {
  return Object.freeze({
    attachments: () => Object.freeze([...runtime.vfs.list('/home/user')]),
    skills: () => Object.freeze(runtime.skills.list().map((skill) => skill.name)),
    skillPackagesEnabled: runtime.skillPackagesEnabled,
    upload: (file: SessionFile) => uploadSessionFile(runtime, file),
    uninstallSkill: (name: string) => runtime.uninstallSkill(name),
    clear: () => runtime.clearSession(),
  })
}

export function uploadSessionFile(runtime: OfficeHostRuntime, file: SessionFile): Promise<void> {
  if (file.name.toLowerCase().endsWith('.zip')) {
    return runtime.installSkillPackage(file.arrayBuffer())
  }
  if (file.name === 'SKILL.md') {
    if (file.size > MAX_SKILL_BYTES) return Promise.reject(new Error('invalid_skill_package'))
    return runtime.installSkill(file.text())
  }
  if (file.size > MAX_VFS_FILE_BYTES) return Promise.reject(new Error('vfs_limit'))
  return runtime.uploadFile(file.name, file.arrayBuffer())
}

export type WorkspacePanelName = 'attachments' | 'skills'

export function composerKeyAction(event: {
  key: string
  shiftKey: boolean
  isComposing: boolean
}): 'send' | 'newline' | 'none' {
  if (event.key !== 'Enter') return 'none'
  return event.shiftKey || event.isComposing ? 'newline' : 'send'
}

interface FocusTarget {
  focus(): void
}

export function focusWorkspacePanel(heading: FocusTarget, opener: FocusTarget): () => void {
  heading.focus()
  return () => opener.focus()
}

export function isTimelineNearBottom(viewport: {
  scrollHeight: number
  scrollTop: number
  clientHeight: number
}): boolean {
  return viewport.scrollHeight - viewport.scrollTop - viewport.clientHeight <= 48
}

const starterPrompts: Record<OfficeHost, string[]> = {
  word: ['帮我写一份项目周报', '写一篇产品发布公告', '列一个活动策划提纲'],
  excel: ['分析这份表格的数据', '整理一份项目进度表', '找出数据中的异常'],
  powerpoint: ['起草一份项目汇报', '优化这份演示文稿', '列一个路演演示提纲'],
  unknown: ['帮我起草一份文档', '总结当前内容', '优化这份材料'],
}

function ProposalReview(props: {
  event: ProposalPresentationEvent
  activeProposalId?: string
  busy: boolean
  applying: boolean
  confirm: (id: string) => void
  reject: () => void
}) {
  const { event } = props
  const presentation = proposalPresentation(event.proposal)
  const hasComparison =
    isLegacyProposal(event.proposal) ||
    (event.proposal.before !== undefined && event.proposal.after !== undefined)
  const canReview = event.state === 'pending' && props.activeProposalId === event.proposal.id
  return (
    <article
      className={`proposal-card proposal-${event.state}`}
      aria-label="Proposed document change"
    >
      <div className="proposal-heading">
        <span className="eyebrow">
          {event.state === 'pending'
            ? 'Approval required'
            : event.state === 'applying'
              ? 'Applying approved change'
              : event.state === 'applied'
                ? 'Change applied'
                : event.state === 'rejected'
                  ? 'Change rejected'
                  : 'Change failed'}
        </span>
        <h2>{presentation.title}</h2>
      </div>
      {!isLegacyProposal(event.proposal) && (
        <div className="proposal-impact">
          <span>{presentation.host}</span>
          <span>{presentation.count} item(s)</span>
          <span>{presentation.targets.join(', ') || 'No named targets'}</span>
        </div>
      )}
      <details className="proposal-preview" open>
        <summary>Review exact impact</summary>
        {hasComparison && (
          <div className="proposal-diff">
            <div className="preview-block">
              <strong>Before</strong>
              <p className="proposal-copy">{presentation.before || '(empty document)'}</p>
            </div>
            <div className="preview-block after">
              <strong>After</strong>
              <p className="proposal-copy">{presentation.after || '(empty document)'}</p>
            </div>
          </div>
        )}
        {presentation.preview && <p className="proposal-copy">{presentation.preview}</p>}
      </details>
      {event.error && <p className="error-text">{event.error}</p>}
      {canReview && (
        <div className="actions">
          <button
            type="button"
            className="secondary"
            disabled={props.applying}
            onClick={props.reject}
          >
            Reject
          </button>
          <button
            type="button"
            disabled={props.applying}
            onClick={() => props.confirm(event.proposal.id)}
          >
            {props.applying ? 'Applying…' : 'Confirm change'}
          </button>
        </div>
      )}
      {event.state === 'applying' && <p className="proposal-state">Applying…</p>}
    </article>
  )
}

function TimelineEvent(props: {
  event: OfficePresentationEvent
  activeProposalId?: string
  busy: boolean
  applying: boolean
  confirm: (id: string) => void
  reject: () => void
}) {
  const { event } = props
  if (event.kind === 'proposal') return <ProposalReview {...props} event={event} />
  if (event.kind === 'tool') {
    return (
      <article className={`tool-event tool-${event.state}`} aria-label="Agent activity">
        <span className="tool-indicator" aria-hidden="true" />
        <p>{event.summary}</p>
      </article>
    )
  }
  return (
    <article
      className={`timeline-message message-${event.kind}`}
      {...(event.kind === 'error' ? { role: 'alert' } : {})}
    >
      <span className="message-role">{event.kind === 'user' ? 'You' : 'WisWork'}</span>
      <p>{event.text}</p>
      {event.kind === 'assistant' && event.streaming && (
        <span className="streaming-cursor" aria-label="Response streaming" />
      )}
    </article>
  )
}

export function AgentWorkspace(props: {
  session: OfficeAgentSession
  ui: OfficeWorkspaceUi
  disconnect: () => void
  host: OfficeHost
  initialPanel?: WorkspacePanelName
  legacy?: boolean
}) {
  const { session, ui, disconnect, host } = props
  const state = useOfficeAgent(session)
  const [instruction, setInstruction] = useState('')
  const [files, setFiles] = useState<readonly string[]>(ui.attachments())
  const [skills, setSkills] = useState<readonly string[]>(ui.skills())
  const [uploadError, setUploadError] = useState('')
  const [panel, setPanel] = useState<WorkspacePanelName | undefined>(props.initialPanel)
  const mounted = useRef(true)
  const panelHeading = useRef<HTMLHeadingElement>(null)
  const panelOpener = useRef<HTMLElement | undefined>(undefined)
  const timeline = useRef<HTMLElement>(null)
  const followLatest = useRef(true)
  useEffect(
    () => () => {
      mounted.current = false
    },
    [],
  )
  useEffect(() => {
    const heading = panelHeading.current
    const opener = panelOpener.current
    if (!panel || !heading || !opener) return
    return focusWorkspacePanel(heading, opener)
  }, [panel])
  useEffect(() => {
    const viewport = timeline.current
    if (!viewport || !followLatest.current || typeof viewport.scrollTo !== 'function') return
    viewport.scrollTo({ top: viewport.scrollHeight, behavior: state.busy ? 'auto' : 'smooth' })
  }, [state.busy, state.timeline])

  function send() {
    if (!instruction.trim()) return
    session.send(instruction)
    setInstruction('')
  }

  const proposal = state.proposal
  const hasTimeline = state.timeline.length > 0
  const showConversationChrome =
    hasTimeline || state.busy || state.applying || Boolean(state.error) || Boolean(proposal)
  const showStatus =
    state.busy || state.applying || Boolean(state.activity) || state.status === 'cancelled'

  return (
    <main
      className={`agent-workspace ${props.legacy ? 'legacy-workspace ' : ''}${panel ? 'has-management ' : ''}${showConversationChrome ? 'has-conversation' : 'is-empty'}`}
      aria-busy={state.busy || state.applying}
    >
      {showConversationChrome && (
        <header className="app-header">
          <div className="editor-identity">
            <span className="connection-dot" aria-hidden="true" />
            <h1>{agentProductLabels[host]}</h1>
            <span className="visually-hidden">Connected to WisWork PC</span>
          </div>
          <div className="header-actions">
            <button
              type="button"
              className="quiet"
              disabled={state.applying}
              onClick={() => {
                session.newTask()
                ui.clear()
                setFiles([])
                setSkills([])
                setPanel(undefined)
                setUploadError('')
              }}
            >
              新对话
            </button>
            <details className="session-menu">
              <summary aria-label="Session menu">•••</summary>
              <button
                type="button"
                disabled={state.applying}
                onClick={(event) => {
                  panelOpener.current =
                    event.currentTarget.closest('details')?.querySelector('summary') ??
                    event.currentTarget
                  setPanel('skills')
                  event.currentTarget.closest('details')?.removeAttribute('open')
                }}
              >
                管理技能
              </button>
              <button type="button" disabled={state.applying} onClick={disconnect}>
                退出登录
              </button>
            </details>
          </div>
        </header>
      )}

      {showStatus && (
        <section className="agent-status" aria-live="polite">
          <span className={`status-dot ${state.busy || state.applying ? 'busy' : ''}`} />
          <strong>
            {state.applying
              ? 'Applying approved change'
              : state.busy
                ? 'Agent is working'
                : 'Agent is ready'}
          </strong>
          <span>{state.activity || (state.status === 'cancelled' ? 'Run stopped' : '')}</span>
        </section>
      )}

      <section
        ref={timeline}
        className="agent-timeline"
        aria-label="Agent conversation"
        aria-live="polite"
        onScroll={(event) => {
          followLatest.current = isTimelineNearBottom(event.currentTarget)
        }}
      >
        {!hasTimeline && (
          <div className="empty-state">
            <h2>让 AI 帮你从零起草</h2>
            <p>
              描述主题、要点或粘贴参考素材，
              <br />
              AI 直接为你写出初稿。
            </p>
            <div className="starter-prompts">
              {starterPrompts[host].map((prompt) => (
                <button
                  key={prompt}
                  type="button"
                  className="prompt-chip"
                  onClick={() => setInstruction(prompt)}
                >
                  {prompt}
                </button>
              ))}
            </div>
          </div>
        )}
        {state.timeline.map((event) => (
          <TimelineEvent
            key={event.id}
            event={event}
            activeProposalId={proposal?.id}
            busy={state.busy}
            applying={state.applying}
            confirm={(id) => void session.confirm(id)}
            reject={() => session.reject()}
          />
        ))}
        {proposal &&
          !state.timeline.some(
            (event) => event.kind === 'proposal' && event.proposal.id === proposal.id,
          ) && (
            <ProposalReview
              event={{
                id: `proposal-${proposal.id}`,
                kind: 'proposal',
                proposal,
                state: 'pending',
              }}
              activeProposalId={proposal.id}
              busy={state.busy}
              applying={state.applying}
              confirm={(id) => void session.confirm(id)}
              reject={() => session.reject()}
            />
          )}
        {state.error && state.errorMessage && (
          <div className="error-banner" role="alert">
            <div>
              <p>{state.errorMessage}</p>
            </div>
            {state.retryable && (
              <button
                type="button"
                className="secondary"
                disabled={state.applying}
                onClick={() => session.retry()}
              >
                {state.error === 'proposal_stale' ? '重新生成' : 'Retry'}
              </button>
            )}
          </div>
        )}
      </section>

      {panel && (
        <section
          className="management-panel"
          role="dialog"
          aria-modal="false"
          aria-labelledby="panel-title"
          onKeyDown={(event) => {
            if (event.key === 'Escape') setPanel(undefined)
          }}
        >
          <div className="panel-heading">
            <h2 id="panel-title" ref={panelHeading} tabIndex={-1}>
              {panel === 'attachments' ? 'Session attachments' : 'Installed skills'}
            </h2>
            <button
              type="button"
              className="icon-button"
              aria-label="Close panel"
              onClick={() => setPanel(undefined)}
            >
              ×
            </button>
          </div>
          {panel === 'attachments' ? (
            <>
              <label
                className="upload-button"
                htmlFor="session-upload"
                aria-disabled={state.applying}
              >
                Add attachment
              </label>
              <input
                id="session-upload"
                className="visually-hidden"
                type="file"
                disabled={state.applying}
                onChange={(event) => {
                  const file = event.currentTarget.files?.[0]
                  if (!file) return
                  setUploadError('')
                  void ui
                    .upload(file)
                    .then(() => mounted.current && setFiles(ui.attachments()))
                    .catch(
                      (error: unknown) => mounted.current && setUploadError(safeUploadError(error)),
                    )
                }}
              />
              <p>Files stay in this bounded session and are cleared on logout.</p>
              {uploadError && (
                <p className="error-text" role="alert">
                  {uploadError}
                </p>
              )}
              <ul>
                {files.map((file) => (
                  <li key={file}>{file.split('/').at(-1)}</li>
                ))}
              </ul>
              {!files.length && <p>No session attachments.</p>}
            </>
          ) : (
            <>
              {ui.skillPackagesEnabled && (
                <>
                  <label
                    className="upload-button"
                    htmlFor="skill-package-upload"
                    aria-disabled={state.applying}
                  >
                    Install skill package
                  </label>
                  <input
                    id="skill-package-upload"
                    className="visually-hidden"
                    type="file"
                    accept=".zip,application/zip"
                    disabled={state.applying}
                    onChange={(event) => {
                      const file = event.currentTarget.files?.[0]
                      if (!file) return
                      setUploadError('')
                      void ui
                        .upload(file)
                        .then(() => mounted.current && setSkills(ui.skills()))
                        .catch(
                          (error: unknown) =>
                            mounted.current && setUploadError(safeUploadError(error)),
                        )
                    }}
                  />
                </>
              )}
              <p>Installed instructions can guide the Agent but cannot add Office authority.</p>
              {uploadError && (
                <p className="error-text" role="alert">
                  {uploadError}
                </p>
              )}
              <ul>
                {skills.map((skill) => (
                  <li key={skill}>
                    <span>{skill}</span>
                    <button
                      type="button"
                      className="quiet"
                      disabled={state.applying}
                      onClick={() => {
                        try {
                          ui.uninstallSkill?.(skill)
                          setSkills(ui.skills())
                        } catch (error) {
                          setUploadError(safeUploadError(error))
                        }
                      }}
                    >
                      Remove
                    </button>
                  </li>
                ))}
              </ul>
              {!skills.length && <p>No installed skills.</p>}
            </>
          )}
        </section>
      )}

      <section className="composer-shell" aria-label="Message WisWork Agent">
        <label className="visually-hidden" htmlFor="instruction">
          Message WisWork Agent
        </label>
        <textarea
          id="instruction"
          value={instruction}
          onChange={(event) => setInstruction(event.target.value)}
          onKeyDown={(event) => {
            if (
              composerKeyAction({
                key: event.key,
                shiftKey: event.shiftKey,
                isComposing: event.nativeEvent.isComposing,
              }) === 'send'
            ) {
              event.preventDefault()
              send()
            }
          }}
          placeholder="描述修改、写作要求，或直接提问"
          rows={3}
          maxLength={12_000}
          disabled={state.busy || state.applying}
        />
        <div className="composer-toolbar">
          <div className="composer-tools">
            <button
              type="button"
              className="icon-button"
              aria-label="Attachments"
              aria-expanded={panel === 'attachments'}
              disabled={state.applying}
              onClick={(event) => {
                panelOpener.current = event.currentTarget
                setPanel(panel === 'attachments' ? undefined : 'attachments')
              }}
            >
              📎
            </button>
            <span className="confirmation-chip">
              <span aria-hidden="true" />
              更改需确认
            </span>
          </div>
          {state.busy ? (
            <button
              type="button"
              className="stop-button"
              disabled={state.applying}
              onClick={() => session.stop()}
            >
              Stop
            </button>
          ) : (
            <button
              className="send-button"
              type="button"
              aria-label="Send message"
              disabled={!instruction.trim() || state.applying}
              onClick={send}
            >
              ↑
            </button>
          )}
        </div>
      </section>
    </main>
  )
}

export function LegacyAgentWorkspace(props: {
  session: OfficeAgentSession
  ui: OfficeWorkspaceUi
  disconnect: () => void
  host: OfficeHost
}) {
  return <AgentWorkspace {...props} legacy />
}

export function workspaceComponentForMode(mode: 'workspace' | 'legacy') {
  return mode === 'legacy' ? LegacyAgentWorkspace : AgentWorkspace
}

function ConfiguredApp() {
  const document = useMemo(() => createOfficeDocumentClient(createBrowserOfficeRuntime()), [])
  const bridge = useMemo(
    () =>
      officeTransportMode(import.meta.env) === 'loopback'
        ? createPcBridgeSession()
        : createOfficeRelaySession(),
    [],
  )
  const bridgeState = useSyncExternalStore(
    (listener) => bridge.subscribe(listener),
    () => bridge.snapshot(),
    () => bridge.snapshot(),
  )
  const [workspace, setWorkspace] = useState<
    { runtime: OfficeHostRuntime; session: OfficeAgentSession; ui: OfficeWorkspaceUi } | undefined
  >()
  const workspaceMode = useMemo(() => officeWorkspaceMode(import.meta.env), [])
  const capabilityFlags = useMemo(() => officeCapabilityFlags(import.meta.env), [])
  const [host, setHost] = useState<OfficeHost>('unknown')
  const [hostSupported, setHostSupported] = useState(false)
  const [status, setStatus] = useState('Connecting to Office…')
  const [busy, setBusy] = useState(true)

  useEffect(() => {
    let active = true
    let created:
      { runtime: OfficeHostRuntime; session: OfficeAgentSession; ui: OfficeWorkspaceUi } | undefined
    void (async () => {
      try {
        const activeHost = await document.initialize()
        if (active) {
          setHost(activeHost)
          setHostSupported(activeHost !== 'unknown')
          if (activeHost !== 'unknown') {
            const runtime = createOfficeHostRuntime(activeHost, {
              enableHostSkills: import.meta.env.VITE_WISWORK_OFFICE_HOST_SKILLS !== '0',
              enableConversions: capabilityFlags.conversions,
              enableSkillPackages: capabilityFlags.skillPackages,
              enableImportMedia: capabilityFlags.importMedia,
              document,
            })
            const session = createOfficeAgentSession({
              transport: createPcBridgeAgentTransport(bridge),
              skill: runtime.skill,
              proposals: runtime.proposals,
            })
            created = { runtime, session, ui: createOfficeWorkspaceUi(runtime) }
            setWorkspace(created)
          }
          setStatus(
            activeHost === 'unknown'
              ? 'office_host_unsupported'
              : `${hostLabels[activeHost]} is ready`,
          )
        }
      } catch {
        if (active) setStatus('office_unavailable')
      } finally {
        if (active) setBusy(false)
      }
    })()
    return () => {
      active = false
      created?.session.authenticationLost()
      created?.runtime.dispose()
      bridge.disconnect()
    }
  }, [bridge, capabilityFlags, document])

  useEffect(() => {
    if (bridgeState.status !== 'connected' && workspace) {
      workspace.session.authenticationLost()
      workspace.runtime.clearSession()
    }
  }, [bridgeState.status, workspace])

  if (busy) return <StatusScreen title="Starting WisWork Agent" detail={status} busy />
  if (!hostSupported) {
    return (
      <StatusScreen title="Unsupported Office host" detail="This host cannot use document tools." />
    )
  }
  if (bridgeState.status !== 'connected') {
    const detail = {
      offline: 'Connect again to create a new secure pairing with WisWork PC.',
      connecting: 'Connecting securely to the WisWork Office Relay…',
      signed_out: 'Sign in to WisWork PC first.',
      pending: bridgeState.verificationCode
        ? `Enter code ${bridgeState.verificationCode} in WisWork PC, then approve the matching request.`
        : 'Enter the pairing code in WisWork PC.',
      waiting_for_pc: bridgeState.verificationCode
        ? `Enter code ${bridgeState.verificationCode} in WisWork PC to continue.`
        : 'Waiting for a signed-in WisWork PC.',
      rejected: 'The connection was rejected in WisWork PC.',
      expired: 'The connection request expired. Try again.',
    }[bridgeState.status]
    return (
      <StatusScreen
        title="Connect to WisWork PC"
        detail={detail}
        busy={bridgeState.status === 'connecting' || bridgeState.status === 'pending'}
      >
        <button
          type="button"
          disabled={bridgeState.status === 'connecting' || bridgeState.status === 'pending'}
          onClick={() => {
            void bridge.connect(host)
          }}
        >
          {bridgeState.status === 'offline'
            ? 'Connect to WisWork PC'
            : bridgeState.status === 'connecting'
              ? 'Looking for WisWork PC…'
              : 'Try again'}
        </button>
      </StatusScreen>
    )
  }
  if (!workspace)
    return <StatusScreen title="Starting WisWork Agent" detail="Loading tools…" busy />
  const disconnect = () => {
    workspace.session.logout()
    workspace.runtime.dispose()
    bridge.disconnect()
  }
  const WorkspaceComponent = workspaceComponentForMode(workspaceMode)
  return (
    <WorkspaceComponent
      session={workspace.session}
      ui={workspace.ui}
      disconnect={disconnect}
      host={host}
    />
  )
}

function StatusScreen(props: {
  title: string
  detail: string
  busy?: boolean
  children?: React.ReactNode
}) {
  return (
    <main className="taskpane centered">
      <section className="welcome-card">
        <span className="eyebrow">WisWork Office</span>
        <h1>{props.title}</h1>
        <p aria-live="polite">{props.detail}</p>
        {props.busy && <span className="loading-line" aria-hidden="true" />}
        {props.children}
      </section>
    </main>
  )
}

export function App() {
  return <ConfiguredApp />
}
