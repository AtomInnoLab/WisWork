import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
import { createOfficeHostRuntime, type OfficeHostRuntime } from './agent/host-runtime.js'
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

type DisplayProposal = OfficeProposal | StructuredProposal

function isLegacyProposal(proposal: DisplayProposal): proposal is OfficeProposal {
  return 'value' in proposal
}

function previewText(value: unknown): string {
  if (value === undefined) return ''
  if (typeof value === 'string') return value
  return JSON.stringify(value, null, 2)
}

export function proposalPresentation(proposal: DisplayProposal) {
  const legacy = isLegacyProposal(proposal)
  return {
    title: legacy
      ? proposal.operation === 'replace'
        ? 'Replace selection'
        : 'Append to selection'
      : proposal.title,
    host: legacy ? undefined : proposal.impact.host,
    count: legacy ? undefined : proposal.impact.count,
    targets: legacy ? [] : [...proposal.impact.targets],
    before: previewText(proposal.before),
    after: previewText(
      legacy
        ? proposal.operation === 'replace'
          ? proposal.value
          : `${proposal.before}${proposal.value}`
        : proposal.after,
    ),
    preview: legacy ? '' : previewText(proposal.preview),
    code: legacy ? undefined : proposal.code,
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

export function uploadSessionFile(runtime: OfficeHostRuntime, file: SessionFile): Promise<void> {
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

const starterPrompts: Record<OfficeHost, string[]> = {
  word: ['Summarize this document', 'Make the selected text clearer'],
  excel: ['Explain the selected data', 'Find patterns in this workbook'],
  powerpoint: ['Review this presentation', 'Improve the selected slide'],
  unknown: ['Help with this document'],
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
        <div className="proposal-diff">
          <div className="preview-block">
            <strong>Before</strong>
            <pre>{presentation.before || '(not available)'}</pre>
          </div>
          <div className="preview-block after">
            <strong>After</strong>
            <pre>{presentation.after || '(described by preview)'}</pre>
          </div>
        </div>
        {presentation.preview && <pre>{presentation.preview}</pre>}
        {presentation.code && <pre className="code-preview">{presentation.code}</pre>}
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
            disabled={props.busy || props.applying}
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
      <article
        className={`tool-event tool-${event.state}`}
        aria-label={`${event.name} tool activity`}
      >
        <span className="tool-indicator" aria-hidden="true" />
        <div>
          <strong>{event.name}</strong>
          <p>{event.summary}</p>
        </div>
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
  runtime: OfficeHostRuntime
  disconnect: () => void
  host: OfficeHost
  initialPanel?: WorkspacePanelName
}) {
  const { session, runtime, disconnect, host } = props
  const state = useOfficeAgent(session)
  const [instruction, setInstruction] = useState('')
  const [files, setFiles] = useState<string[]>(runtime.vfs.list('/home/user'))
  const [uploadError, setUploadError] = useState('')
  const [panel, setPanel] = useState<WorkspacePanelName | undefined>(props.initialPanel)
  const mounted = useRef(true)
  useEffect(
    () => () => {
      mounted.current = false
    },
    [],
  )

  function send() {
    if (!instruction.trim()) return
    session.send(instruction)
    setInstruction('')
  }

  const proposal = state.proposal
  const hasTimeline = state.timeline.length > 0

  return (
    <main className="agent-workspace" aria-busy={state.busy || state.applying}>
      <header className="app-header">
        <div>
          <span className="eyebrow">WisWork Agent</span>
          <h1>{hostLabels[host]}</h1>
          <p className="connection-line">
            <span className="connection-dot" />
            PC connected
          </p>
        </div>
        <div className="header-actions">
          <button
            type="button"
            className="quiet"
            disabled={state.applying}
            onClick={() => {
              session.newTask()
              runtime.clearSession()
              setFiles([])
              setPanel(undefined)
              setUploadError('')
            }}
          >
            New task
          </button>
          <details className="session-menu">
            <summary aria-label="Session menu">•••</summary>
            <button type="button" disabled={state.applying} onClick={disconnect}>
              Log out
            </button>
          </details>
        </div>
      </header>

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

      <section className="agent-timeline" aria-label="Agent conversation" aria-live="polite">
        {!hasTimeline && (
          <div className="empty-state">
            <span className="agent-mark" aria-hidden="true">
              W
            </span>
            <h2>What are we working on?</h2>
            <p>
              Ask WisWork to read, explain, or prepare a change. Document edits always need
              approval.
            </p>
            <div className="starter-prompts">
              {starterPrompts[host].map((prompt) => (
                <button
                  key={prompt}
                  type="button"
                  className="prompt-chip"
                  onClick={() => session.send(prompt)}
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
        {state.error && (
          <div className="error-banner" role="alert">
            <p>{state.error}</p>
            <button
              type="button"
              className="secondary"
              disabled={state.applying}
              onClick={() => session.retry()}
            >
              Retry
            </button>
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
            <h2 id="panel-title">
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
              <label className="upload-button" htmlFor="session-upload">
                Add attachment
              </label>
              <input
                id="session-upload"
                className="visually-hidden"
                type="file"
                onChange={(event) => {
                  const file = event.currentTarget.files?.[0]
                  if (!file) return
                  setUploadError('')
                  void uploadSessionFile(runtime, file)
                    .then(() => mounted.current && setFiles(runtime.vfs.list('/home/user')))
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
              <p>Installed instructions can guide the Agent but cannot add Office authority.</p>
              <ul>
                {runtime.skills.list().map((skill) => (
                  <li key={skill.name}>{skill.name}</li>
                ))}
              </ul>
              {!runtime.skills.list().length && <p>No installed skills.</p>}
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
          placeholder={`Ask about ${hostLabels[host]}…`}
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
              onClick={() => setPanel(panel === 'attachments' ? undefined : 'attachments')}
            >
              ＋
            </button>
            <button
              type="button"
              className="tool-button"
              aria-expanded={panel === 'skills'}
              onClick={() => setPanel(panel === 'skills' ? undefined : 'skills')}
            >
              Skills
            </button>
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
    { runtime: OfficeHostRuntime; session: OfficeAgentSession } | undefined
  >()
  const [host, setHost] = useState<OfficeHost>('unknown')
  const [hostSupported, setHostSupported] = useState(false)
  const [status, setStatus] = useState('Connecting to Office…')
  const [busy, setBusy] = useState(true)

  useEffect(() => {
    let active = true
    let created: { runtime: OfficeHostRuntime; session: OfficeAgentSession } | undefined
    void (async () => {
      try {
        const activeHost = await document.initialize()
        if (active) {
          setHost(activeHost)
          setHostSupported(activeHost !== 'unknown')
          if (activeHost !== 'unknown') {
            const runtime = createOfficeHostRuntime(activeHost, {
              enableHostSkills: import.meta.env.VITE_WISWORK_OFFICE_HOST_SKILLS !== '0',
              document,
            })
            const session = createOfficeAgentSession({
              transport: createPcBridgeAgentTransport(bridge),
              skill: runtime.skill,
              proposals: runtime.proposals,
            })
            created = { runtime, session }
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
  }, [bridge, document])

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
  return (
    <AgentWorkspace
      session={workspace.session}
      runtime={workspace.runtime}
      disconnect={() => {
        workspace.session.logout()
        workspace.runtime.dispose()
        bridge.disconnect()
      }}
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
