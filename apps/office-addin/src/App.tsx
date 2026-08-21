import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
import { createOfficeHostRuntime, type OfficeHostRuntime } from './agent/host-runtime.js'
import type { OfficeProposal, StructuredProposal } from './agent/proposal-controller.js'
import { createPcBridgeAgentTransport } from './agent/transport.js'
import {
  createOfficeAgentSession,
  useOfficeAgent,
  type OfficeAgentSession,
} from './agent/use-office-agent.js'
import { createPcBridgeSession } from './pc-bridge/session.js'
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

export function AgentWorkspace(props: {
  session: OfficeAgentSession
  runtime: OfficeHostRuntime
  disconnect: () => void
  host: OfficeHost
}) {
  const { session, runtime, disconnect, host } = props
  const state = useOfficeAgent(session)
  const [instruction, setInstruction] = useState('')
  const [files, setFiles] = useState<string[]>(runtime.vfs.list('/home/user'))
  const [uploadError, setUploadError] = useState('')
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
  const presentation = proposal ? proposalPresentation(proposal) : undefined

  return (
    <main className="taskpane">
      <header className="app-header">
        <div>
          <span className="eyebrow">WisWork Agent</span>
          <h1>Work with your selection</h1>
          <p>{hostLabels[host]} is connected. Edits always require your approval.</p>
        </div>
        <button
          type="button"
          className="quiet"
          disabled={state.applying}
          onClick={() => {
            disconnect()
          }}
        >
          Log out
        </button>
      </header>

      <section className="status-card" aria-live="polite">
        <span className={`status-dot ${state.busy || state.applying ? 'busy' : ''}`} />
        <div>
          <strong>
            {state.applying
              ? 'Applying approved change'
              : state.busy
                ? 'Agent is working'
                : 'Agent is ready'}
          </strong>
          <p>
            {state.activity ||
              (state.status === 'cancelled' ? 'Run stopped' : 'Ask about the current selection.')}
          </p>
        </div>
      </section>

      {(state.assistantText || state.error) && (
        <section className="message-card" aria-live="polite">
          {state.assistantText && <p className="assistant-text">{state.assistantText}</p>}
          {state.error && <p className="error-text">{state.error}</p>}
        </section>
      )}

      {proposal && (
        <section className="proposal-card" aria-label="Proposed document change">
          <div className="proposal-heading">
            <span className="eyebrow">Approval required</span>
            <h2>{presentation?.title}</h2>
          </div>
          {!isLegacyProposal(proposal) && (
            <div className="preview-block">
              <strong>Impact</strong>
              <p>
                {presentation?.host}: {presentation?.count} item(s)
              </p>
              <pre>{presentation?.targets.join('\n') || '(no named targets)'}</pre>
            </div>
          )}
          <div className="preview-block">
            <strong>Before</strong>
            <pre>{presentation?.before || '(not available)'}</pre>
          </div>
          <div className="preview-block after">
            <strong>After</strong>
            <pre>{presentation?.after || '(described by preview)'}</pre>
          </div>
          {!isLegacyProposal(proposal) && (
            <div className="preview-block">
              <strong>Preview</strong>
              <pre>{presentation?.preview}</pre>
            </div>
          )}
          {!isLegacyProposal(proposal) && proposal.code && (
            <div className="preview-block code-preview">
              <strong>Code</strong>
              <pre>{presentation?.code}</pre>
            </div>
          )}
          <div className="actions">
            <button
              type="button"
              className="secondary"
              disabled={state.applying}
              onClick={() => session.reject()}
            >
              Reject
            </button>
            <button
              type="button"
              disabled={state.busy || state.applying}
              onClick={() => void session.confirm(proposal.id)}
            >
              {state.applying ? 'Applying…' : 'Confirm change'}
            </button>
          </div>
        </section>
      )}

      <section className="composer-card" aria-label="Session files">
        <label htmlFor="session-upload">Session files</label>
        <input
          id="session-upload"
          type="file"
          onChange={(event) => {
            const file = event.currentTarget.files?.[0]
            if (!file) return
            setUploadError('')
            const operation =
              file.name === 'SKILL.md'
                ? runtime.installSkill(file.text())
                : runtime.uploadFile(file.name, file.arrayBuffer())
            void operation
              .then(() => {
                if (mounted.current) setFiles(runtime.vfs.list('/home/user'))
              })
              .catch((error: unknown) => {
                if (!mounted.current) return
                setUploadError(safeUploadError(error))
              })
          }}
        />
        <p>Upload a file, or upload a file named SKILL.md to install its bounded skill package.</p>
        {uploadError && <p className="error-text">{uploadError}</p>}
        <p>{files.length ? files.join(', ') : 'No uploaded files in this session.'}</p>
        <p>
          Installed skills:{' '}
          {runtime.skills
            .list()
            .map((skill) => skill.name)
            .join(', ') || 'none'}
        </p>
      </section>

      <section className="composer-card">
        <label htmlFor="instruction">What should the Agent do?</label>
        <textarea
          id="instruction"
          value={instruction}
          onChange={(event) => setInstruction(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter' && !event.shiftKey) {
              event.preventDefault()
              send()
            }
          }}
          placeholder="For example: make the selected paragraph clearer"
          rows={4}
          maxLength={12_000}
          disabled={state.busy || state.applying}
        />
        <div className="actions">
          {state.busy ? (
            <button
              type="button"
              className="secondary"
              disabled={state.applying}
              onClick={() => session.stop()}
            >
              Stop
            </button>
          ) : (
            <button type="button" disabled={!instruction.trim() || state.applying} onClick={send}>
              Send
            </button>
          )}
        </div>
      </section>
    </main>
  )
}

function ConfiguredApp() {
  const document = useMemo(() => createOfficeDocumentClient(createBrowserOfficeRuntime()), [])
  const bridge = useMemo(() => createPcBridgeSession(), [])
  const bridgeState = useSyncExternalStore(bridge.subscribe, bridge.snapshot, bridge.snapshot)
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
      offline: 'Open WisWork PC, sign in, then retry.',
      signed_out: 'Sign in to WisWork PC first.',
      pending: bridgeState.verificationCode
        ? `Confirm code ${bridgeState.verificationCode} in WisWork PC, then approve.`
        : 'Approve this connection in WisWork PC.',
      rejected: 'The connection was rejected in WisWork PC.',
      expired: 'The connection request expired. Try again.',
    }[bridgeState.status]
    return (
      <StatusScreen
        title="Connect to WisWork PC"
        detail={detail}
        busy={bridgeState.status === 'pending'}
      >
        <button
          type="button"
          disabled={bridgeState.status === 'pending'}
          onClick={() => {
            void bridge.connect(host)
          }}
        >
          {bridgeState.status === 'offline' ? 'Connect to WisWork PC' : 'Try again'}
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
