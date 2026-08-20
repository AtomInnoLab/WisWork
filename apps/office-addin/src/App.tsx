import { useEffect, useMemo, useState, useSyncExternalStore } from 'react'
import { createOfficeSkill } from './agent/office-skill.js'
import { createProposalController } from './agent/proposal-controller.js'
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

function AgentWorkspace(props: {
  session: OfficeAgentSession
  disconnect: () => void
  host: OfficeHost
}) {
  const { session, disconnect, host } = props
  const state = useOfficeAgent(session)
  const [instruction, setInstruction] = useState('')

  function send() {
    if (!instruction.trim()) return
    session.send(instruction)
    setInstruction('')
  }

  const proposal = state.proposal
  const after = proposal
    ? proposal.operation === 'replace'
      ? proposal.value
      : `${proposal.before}${proposal.value}`
    : ''

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
            <h2>
              {proposal.operation === 'replace' ? 'Replace selection' : 'Append to selection'}
            </h2>
          </div>
          <div className="preview-block">
            <strong>Before</strong>
            <pre>{proposal.before || '(empty selection)'}</pre>
          </div>
          <div className="preview-block after">
            <strong>After</strong>
            <pre>{after}</pre>
          </div>
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
  const proposals = useMemo(() => createProposalController(document), [document])
  const session = useMemo(
    () =>
      createOfficeAgentSession({
        transport: createPcBridgeAgentTransport(bridge),
        skill: createOfficeSkill(document, proposals),
        proposals,
      }),
    [bridge, document, proposals],
  )
  const [host, setHost] = useState<OfficeHost>('unknown')
  const [hostSupported, setHostSupported] = useState(false)
  const [status, setStatus] = useState('Connecting to Office…')
  const [busy, setBusy] = useState(true)

  useEffect(() => {
    let active = true
    void (async () => {
      try {
        const activeHost = await document.initialize()
        if (active) {
          setHost(activeHost)
          setHostSupported(activeHost !== 'unknown')
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
    }
  }, [document])

  useEffect(() => {
    if (bridgeState.status !== 'connected') session.authenticationLost()
  }, [bridgeState.status, session])

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
      pending: 'Approve this connection in WisWork PC.',
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
  return (
    <AgentWorkspace
      session={session}
      disconnect={() => {
        session.logout()
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
