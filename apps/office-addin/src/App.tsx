import { useEffect, useMemo, useState } from 'react'
import { createOfficeSkill } from './agent/office-skill.js'
import { createProposalController } from './agent/proposal-controller.js'
import { createOfficeAgentTransport } from './agent/transport.js'
import {
  createOfficeAgentSession,
  useOfficeAgent,
  type OfficeAgentSession,
} from './agent/use-office-agent.js'
import { createBrowserAuth, type BrowserAuth } from './auth/browser-auth.js'
import { captureAndScrubOAuthCallback } from './auth/oauth-callback.js'
import { loadRuntimeConfig, type RuntimeConfig } from './config.js'
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

const safeAuthError = (error: unknown) =>
  error instanceof Error && ['invalid_callback', 'token_exchange_failed'].includes(error.message)
    ? error.message
    : 'sign_in_failed'

function AgentWorkspace(props: {
  session: OfficeAgentSession
  auth: BrowserAuth
  host: OfficeHost
  onLogout(): void
}) {
  const { session, auth, host, onLogout } = props
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
            session.logout()
            auth.logout()
            onLogout()
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

function ConfiguredApp({ config }: { config: RuntimeConfig }) {
  const document = useMemo(() => createOfficeDocumentClient(createBrowserOfficeRuntime()), [])
  const auth = useMemo(() => createBrowserAuth(config), [config])
  const proposals = useMemo(() => createProposalController(document), [document])
  const session = useMemo(
    () =>
      createOfficeAgentSession({
        transport: createOfficeAgentTransport(config, auth),
        skill: createOfficeSkill(document, proposals),
        proposals,
      }),
    [auth, config, document, proposals],
  )
  const [host, setHost] = useState<OfficeHost>('unknown')
  const [signedIn, setSignedIn] = useState(auth.isAuthenticated())
  const [status, setStatus] = useState('Connecting to Office…')
  const [busy, setBusy] = useState(true)

  useEffect(() => {
    let active = true
    void (async () => {
      try {
        const capturedCallback = captureAndScrubOAuthCallback(
          config.callbackUrl,
          window.location.href,
          (cleanUrl) => window.history.replaceState({}, '', cleanUrl),
        )
        if (capturedCallback) {
          await auth.consumeCallback(capturedCallback)
          if (active) setSignedIn(true)
        }
        const activeHost = await document.initialize()
        if (active) {
          setHost(activeHost)
          setStatus(`${hostLabels[activeHost]} is ready`)
        }
      } catch (error) {
        if (active) setStatus(safeAuthError(error))
      } finally {
        if (active) setBusy(false)
      }
    })()
    return () => {
      active = false
    }
  }, [auth, config.callbackUrl, document])

  if (busy) return <StatusScreen title="Starting WisWork Agent" detail={status} busy />
  if (!signedIn) {
    return (
      <StatusScreen title="WisWork Agent for Office" detail={status}>
        <button
          type="button"
          onClick={() => {
            setBusy(true)
            void auth
              .startAuthorization()
              .then((url) => window.location.assign(url))
              .catch((error) => {
                setStatus(safeAuthError(error))
                setBusy(false)
              })
          }}
        >
          Sign in with WisWork
        </button>
      </StatusScreen>
    )
  }
  return (
    <AgentWorkspace session={session} auth={auth} host={host} onLogout={() => setSignedIn(false)} />
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
  const runtime = useMemo(
    () => loadRuntimeConfig(import.meta.env, { production: import.meta.env.PROD }),
    [],
  )
  if (runtime.status === 'unavailable') {
    return (
      <StatusScreen
        title="Agent unavailable"
        detail="This add-in is not configured for the WisWork Gateway. Contact your administrator."
      />
    )
  }
  return <ConfiguredApp config={runtime.config} />
}
