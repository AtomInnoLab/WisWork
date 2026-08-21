import { AgentLoop, type AgentSkill, type AgentTransport } from '@wiswork/agent-core'
import { useSyncExternalStore } from 'react'
import type {
  OfficeProposal,
  ProposalController,
  StructuredProposal,
  StructuredProposalController,
} from './proposal-controller.js'

export type AgentSessionStatus = 'idle' | 'working' | 'done' | 'cancelled' | 'error'

export interface OfficeAgentSnapshot {
  assistantText: string
  activity: string
  busy: boolean
  applying: boolean
  status: AgentSessionStatus
  error?: string
  proposal?: OfficeProposal | StructuredProposal
}

export interface OfficeAgentSession {
  snapshot(): OfficeAgentSnapshot
  subscribe(listener: () => void): () => void
  send(instruction: string): void
  stop(): void
  confirm(id: string): Promise<void>
  reject(): void
  logout(): void
  authenticationLost(): void
}

const safeConfirmationError = (error: unknown): string =>
  error instanceof Error && ['proposal_missing', 'proposal_stale'].includes(error.message)
    ? error.message
    : 'office_write_failed'

export function createOfficeAgentSession(dependencies: {
  transport: AgentTransport
  skill: AgentSkill
  proposals: ProposalController | StructuredProposalController
}): OfficeAgentSession {
  const { proposals } = dependencies
  const listeners = new Set<() => void>()
  let state: Omit<OfficeAgentSnapshot, 'proposal'> = {
    assistantText: '',
    activity: '',
    busy: false,
    applying: false,
    status: 'idle',
  }
  let cached: OfficeAgentSnapshot = { ...state, proposal: proposals.pending() }

  const publish = (next: Partial<typeof state> = {}) => {
    state = { ...state, ...next }
    cached = { ...state, proposal: proposals.pending() }
    listeners.forEach((listener) => listener())
  }

  const loop = new AgentLoop({
    transport: dependencies.transport,
    skill: dependencies.skill,
    events: {
      onText: (assistantText) => publish({ assistantText }),
      onToolStart: (call) => publish({ activity: `Running ${call.name}` }),
      onToolExecuted: (event) => publish({ activity: event.execution.summary }),
      onTurnEnd: () => publish({ activity: 'Thinking…' }),
      onDone: (result) =>
        publish({
          busy: false,
          activity: '',
          status: result.cancelled ? 'cancelled' : 'done',
        }),
      onError: (error) => publish({ busy: false, activity: '', status: 'error', error }),
    },
  })

  return {
    snapshot: () => cached,
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    send(instruction) {
      const value = instruction.trim()
      if (!value || loop.busy || state.applying) return
      proposals.newTurn()
      publish({
        assistantText: '',
        activity: 'Thinking…',
        busy: true,
        status: 'working',
        error: undefined,
      })
      loop.run(value)
    },
    stop() {
      if (state.applying) return
      loop.cancel()
    },
    async confirm(id) {
      if (state.applying || loop.busy) return
      publish({ applying: true, error: undefined })
      try {
        await proposals.confirm(id)
        publish({ activity: 'Document updated', error: undefined })
      } catch (error) {
        publish({ error: safeConfirmationError(error), status: 'error' })
      } finally {
        publish({ applying: false })
      }
    },
    reject() {
      if (state.applying) return
      proposals.reject()
      publish({ activity: 'Proposal rejected', error: undefined })
    },
    logout() {
      if (state.applying) return
      loop.reset()
      proposals.logout()
      publish({
        assistantText: '',
        activity: '',
        busy: false,
        applying: false,
        status: 'idle',
        error: undefined,
      })
    },
    authenticationLost() {
      loop.reset()
      proposals.logout()
      publish({
        assistantText: '',
        activity: '',
        busy: false,
        applying: false,
        status: 'idle',
        error: undefined,
      })
    },
  }
}

export function bindAuthLoss(
  auth: { subscribeAuthLoss(listener: () => void): () => void },
  session: OfficeAgentSession,
  signedOut: () => void,
): () => void {
  return auth.subscribeAuthLoss(() => {
    session.authenticationLost()
    signedOut()
  })
}

export function useOfficeAgent(session: OfficeAgentSession): OfficeAgentSnapshot {
  return useSyncExternalStore(session.subscribe, session.snapshot, session.snapshot)
}
