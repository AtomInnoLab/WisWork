import { AgentLoop, type AgentSkill, type AgentTransport } from '@wiswork/agent-core'
import { useSyncExternalStore } from 'react'
import type {
  OfficeProposal,
  ProposalController,
  StructuredProposal,
  StructuredProposalController,
} from './proposal-controller.js'
import {
  appendPresentationEvent,
  boundedText,
  emptyPresentationTimeline,
  replacePresentationEvent,
  type OfficePresentationTimeline,
  type ProposalPresentationEvent,
} from './presentation-state.js'

export type AgentSessionStatus = 'idle' | 'working' | 'done' | 'cancelled' | 'error'

export interface OfficeAgentSnapshot {
  assistantText: string
  activity: string
  busy: boolean
  applying: boolean
  status: AgentSessionStatus
  error?: string
  proposal?: OfficeProposal | StructuredProposal
  timeline: OfficePresentationTimeline
}

export interface OfficeAgentSession {
  snapshot(): OfficeAgentSnapshot
  subscribe(listener: () => void): () => void
  send(instruction: string): void
  stop(): void
  confirm(id: string): Promise<void>
  reject(): void
  newTask(): void
  retry(): void
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
    timeline: emptyPresentationTimeline(),
  }
  let cached: OfficeAgentSnapshot = { ...state, proposal: proposals.pending() }

  const publish = (next: Partial<typeof state> = {}) => {
    state = { ...state, ...next }
    cached = { ...state, proposal: proposals.pending() }
    listeners.forEach((listener) => listener())
  }

  let nextEventId = 0
  let activeAssistantId: string | undefined
  let lastInstruction = ''
  const eventId = () => `event-${++nextEventId}`
  const append = (event: Parameters<typeof appendPresentationEvent>[1]) => {
    state = { ...state, timeline: appendPresentationEvent(state.timeline, event) }
  }
  const replace = (id: string, update: Parameters<typeof replacePresentationEvent>[2]) => {
    state = { ...state, timeline: replacePresentationEvent(state.timeline, id, update) }
  }
  const pendingProposalEvent = () =>
    [...state.timeline]
      .reverse()
      .find(
        (event): event is ProposalPresentationEvent =>
          event.kind === 'proposal' && event.state === 'pending',
      )
  const appendPendingProposal = () => {
    const proposal = proposals.pending()
    if (
      proposal &&
      !state.timeline.some(
        (event) => event.kind === 'proposal' && event.proposal.id === proposal.id,
      )
    ) {
      append({ id: eventId(), kind: 'proposal', proposal, state: 'pending' })
    }
  }
  const clearConversation = () => {
    activeAssistantId = undefined
    state = {
      ...state,
      assistantText: '',
      activity: '',
      busy: false,
      applying: false,
      status: 'idle',
      error: undefined,
      timeline: emptyPresentationTimeline(),
    }
  }

  const loop = new AgentLoop({
    transport: dependencies.transport,
    skill: dependencies.skill,
    events: {
      onText: (assistantText) => {
        if (!activeAssistantId) {
          activeAssistantId = eventId()
          append({
            id: activeAssistantId,
            kind: 'assistant',
            text: boundedText(assistantText),
            streaming: true,
          })
        } else {
          replace(activeAssistantId, (event) => ({
            ...event,
            text: boundedText(assistantText),
            streaming: true,
          }))
        }
        publish({ assistantText: boundedText(assistantText) })
      },
      onToolStart: (call) => {
        if (activeAssistantId) {
          replace(activeAssistantId, (event) => ({ ...event, streaming: false }))
          activeAssistantId = undefined
        }
        append({
          id: eventId(),
          kind: 'tool',
          callId: call.id,
          name: boundedText(call.name),
          summary: `Running ${boundedText(call.name)}`,
          state: 'running',
        })
        publish({ activity: `Running ${call.name}` })
      },
      onToolExecuted: (event) => {
        const tool = [...state.timeline]
          .reverse()
          .find((item) => item.kind === 'tool' && item.callId === event.call.id)
        if (tool) {
          replace(tool.id, (item) => {
            if (item.kind !== 'tool') return item
            return {
              ...item,
              summary: boundedText(event.execution.summary),
              state: event.execution.isError ? 'error' : 'complete',
            }
          })
        }
        appendPendingProposal()
        publish({ activity: event.execution.summary })
      },
      onTurnEnd: () => {
        activeAssistantId = undefined
        publish({ activity: 'Thinking…' })
      },
      onDone: (result) => {
        if (activeAssistantId) {
          replace(activeAssistantId, (event) => ({ ...event, streaming: false }))
          activeAssistantId = undefined
        }
        publish({
          busy: false,
          activity: '',
          status: result.cancelled ? 'cancelled' : 'done',
        })
      },
      onError: (error) => {
        const safeError = boundedText(error)
        activeAssistantId = undefined
        append({ id: eventId(), kind: 'error', text: safeError, code: safeError })
        publish({ busy: false, activity: '', status: 'error', error: safeError })
      },
    },
  })

  const startRun = (instruction: string) => {
    const value = instruction.trim()
    if (!value || loop.busy || state.applying) return
    proposals.newTurn()
    lastInstruction = value
    activeAssistantId = undefined
    append({ id: eventId(), kind: 'user', text: boundedText(value) })
    publish({
      assistantText: '',
      activity: 'Thinking…',
      busy: true,
      status: 'working',
      error: undefined,
    })
    loop.run(value)
  }

  return {
    snapshot: () => cached,
    subscribe(listener) {
      listeners.add(listener)
      return () => listeners.delete(listener)
    },
    send(instruction) {
      startRun(instruction)
    },
    stop() {
      if (state.applying) return
      loop.cancel()
    },
    async confirm(id) {
      if (state.applying || loop.busy) return
      const event = pendingProposalEvent()
      if (event?.proposal.id === id) {
        replace(event.id, (item) =>
          item.kind === 'proposal' ? { ...item, state: 'applying', error: undefined } : item,
        )
      }
      publish({ applying: true, error: undefined })
      try {
        await proposals.confirm(id)
        if (event)
          replace(event.id, (item) =>
            item.kind === 'proposal' ? { ...item, state: 'applied' } : item,
          )
        publish({ activity: 'Document updated', error: undefined })
      } catch (error) {
        const safeError = safeConfirmationError(error)
        if (event)
          replace(event.id, (item) =>
            item.kind === 'proposal' ? { ...item, state: 'error', error: safeError } : item,
          )
        publish({ error: safeError, status: 'error' })
      } finally {
        publish({ applying: false })
      }
    },
    reject() {
      if (state.applying) return
      const event = pendingProposalEvent()
      proposals.reject()
      if (event)
        replace(event.id, (item) =>
          item.kind === 'proposal' ? { ...item, state: 'rejected' } : item,
        )
      publish({ activity: 'Proposal rejected', error: undefined })
    },
    newTask() {
      if (state.applying) return
      loop.reset()
      proposals.logout()
      lastInstruction = ''
      clearConversation()
      publish()
    },
    retry() {
      if (!lastInstruction || loop.busy || state.applying) return
      const instruction = lastInstruction
      startRun(instruction)
    },
    logout() {
      if (state.applying) return
      loop.reset()
      proposals.logout()
      lastInstruction = ''
      clearConversation()
      publish()
    },
    authenticationLost() {
      loop.reset()
      proposals.logout()
      lastInstruction = ''
      clearConversation()
      publish()
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
