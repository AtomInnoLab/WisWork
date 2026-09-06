import type { ToolDisplay } from '@wiswork/agent-core'
import type { OfficeProposal, StructuredProposal } from './proposal-controller.js'

export const MAX_PRESENTATION_EVENTS = 100
export const MAX_PRESENTATION_TEXT = 12_000

export function presentationProgressLabel(timeline: OfficePresentationTimeline): string {
  let currentTurnStart = 0
  for (let index = timeline.length - 1; index >= 0; index--) {
    if (timeline[index]?.kind !== 'user') continue
    currentTurnStart = index + 1
    break
  }
  return timeline.slice(currentTurnStart).some((event) => event.kind === 'tool')
    ? '继续处理中'
    : '思考中'
}

export interface OfficeClarificationQuestion {
  readonly id: string
  readonly label: string
  readonly description?: string
  readonly options: readonly string[]
}

export type PresentationProposal = OfficeProposal | StructuredProposal

interface PresentationEventBase {
  readonly id: string
}

export interface TextPresentationEvent extends PresentationEventBase {
  readonly kind: 'user' | 'assistant' | 'error' | 'system'
  readonly text: string
  readonly streaming?: boolean
  readonly code?: string
}

export interface ToolPresentationEvent extends PresentationEventBase {
  readonly kind: 'tool'
  readonly callId: string
  readonly name: string
  readonly summary: string
  readonly state: 'running' | 'complete' | 'error'
  readonly durationMs?: number
  readonly output?: string
  readonly display?: ToolDisplay
}

export interface ProposalPresentationEvent extends PresentationEventBase {
  readonly kind: 'proposal'
  readonly proposal: PresentationProposal
  readonly state: 'pending' | 'applying' | 'applied' | 'uncertain' | 'rejected' | 'error'
  readonly error?: string
}

export type OfficePresentationEvent =
  TextPresentationEvent | ToolPresentationEvent | ProposalPresentationEvent

export type OfficePresentationTimeline = readonly OfficePresentationEvent[]

export function boundedText(value: string): string {
  return value.length <= MAX_PRESENTATION_TEXT
    ? value
    : `${value.slice(0, MAX_PRESENTATION_TEXT - 1)}…`
}

export function freezeEvent<T extends OfficePresentationEvent>(event: T): T {
  return Object.freeze(event)
}

export function appendPresentationEvent(
  timeline: OfficePresentationTimeline,
  event: OfficePresentationEvent,
): OfficePresentationTimeline {
  return Object.freeze([...timeline, freezeEvent(event)].slice(-MAX_PRESENTATION_EVENTS))
}

export function replacePresentationEvent(
  timeline: OfficePresentationTimeline,
  id: string,
  replace: (event: OfficePresentationEvent) => OfficePresentationEvent,
): OfficePresentationTimeline {
  let changed = false
  const next = timeline.map((event) => {
    if (event.id !== id) return event
    changed = true
    return freezeEvent(replace(event))
  })
  return changed ? Object.freeze(next) : timeline
}

export const emptyPresentationTimeline = (): OfficePresentationTimeline => Object.freeze([])
