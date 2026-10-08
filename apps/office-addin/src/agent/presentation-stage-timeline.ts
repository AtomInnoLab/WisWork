import { acpPresentationStage, type PresentationStage } from '@wiswork/agent-harness'
import type { OfficePresentationEvent, ToolPresentationEvent } from './presentation-state.js'

export interface PresentationStageGroup {
  readonly kind: 'stage'
  readonly id: string
  readonly stage: PresentationStage
  readonly events: readonly ToolPresentationEvent[]
  readonly running: number
  readonly ended: number
  readonly failed: number
}

export type PresentationTimelineItem = OfficePresentationEvent | PresentationStageGroup

/** Display projection only; tool termination does not prove page or project completion. */
export function presentationStageTimeline(
  timeline: readonly OfficePresentationEvent[],
): readonly PresentationTimelineItem[] {
  const items: (
    | OfficePresentationEvent
    | { kind: 'stage'; stage: PresentationStage; events: ToolPresentationEvent[] }
  )[] = []
  for (const event of timeline) {
    const stage = event.kind === 'tool' ? acpPresentationStage(event.name) : undefined
    if (!stage || event.kind !== 'tool') {
      items.push(event)
      continue
    }
    const previous = items.at(-1)
    if (previous?.kind === 'stage' && previous.stage === stage) previous.events.push(event)
    else items.push({ kind: 'stage', stage, events: [event] })
  }
  return Object.freeze(
    items.map((item) =>
      item.kind !== 'stage'
        ? item
        : Object.freeze({
            ...item,
            id: item.events[0]!.id,
            events: Object.freeze(item.events),
            running: item.events.filter((event) => event.state === 'running').length,
            ended: item.events.filter((event) => event.state === 'complete').length,
            failed: item.events.filter((event) => event.state === 'error').length,
          }),
    ),
  )
}
