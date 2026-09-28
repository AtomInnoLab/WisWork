import { describe, expect, it } from 'vitest'
import { presentationStageTimeline } from '../src/agent/presentation-stage-timeline.js'
import type {
  OfficePresentationEvent,
  ToolPresentationEvent,
} from '../src/agent/presentation-state.js'

const tool = (
  id: string,
  name: string,
  state: ToolPresentationEvent['state'] = 'complete',
): ToolPresentationEvent =>
  Object.freeze({ id, kind: 'tool', callId: `call-${id}`, name, summary: 'Operation', state })

describe('presentation stage timeline projection', () => {
  it('groups adjacent planning operations and retains their identities and progress', () => {
    const events = Object.freeze([
      tool('1', 'read_presentation_plan'),
      tool('2', 'save_presentation_plan', 'running'),
    ])
    const groups = presentationStageTimeline(events)
    expect(groups).toHaveLength(1)
    expect(groups[0]).toMatchObject({
      kind: 'stage',
      id: '1',
      stage: 'planning',
      running: 1,
      ended: 1,
      failed: 0,
      events,
    })
    expect(events).toHaveLength(2)
    expect(events[1]?.state).toBe('running')
  })

  it('keeps different stages and unknown tools separate', () => {
    const unknown = tool('3', 'unknown_tool')
    const groups = presentationStageTimeline([
      tool('1', 'read_presentation_plan'),
      tool('2', 'run_presentation_production'),
      unknown,
      tool('4', 'read_presentation_plan'),
    ])
    expect(groups.map((item) => (item.kind === 'stage' ? item.stage : item.kind))).toEqual([
      'planning',
      'production',
      'tool',
      'planning',
    ])
    expect(groups[2]).toBe(unknown)
  })

  it('preserves messages, errors, and pending confirmation as causal boundaries', () => {
    const boundaries: OfficePresentationEvent[] = [
      { id: 'user', kind: 'user', text: 'Continue' },
      { id: 'assistant', kind: 'assistant', text: 'Review the proposed change' },
      { id: 'error', kind: 'error', text: 'Check the saved result', code: 'network_error' },
      {
        id: 'proposal',
        kind: 'proposal',
        state: 'pending',
        proposal: {
          id: 'p1',
          operation: 'replace',
          before: 'old',
          value: 'new',
          fingerprint: 'fingerprint',
        },
      },
    ]
    const events = boundaries.flatMap((event, index) => [
      tool(`t-${index}`, 'save_presentation_plan'),
      event,
    ])
    const groups = presentationStageTimeline(events)
    expect(groups).toHaveLength(8)
    boundaries.forEach((boundary, index) => expect(groups[index * 2 + 1]).toBe(boundary))
  })

  it('updates counts in place by event identity without mutating the prior projection', () => {
    const first = presentationStageTimeline([
      tool('1', 'run_presentation_production', 'running'),
      tool('2', 'run_presentation_production', 'error'),
    ])
    const next = presentationStageTimeline([
      tool('1', 'run_presentation_production'),
      tool('2', 'run_presentation_production', 'error'),
    ])
    expect(first[0]).toMatchObject({ id: '1', running: 1, ended: 0, failed: 1 })
    expect(next[0]).toMatchObject({ id: '1', running: 0, ended: 1, failed: 1 })
    expect(next).toHaveLength(1)
  })
})
