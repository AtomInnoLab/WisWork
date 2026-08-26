import { describe, expect, it } from 'vitest'
import {
  MAX_PRESENTATION_EVENTS,
  MAX_PRESENTATION_TEXT,
  appendPresentationEvent,
  boundedText,
  emptyPresentationTimeline,
} from '../src/agent/presentation-state.js'

describe('Office presentation state bounds', () => {
  it('retains only the newest bounded events without mutating prior snapshots', () => {
    let timeline = emptyPresentationTimeline()
    const first = timeline
    for (let index = 0; index < MAX_PRESENTATION_EVENTS + 5; index += 1) {
      timeline = appendPresentationEvent(timeline, {
        id: `event-${index}`,
        kind: 'system',
        text: `event ${index}`,
      })
    }
    expect(first).toEqual([])
    expect(timeline).toHaveLength(MAX_PRESENTATION_EVENTS)
    expect(timeline[0]?.id).toBe('event-5')
    expect(Object.isFrozen(timeline)).toBe(true)
  })

  it('bounds presentation text before it reaches the UI', () => {
    const text = boundedText('x'.repeat(MAX_PRESENTATION_TEXT + 100))
    expect(text).toHaveLength(MAX_PRESENTATION_TEXT)
    expect(text.endsWith('…')).toBe(true)
  })
})
