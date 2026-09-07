import { describe, expect, it } from 'vitest'
import {
  presentationTimelineBlockOrder,
  shouldShowStreamingProgress,
} from '../src/renderer/ai/streaming-progress'

describe('agent streaming progress', () => {
  it('remains visible while a completion-review continuation runs after existing text', () => {
    expect(
      shouldShowStreamingProgress({
        role: 'assistant',
        text: 'Questionnaire answers received.',
        streaming: true,
      }),
    ).toBe(true)
  })

  it('does not show after the assistant turn settles', () => {
    expect(
      shouldShowStreamingProgress({ role: 'assistant', text: 'Finished.', streaming: false }),
    ).toBe(false)
  })

  it('lets an active tool group own the running state instead of adding a second placeholder', () => {
    expect(
      shouldShowStreamingProgress({
        role: 'assistant',
        text: 'I will inspect the current slides first.',
        streaming: true,
        tools: [{ running: true }],
      }),
    ).toBe(false)
  })

  it('places each tool group after its model-authored message and omits empty message shells', () => {
    expect(
      presentationTimelineBlockOrder({
        role: 'assistant',
        text: 'I will inspect the current slides first.',
        tools: [{ running: false }],
      }),
    ).toEqual(['message', 'tools'])
    expect(
      presentationTimelineBlockOrder({
        role: 'assistant',
        text: '',
        tools: [{ running: true }],
      }),
    ).toEqual(['tools'])
  })
})
