import { describe, expect, it } from 'vitest'
import { shouldShowStreamingProgress } from '../src/renderer/ai/streaming-progress'

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
})
