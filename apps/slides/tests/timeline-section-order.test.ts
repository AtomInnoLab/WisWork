import { describe, expect, it } from 'vitest'
import source from '../src/renderer/ai/AiPanel.tsx?raw'

describe('PC agent timeline section order', () => {
  it('does not run or render a host-owned quality timeline', () => {
    expect(source).not.toContain('publishAppliedDeterministicQuality({')
    expect(source).not.toContain('{qualityTimeline.map((receipt) => {')
  })
})
