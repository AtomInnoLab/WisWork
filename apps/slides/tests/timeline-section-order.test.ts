import { describe, expect, it } from 'vitest'
import source from '../src/renderer/ai/AiPanel.tsx?raw'

describe('PC agent timeline section order', () => {
  it('renders quality receipts after the current conversation turn', () => {
    const chatStart = source.indexOf('{chat.map((entry, i) => {')
    const qualityStart = source.indexOf('{qualityTimeline.map((receipt) => {')

    expect(chatStart).toBeGreaterThan(-1)
    expect(qualityStart).toBeGreaterThan(chatStart)
  })
})
