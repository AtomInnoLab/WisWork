import { describe, expect, it } from 'vitest'
import {
  friendlyEnhancedError,
  shouldMarkEnhancedMessageUndelivered,
} from '../src/renderer/ai/enhanced-error-copy'

describe('Enhanced Slides error copy', () => {
  it('never exposes internal runtime error codes to users', () => {
    expect(friendlyEnhancedError('enhanced_turn_failed', '生成失败', '请求超时')).toBe('生成失败')
    expect(friendlyEnhancedError('enhanced_turn_timeout', '生成失败', '请求超时')).toBe('请求超时')
    expect(friendlyEnhancedError('useful provider message', '生成失败', '请求超时')).toBe(
      'useful provider message',
    )
  })

  it('does not claim a failed request had no effect after tools already ran', () => {
    expect(shouldMarkEnhancedMessageUndelivered(0)).toBe(true)
    expect(shouldMarkEnhancedMessageUndelivered(1)).toBe(false)
  })
})
