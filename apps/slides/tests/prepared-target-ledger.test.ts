import { describe, expect, it, vi } from 'vitest'
import { PreparedTargetLedger } from '../src/main/operations/prepared-target-ledger'

const response = { status: 'busy' as const }

describe('PreparedTargetLedger', () => {
  it('consumes enrollment on completion but replays bounded metadata', () => {
    const ledger = new PreparedTargetLedger(2, 1_000)
    ledger.set('tx-1', { slideIndex: 0, sourceId: '2' }, response, {
      slideId: 'slide-1',
      sourceId: '2',
      elementId: 'element-1',
    })
    expect(ledger.enrollment('element-1')).toBeDefined()
    ledger.complete('tx-1')
    expect(ledger.enrollment('element-1')).toBeUndefined()
    expect(ledger.get('tx-1', { slideIndex: 0, sourceId: '2' })).toEqual(response)
    expect(ledger.get('tx-1', { slideIndex: 1, sourceId: '2' })).toEqual({
      status: 'conflict',
      code: 'target_stale',
    })
  })

  it('recovers capacity through completed LRU eviction and active TTL expiry', () => {
    vi.useFakeTimers()
    try {
      const ledger = new PreparedTargetLedger(2, 100)
      expect(ledger.set('a', { slideIndex: 0 }, response)).toBe(true)
      ledger.complete('a')
      expect(ledger.set('b', { slideIndex: 0 }, response)).toBe(true)
      expect(ledger.set('c', { slideIndex: 0 }, response)).toBe(true)
      expect(ledger.get('a', { slideIndex: 0 })).toBeUndefined()
      vi.advanceTimersByTime(101)
      expect(ledger.set('d', { slideIndex: 0 }, response)).toBe(true)
    } finally {
      vi.useRealTimers()
    }
  })

  it('never consumes or exposes another preparation token enrollment', () => {
    const ledger = new PreparedTargetLedger(4, 1_000)
    ledger.set('token-a', { slideIndex: 0, sourceId: '2' }, response, {
      slideId: 'slide-1',
      sourceId: '2',
      elementId: 'element-a',
    })
    ledger.set('token-b', { slideIndex: 1, sourceId: '3' }, response, {
      slideId: 'slide-2',
      sourceId: '3',
      elementId: 'element-b',
    })
    ledger.complete('token-a')
    expect(ledger.enrollment('element-a')).toBeUndefined()
    expect(ledger.enrollment('element-b')).toMatchObject({ sourceId: '3' })
    expect(ledger.get('token-b', { slideIndex: 0, sourceId: '2' })).toEqual({
      status: 'conflict',
      code: 'target_stale',
    })
  })
})
