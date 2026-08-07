import { describe, expect, it } from 'vitest'
import { clientToPdfPoint, pdfPointToClient } from '../src/renderer/pdf/synctex.js'

describe('SyncTeX UI coordinates', () => {
  it('round trips CSS client coordinates through scaled PDF page points', () => {
    const rect = { left: 100, top: 50, width: 600, height: 800 }
    const point = clientToPdfPoint(rect, 2, 340, 410)
    expect(point).toEqual({ x: 120, y: 180 })
    if (!point) throw new Error('Expected an in-page SyncTeX point')
    expect(pdfPointToClient(rect, 2, point.x, point.y)).toEqual({ clientX: 340, clientY: 410 })
  })

  it('rejects coordinates outside the page or invalid scales', () => {
    const rect = { left: 0, top: 0, width: 100, height: 200 }
    expect(clientToPdfPoint(rect, 1, -1, 20)).toBeNull()
    expect(clientToPdfPoint(rect, 0, 20, 20)).toBeNull()
  })
})
