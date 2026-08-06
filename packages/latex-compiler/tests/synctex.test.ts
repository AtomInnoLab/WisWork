import { gzipSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { parseSyncTeX } from '../src/synctex.js'

describe('SyncTeX mapping', () => {
  it('maps isolated input paths back to relative source paths in both directions', () => {
    const data = gzipSync(
      [
        'SyncTeX Version:1',
        'Input:1:/tmp/job/input/main.tex',
        'Input:2:/tmp/job/input/chapters/a.tex',
        'Unit:65781.76',
        '{1',
        'x1,4:10,20',
        'x2,8:30,40',
        '}1',
      ].join('\n'),
    )
    const index = parseSyncTeX(data, '/tmp/job/input')
    expect(index.forward('chapters/a.tex', 8)).toEqual({ page: 1, x: 30, y: 40 })
    expect(index.inverse(1, 11, 19)).toEqual({ path: 'main.tex', line: 4 })
  })

  it('ignores inputs outside the isolated input root', () => {
    const data = gzipSync('Input:1:/etc/passwd\n{1\nx1,1:1,1\n}1')
    expect(parseSyncTeX(data, '/tmp/job/input').forward('passwd', 1)).toBeNull()
  })
  it('parses relative inputs and real box/glue/kern records with coordinate scaling', () => {
    const data = gzipSync(
      [
        'SyncTeX Version:1',
        'Input:1:main.tex',
        'Input:2:chapters/a.tex',
        'Magnification:1000',
        'Unit:1',
        'X Offset:65536',
        'Y Offset:131072',
        'Content:',
        '{1',
        '[1,4:65536,131072:0,0,0',
        'x1,4:65536,131072',
        '(2,8:131072,196608:0,0,0',
        'k2,8:131072,196608:0',
        ')',
        ']',
        '}1',
      ].join('\n'),
    )
    const index = parseSyncTeX(data, '/tmp/job/input')
    const forward = index.forward('main.tex', 4)
    expect(forward?.page).toBe(1)
    expect(forward?.x).toBeCloseTo(131072 / 65781.76, 5)
    expect(forward?.y).toBeCloseTo(262144 / 65781.76, 5)
    expect(index.inverse(1, 3, 5)).toEqual({ path: 'chapters/a.tex', line: 8 })
  })

  it('scales preamble unit by integer magnification but translates preamble offsets independently', () => {
    const data = gzipSync(
      [
        'Input:1:main.tex',
        'Magnification:2000',
        'Unit:65781.76',
        'X Offset:65781.76',
        'Y Offset:131563.52',
        '{1',
        'x1,4:2,3',
        '}1',
      ].join('\n'),
    )
    const point = parseSyncTeX(data, '/tmp/job/input').forward('main.tex', 4)
    expect(point?.x).toBeCloseTo(5, 6)
    expect(point?.y).toBeCloseTo(8, 6)
  })

  it('applies Post Scriptum transforms after point, math, and form records', () => {
    const data = gzipSync(
      [
        'Input:1:main.tex',
        'Unit:65781.76',
        '{1',
        'x1,4:2,2',
        '$1,5:6,2',
        'f7:65782,0',
        'x1,6:2,6',
        '}1',
        'Post scriptum:',
        'Magnification:0.5',
        'X Offset:1bp',
        'Y Offset:1bp',
      ].join('\n'),
    )
    const index = parseSyncTeX(data, '/tmp/job/input')
    expect(index.forward('main.tex', 4)).toEqual({ page: 1, x: 2, y: 2 })
    expect(index.forward('main.tex', 5)?.x).toBeCloseTo(4, 4)
    expect(index.forward('main.tex', 6)?.y).toBeCloseTo(4, 4)
  })

  it('rejects oversized decompressed SyncTeX data', () => {
    const data = gzipSync(`Input:1:main.tex\n${'x'.repeat(1024)}`)
    expect(() => parseSyncTeX(data, '/tmp/job/input', { maxDecompressedBytes: 64 })).toThrow(
      /limit|large|size/i,
    )
  })
  it('limits inputs and records while accepting sparse content within the byte limit', () => {
    const inputs = gzipSync('Input:1:a.tex\nInput:2:b.tex')
    expect(() => parseSyncTeX(inputs, '/tmp/job/input', { maxInputs: 1 })).toThrow(/input limit/i)
    const records = gzipSync('Input:1:main.tex\n{1\nx1,1:0,0\n$1,2:0,0\nf7:0,0\n}1')
    expect(() => parseSyncTeX(records, '/tmp/job/input', { maxRecords: 2 })).toThrow(
      /record limit/i,
    )
    const sparse = gzipSync(`Input:1:main.tex\n${' '.repeat(1024 * 1024)}\n{1\nx1,1:0,0\n}1`)
    expect(
      parseSyncTeX(sparse, '/tmp/job/input', { maxDecompressedBytes: 2 * 1024 * 1024 }).forward(
        'main.tex',
        1,
      ),
    ).not.toBeNull()
  })

  it('accepts optional source columns and expands compressed vertical points', () => {
    const data = gzipSync(
      ['Input:1:main.tex', 'Unit:65781.76', '{1', 'x1,4,2:10,20', 'x1,5:30,=', '}1'].join('\n'),
    )
    const index = parseSyncTeX(data, '/tmp/job/input')
    expect(index.forward('main.tex', 4)).toEqual({ page: 1, x: 10, y: 20 })
    expect(index.forward('main.tex', 5)).toEqual({ page: 1, x: 30, y: 20 })
  })

  it('updates compressed vertical state for full and compressed form references', () => {
    const data = gzipSync(
      [
        'Input:1:main.tex',
        'Unit:65781.76',
        '{1',
        'x1,1:0,10',
        'f7:0,20',
        'x1,2:0,=',
        'f8:0,=',
        '}1',
      ].join('\n'),
    )
    const index = parseSyncTeX(data, '/tmp/job/input')
    expect(index.forward('main.tex', 2)).toEqual({ page: 1, x: 0, y: 20 })
    expect(() => parseSyncTeX(data, '/tmp/job/input', { maxRecords: 3 })).toThrow(/record limit/i)
  })
})
