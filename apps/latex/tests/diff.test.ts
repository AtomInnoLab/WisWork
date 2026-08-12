import { describe, expect, it } from 'vitest'
import { buildLineDiff } from '../src/renderer/ai/diff.js'

describe('bounded proposal line diff', () => {
  it('emits line-numbered hunks with unchanged context', () => {
    const diff = buildLineDiff(
      'one\ntwo\nthree\nfour\nfive',
      'one\ntwo changed\nthree\nfour\nfive!',
    )

    expect(diff.hunks).toHaveLength(2)
    expect(diff.hunks[0]?.lines).toEqual([
      { kind: 'context', text: 'one', beforeLine: 1, afterLine: 1 },
      { kind: 'remove', text: 'two', beforeLine: 2, afterLine: null },
      { kind: 'add', text: 'two changed', beforeLine: null, afterLine: 2 },
      { kind: 'context', text: 'three', beforeLine: 3, afterLine: 3 },
    ])
    expect(diff.hunks[1]?.lines.at(-1)).toEqual({
      kind: 'add',
      text: 'five!',
      beforeLine: null,
      afterLine: 5,
    })
    expect(diff.summary).toEqual({ added: 2, removed: 2, atLeast: false })
  })

  it('bounds work and rendered output for enormous proposals', () => {
    const before = Array.from(
      { length: 5_000 },
      (_, index) => `before-${index}-${'x'.repeat(100)}`,
    ).join('\n')
    const after = Array.from(
      { length: 5_000 },
      (_, index) => `after-${index}-${'y'.repeat(100)}`,
    ).join('\n')
    const diff = buildLineDiff(before, after, {
      maxInputLines: 80,
      maxOutputLines: 25,
      maxOutputChars: 600,
    })

    expect(diff.truncated).toBe(true)
    expect(diff.hunks.flatMap((hunk) => hunk.lines).length).toBeLessThanOrEqual(25)
    expect(
      diff.hunks.flatMap((hunk) => hunk.lines).reduce((total, line) => total + line.text.length, 0),
    ).toBeLessThanOrEqual(600)
  })

  it('locates the first difference after a long common prefix with bounded retained context', () => {
    const before = Array.from({ length: 2_000 }, (_, index) => `line ${index}`)
    const after = [...before]
    after[1_500] = 'changed late line'
    const diff = buildLineDiff(before.join('\n'), after.join('\n'), { maxInputLines: 20 })

    expect(diff.truncated).toBe(true)
    expect(diff.summary.atLeast).toBe(true)
    expect(diff.hunks.flatMap((hunk) => hunk.lines).map((line) => line.afterLine)).toContain(1_501)
    expect(diff.hunks.flatMap((hunk) => hunk.lines).map((line) => line.text)).toContain(
      'changed late line',
    )
  })

  it('finds a first difference on line 10000 without retaining the common prefix', () => {
    const before = Array.from({ length: 10_001 }, (_, index) => `x${index}`)
    const after = [...before]
    after[9_999] = 'late difference'
    const diff = buildLineDiff(before.join('\n'), after.join('\n'))

    expect(diff.hunks.flatMap((hunk) => hunk.lines).map((line) => line.afterLine)).toContain(10_000)
    expect(diff.hunks.flatMap((hunk) => hunk.lines).length).toBeLessThanOrEqual(240)
  })

  it('reports when the change location is beyond the scan budget', () => {
    const before = `${'same\n'.repeat(100)}before`
    const after = `${'same\n'.repeat(100)}after`
    const diff = buildLineDiff(before, after, { maxScanChars: 100 })

    expect(diff).toMatchObject({
      truncated: true,
      notice: 'change-location-beyond-preview-budget',
      summary: { added: 0, removed: 0, atLeast: true },
    })
    expect(diff.hunks).toEqual([])
  })

  it('never emits a broken Unicode surrogate when enforcing character limits', () => {
    const diff = buildLineDiff('', '😀😀😀', { maxOutputChars: 1 })
    const emitted = diff.hunks
      .flatMap((hunk) => hunk.lines)
      .map((line) => line.text)
      .join('')
    expect(emitted).toBe('😀')
    expect(Array.from(emitted)).toHaveLength(1)
  })

  it('describes a new file without manufacturing removed lines', () => {
    const diff = buildLineDiff(null, 'alpha\nbeta')
    expect(diff.summary).toEqual({ added: 2, removed: 0, atLeast: false })
    expect(diff.hunks[0]?.lines.every((line) => line.kind === 'add')).toBe(true)
  })
})
