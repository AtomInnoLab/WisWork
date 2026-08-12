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
    expect(diff.summary).toEqual({ added: 2, removed: 2 })
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

  it('finds a localized change beyond the initial input budget', () => {
    const before = Array.from({ length: 2_000 }, (_, index) => `line ${index}`)
    const after = [...before]
    after[1_500] = 'changed late line'
    const diff = buildLineDiff(before.join('\n'), after.join('\n'), { maxInputLines: 20 })

    expect(diff.hunks.flatMap((hunk) => hunk.lines).map((line) => line.beforeLine)).toContain(1_501)
    expect(diff.hunks.flatMap((hunk) => hunk.lines).map((line) => line.text)).toContain(
      'changed late line',
    )
  })

  it('describes a new file without manufacturing removed lines', () => {
    const diff = buildLineDiff(null, 'alpha\nbeta')
    expect(diff.summary).toEqual({ added: 2, removed: 0 })
    expect(diff.hunks[0]?.lines.every((line) => line.kind === 'add')).toBe(true)
  })
})
