import { describe, expect, it } from 'vitest'
import { createBlankPptx, fingerprintPresentation, openPptx } from '../src/index'

describe('fingerprintPresentation', () => {
  it('is stable for repeated reads and includes package-level concurrent state', async () => {
    const opened = await openPptx(await createBlankPptx())
    const before = await fingerprintPresentation(opened)
    expect(await fingerprintPresentation(opened)).toBe(before)

    opened.archive.entries.set('ppt/comments/comment1.xml', new TextEncoder().encode('<comments/>'))
    expect(await fingerprintPresentation(opened)).not.toBe(before)
  })
})
