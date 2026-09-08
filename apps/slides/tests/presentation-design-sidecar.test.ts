import { mkdtempSync, readFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  flushPresentationDesignSidecar,
  savePresentationDesignSidecar,
} from '../src/main/presentation-design-sidecar'

describe('presentation design sidecar', () => {
  it('keeps an untitled design and writes it beside the first saved deck', () => {
    const senderId = 91357
    const deckPath = join(mkdtempSync(join(tmpdir(), 'wiswork-design-')), 'deck.pptx')

    expect(savePresentationDesignSidecar(senderId, undefined, 'Background: #0A0A0A')).toBe(true)
    expect(flushPresentationDesignSidecar(senderId, deckPath)).toBe(true)
    expect(readFileSync(deckPath.replace(/\.pptx$/, '.design.md'), 'utf8')).toContain(
      'Background: #0A0A0A',
    )
  })
})
