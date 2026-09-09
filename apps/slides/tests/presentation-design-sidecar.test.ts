import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  flushPresentationDesignSidecar,
  readPresentationDesignSidecar,
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

  it('reads external edits and keeps the edited document for an untitled deck', () => {
    const senderId = 91358
    const deckPath = join(mkdtempSync(join(tmpdir(), 'wiswork-design-')), 'deck.pptx')
    const designPath = deckPath.replace(/\.pptx$/, '.design.md')
    writeFileSync(designPath, '# DESIGN.md\n\nAccent: #10B981')

    expect(readPresentationDesignSidecar(senderId, deckPath)).toContain('Accent: #10B981')
    expect(savePresentationDesignSidecar(senderId, undefined, 'Accent: #3B82F6')).toBe(true)
    expect(readPresentationDesignSidecar(senderId, undefined)).toContain('Accent: #3B82F6')
  })

  it('persists a complete rendered contract without nesting a second DESIGN.md heading', () => {
    const senderId = 91359
    const deckPath = join(mkdtempSync(join(tmpdir(), 'wiswork-design-')), 'deck.pptx')
    const designMd = '# DESIGN.md\n\nStatus: producing\nRevision: 4\n\n## Brief\n\n- Topic: Nature'

    expect(savePresentationDesignSidecar(senderId, deckPath, designMd)).toBe(true)
    expect(readFileSync(deckPath.replace(/\.pptx$/, '.design.md'), 'utf8')).toBe(designMd)
  })

  it('supports uppercase PowerPoint file extensions', () => {
    const senderId = 91360
    const deckPath = join(mkdtempSync(join(tmpdir(), 'wiswork-design-')), 'deck.PPTX')

    expect(savePresentationDesignSidecar(senderId, deckPath, 'Accent: #3B82F6')).toBe(true)
    expect(readPresentationDesignSidecar(senderId, deckPath)).toContain('Accent: #3B82F6')
  })
})
