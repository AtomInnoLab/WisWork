import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')

describe('removed synthetic generation progress mechanism', () => {
  it('has no dead progress producer, renderer card, or dedicated styles', () => {
    const skill = read('../src/renderer/ai/slides-skill.ts')
    const panel = read('../src/renderer/ai/AiPanel.tsx')
    const styles = read('../src/renderer/styles.css')

    expect(skill).not.toContain('DeckProgressEvent')
    expect(skill).not.toContain('onProgress')
    expect(panel).not.toContain('DeckProgressCard')
    expect(panel).not.toContain('deckProgress')
    expect(panel).not.toContain('onDeckProgress')
    expect(styles).not.toContain('.deck-progress-')
  })
})
