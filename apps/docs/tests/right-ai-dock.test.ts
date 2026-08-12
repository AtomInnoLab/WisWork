import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')

describe('Docs right AI dock', () => {
  it('renders the persistent AI dock after the document workspace content', () => {
    const source = read('../src/renderer/App.tsx')
    expect(source.indexOf('<div className="editor-area">')).toBeLessThan(
      source.indexOf('<div className={`ai-dock${showAi'),
    )
    expect(source.match(/<AiPanel/g)).toHaveLength(1)
  })

  it('resizes from the dock left edge', () => {
    const panel = read('../src/renderer/ai/AiPanel.tsx')
    const styles = read('../src/renderer/styles.css')
    expect(panel).toContain('window.innerWidth - ev.clientX')
    expect(styles).toMatch(/\.ai-panel-resizer\s*{[^}]*left:\s*-?\d+px/s)
  })
})
