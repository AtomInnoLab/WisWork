import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')

describe('Slides right AI dock', () => {
  it('renders one persistent AI dock after the normal editing surface', () => {
    const source = read('../src/renderer/App.tsx')
    expect(source.indexOf('className="stage-col"')).toBeLessThan(
      source.indexOf('<div className={`ai-dock${showAi'),
    )
    expect(source.match(/<AiPanel/g)).toHaveLength(1)
  })

  it('resizes from the dock left edge', () => {
    const panel = read('../src/renderer/ai/AiPanel.tsx')
    const styles = read('../src/renderer/styles.css')
    expect(panel).toContain('window.innerWidth - ev.clientX')
    expect(styles).toMatch(/\.ai-panel-resizer\s*{[^}]*left:\s*0/s)
  })
})
