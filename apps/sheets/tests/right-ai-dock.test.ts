import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')

describe('Sheets right AI dock', () => {
  it('renders the persistent AI panel after the workbook', () => {
    const source = read('../src/renderer/ExcelShell.tsx')
    expect(source.indexOf('<div className="sheet-main">')).toBeLessThan(
      source.indexOf('<AiChatPanel'),
    )
    expect(source.match(/<AiChatPanel/g)).toHaveLength(1)
  })

  it('places the resizable panel in the final grid column', () => {
    const panel = read('../src/renderer/ai/AiChatPanel.tsx')
    const styles = read('../src/renderer/styles.css')
    expect(panel).toContain('window.innerWidth - ev.clientX')
    expect(styles).toContain('grid-template-columns: minmax(0, 1fr) var(--copilot-width, 360px)')
    expect(styles).toMatch(/\.ai-panel-resizer\s*{[^}]*left:\s*0/s)
  })
})
