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

  it('keeps the DESIGN.md dialog inside the visible dock instead of the scroll content', () => {
    const styles = read('../src/renderer/styles.css')
    expect(styles).toMatch(/\.ai-panel\s*{[^}]*min-height:\s*0[^}]*overflow:\s*hidden/s)
    expect(styles).toMatch(/\.ai-design-backdrop\s*{[^}]*position:\s*absolute[^}]*inset:\s*0/s)
    expect(styles).toMatch(/\.ai-design-backdrop\s*{[^}]*z-index:\s*40/s)
    expect(styles).toMatch(/\.ai-design-dialog\s*{[^}]*max-height:\s*calc\(100% - 36px\)/s)
  })

  it('offers the native WisWork Markdown editor with an in-panel fallback', () => {
    const panel = read('../src/renderer/ai/AiPanel.tsx')
    const preload = read('../src/preload/index.ts')
    expect(panel).toContain('openDesignSidecar')
    expect(panel).toContain("'在 WisWork 中打开'")
    expect(preload).toContain("ipcRenderer.invoke('ai:open-design-sidecar')")
    expect(panel).toContain('setDesignEditorEditing(true)')
  })
})
