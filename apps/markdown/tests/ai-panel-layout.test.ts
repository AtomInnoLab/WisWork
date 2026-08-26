import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const root = resolve(__dirname, '../../..')
const read = (path: string) => readFileSync(resolve(root, path), 'utf8')

describe('Markdown AI panel layout', () => {
  it('renders the document before the right-side AI dock', () => {
    const app = read('apps/markdown/src/renderer/App.tsx')
    expect(app.indexOf('className="app-content"')).toBeLessThan(app.indexOf('className={`ai-dock'))
  })

  it('resizes from the dock left edge and uses a left divider', () => {
    const panel = read('apps/markdown/src/renderer/ai/AiPanel.tsx')
    const css = read('apps/markdown/src/renderer/styles.css')
    expect(panel).toContain('window.innerWidth - ev.clientX')
    expect(css).toMatch(/\.copilot\s*\{[\s\S]*?border-left:/)
    expect(css).toMatch(/\.ai-panel-resizer\s*\{[\s\S]*?left: 0;/)
  })
})
