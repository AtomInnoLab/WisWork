import { readFileSync } from 'node:fs'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { AiPanel } from '../src/renderer/ai/AiPanel.js'
import type { AgentContext } from '../src/renderer/ai/agent-context.js'

describe('LaTeX AI dock', () => {
  it('has only WisWork AI and 编译 tabs, with editing kept inside AI', () => {
    const aiHtml = renderToStaticMarkup(
      createElement(AiPanel, {
        projectId: 'project-1',
        activeTab: 'ai',
        compilePanel: createElement('div', null, 'compile evidence'),
      }),
    )
    expect(aiHtml.match(/role="tab"/g)).toHaveLength(2)
    expect(aiHtml).toContain('WisWork AI')
    expect(aiHtml).toContain('编译')
    expect(aiHtml).toContain('Edit LaTeX with WisWork AI')
    expect(aiHtml).not.toContain('compile evidence')

    const compileHtml = renderToStaticMarkup(
      createElement(AiPanel, {
        projectId: 'project-1',
        activeTab: 'compile',
        compilePanel: createElement('div', null, 'compile evidence'),
      }),
    )
    expect(compileHtml).toContain('compile evidence')
    expect(compileHtml).not.toContain('Edit LaTeX with WisWork AI')
    expect(compileHtml).not.toContain('>编辑<')
  })

  it('uses the shared WisWork chat design and product app icon', () => {
    const html = renderToStaticMarkup(
      createElement(AiPanel, { projectId: 'project-1', disabled: false }),
    )
    expect(html).toContain('class="latex-ai-dock"')
    expect(html).toContain('class="ai-panel-header"')
    expect(html).toContain('class="ai-chat"')
    expect(html).toContain('class="ai-input-box"')
    expect(html).toContain('class="ai-brand-icon"')
    expect(html).toContain('Loading project chat…')
  })

  it('stays after the main work area and resizes from its left edge', () => {
    const app = readFileSync(new URL('../src/renderer/App.tsx', import.meta.url), 'utf8')
    const panel = readFileSync(new URL('../src/renderer/ai/AiPanel.tsx', import.meta.url), 'utf8')
    const styles = readFileSync(new URL('../src/renderer/styles.css', import.meta.url), 'utf8')
    expect(app.indexOf('className="latex-main-area"')).toBeLessThan(app.indexOf('{projectId && ('))
    expect(app.match(/<AiPanel/g)).toHaveLength(1)
    expect(app).toContain('open={aiOpen}')
    expect(panel).toContain('window.innerWidth - event.clientX')
    expect(styles).toMatch(/\.ai-panel-resizer\s*{[^}]*left:\s*-?\d+px/s)
  })

  it('shows removable editor and diagnostic context above the composer', () => {
    const context: AgentContext = {
      activeFile: 'main.tex',
      cursorLine: 12,
      selection: { startLine: 10, endLine: 12, text: 'selected', truncated: false },
      diagnostic: {
        path: 'main.tex',
        line: 12,
        column: 1,
        severity: 'error',
        message: 'Undefined control sequence',
      },
    }
    const html = renderToStaticMarkup(
      createElement(AiPanel, {
        projectId: 'project-1',
        disabled: false,
        context,
        onRemoveContext: () => undefined,
      }),
    )
    expect(html).toContain('main.tex:12')
    expect(html).toContain('Selection lines 10–12')
    expect(html).toContain('Error at main.tex:12')
    expect(html.match(/aria-label="Remove [^"]+ context"/g)).toHaveLength(3)
  })

  it('blocks sensitive context chips and explains the boundary', () => {
    const html = renderToStaticMarkup(
      createElement(AiPanel, {
        projectId: 'project-1',
        context: {
          activeFile: 'secrets/secret.tex',
          selection: { startLine: 1, endLine: 1, text: 'TOP_SECRET_VALUE', truncated: false },
        },
        sensitiveContextBlocked: true,
      }),
    )
    expect(html).not.toContain('secret.tex')
    expect(html).not.toContain('TOP_SECRET_VALUE')
    expect(html).toContain('cannot be attached as AI context')
  })
})
