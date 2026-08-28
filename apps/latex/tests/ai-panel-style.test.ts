import { readFileSync } from 'node:fs'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { AgentActivity, AiPanel } from '../src/renderer/ai/AiPanel.js'
import type { AgentContext } from '../src/renderer/ai/agent-context.js'

describe('LaTeX AI dock', () => {
  it('keeps account identity out of the LaTeX dock header', () => {
    const html = renderToStaticMarkup(
      createElement(AiPanel, {
        projectId: 'project-1',
      }),
    )
    expect(html).not.toContain('latex-account')
    expect(html).not.toContain('writer@example.com')
  })

  it('uses the same icon-only collapsed rail as the other document editors', () => {
    const html = renderToStaticMarkup(
      createElement(AiPanel, { projectId: 'project-1', open: false }),
    )
    expect(html).toContain('class="latex-ai-rail"')
    expect(html).toContain('aria-label="Expand AI panel"')
    expect(html).toContain('viewBox="0 0 130 130.025"')
    expect(html).not.toContain('role="tab"')
    expect(html).not.toContain('编译')
    const styles = readFileSync(new URL('../src/renderer/styles.css', import.meta.url), 'utf8')
    expect(styles).toMatch(/\.latex-ai-rail\s*{[^}]*flex:\s*0 0 38px/s)
  })

  it('keeps compilation out of the AI dock', () => {
    const html = renderToStaticMarkup(
      createElement(AiPanel, {
        projectId: 'project-1',
      }),
    )
    expect(html).toContain('WisWork AI')
    expect(html).toContain('Edit LaTeX with WisWork AI')
    expect(html).not.toContain('编译')
    expect(html).not.toContain('dock-compile-content')
  })

  it('uses the canonical document AI title and collapse icons', () => {
    const html = renderToStaticMarkup(
      createElement(AiPanel, {
        projectId: 'project-1',
        disabled: false,
        onCollapse: () => undefined,
      }),
    )
    expect(html).toContain('class="latex-ai-dock"')
    expect(html).toContain('class="ai-panel-header"')
    expect(html).toContain('class="ai-panel-title"')
    expect(html).toContain('class="ai-chat"')
    expect(html).toContain('class="ai-input-box"')
    expect(html).toContain('viewBox="0 0 130 130.025"')
    expect(html).toContain('M 4.54 8 h 3.39')
    expect(html).not.toContain('class="workspace-tabs"')
    expect(html).not.toContain('>›</button>')
    expect(html).toContain('M13 3.5v4a2.5')
    expect(html).not.toContain('>Send</button>')
    expect(html).toContain('Loading project chat…')
  })

  it('shows agent events with the shared collapsible work-group treatment', () => {
    const running = renderToStaticMarkup(
      createElement(AgentActivity, {
        entries: [
          { id: 'read', kind: 'read', label: 'Read main.tex', state: 'success' },
          { id: 'compile', kind: 'compile', label: 'Compile project', state: 'running' },
        ],
      }),
    )
    expect(running).toContain('class="ai-work-group"')
    expect(running).toContain('class="ai-work-group-summary running"')
    expect(running).toContain('Working')
    expect(running.match(/class="ai-step-row"/g)).toHaveLength(2)
    expect(running).toContain('class="ai-step-icon running"')

    const completed = renderToStaticMarkup(
      createElement(AgentActivity, {
        entries: [
          { id: 'read', kind: 'read', label: 'Read main.tex', state: 'success' },
          { id: 'compile', kind: 'compile', label: 'Compile failed', state: 'error' },
        ],
      }),
    )
    expect(completed).toContain('Worked · 2 steps')
    expect(completed).toContain('class="ai-step-icon error"')
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
