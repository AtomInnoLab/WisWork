import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { AiPanel } from '../src/renderer/ai/AiPanel.js'

describe('LaTeX AI dock', () => {
  it('uses the shared WisWork chat design and product app icon', () => {
    const html = renderToStaticMarkup(
      createElement(AiPanel, { projectId: 'project-1', disabled: false }),
    )
    expect(html).toContain('class="latex-ai-dock"')
    expect(html).toContain('class="ai-panel-header"')
    expect(html).toContain('class="ai-chat"')
    expect(html).toContain('class="ai-input-box"')
    expect(html).toContain('class="ai-brand-icon"')
  })
})
