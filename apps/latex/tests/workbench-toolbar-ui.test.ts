import { readFileSync } from 'node:fs'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import { WorkbenchToolbar } from '../src/renderer/workbench/WorkbenchToolbar.js'

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')

describe('LaTeX workbench toolbar', () => {
  it('puts document commands and workspace visibility controls above the editor', () => {
    const app = read('../src/renderer/App.tsx')
    const styles = read('../src/renderer/styles.css')

    expect(app).toContain('<WorkbenchToolbar')
    expect(app.indexOf('<WorkbenchToolbar')).toBeLessThan(
      app.indexOf('className={`latex-main-area'),
    )
    expect(app).toContain('onSave={() =>')
    expect(app).toContain('onCompile={compileProject}')
    expect(app).toContain('onToggleFiles={() => setFilesOpen((open) => !open)}')
    expect(app).toContain('onTogglePreview={() => setPreviewOpen((open) => !open)}')
    expect(app).toContain('onToggleAi={() => setAiOpen((open) => !open)}')
    expect(app).toContain('onEditorCommand={(command) =>')
    expect(styles).toMatch(/\.latex-workbench-toolbar\s*{[^}]*display:\s*flex/s)
    expect(styles).toMatch(/\.latex-workbench-body\s*{[^}]*min-height:\s*0/s)
    expect(styles).toMatch(/\.latex-toolbar-tabs\s*{[^}]*min-height:\s*39px/s)
    expect(styles).toMatch(/\.latex-toolbar-body\s*{[^}]*height:\s*80px/s)
  })

  it('exposes save, compile, and pressed workspace controls accessibly', () => {
    const html = renderToStaticMarkup(
      createElement(WorkbenchToolbar, {
        activePath: 'chapters/main.tex',
        dirty: true,
        disabled: false,
        compiling: false,
        filesOpen: true,
        previewOpen: false,
        aiOpen: true,
        onSave: vi.fn(),
        onCompile: vi.fn(),
        onToggleFiles: vi.fn(),
        onTogglePreview: vi.fn(),
        onToggleAi: vi.fn(),
      }),
    )

    expect(html).toContain('role="toolbar"')
    expect(html).toContain('aria-label="Save"')
    expect(html).toContain('chapters/main.tex')
    expect(html).toContain('aria-label="Unsaved changes"')
    expect(html.match(/aria-pressed="true"/g)).toHaveLength(2)
    expect(html.match(/aria-pressed="false"/g)).toHaveLength(1)
    for (const command of [
      'Undo',
      'Redo',
      'Bold',
      'Italic',
      'Underline',
      'Section',
      'Subsection',
      'Bulleted list',
      'Numbered list',
      'Inline math',
      'Equation',
      'Figure',
      'Table',
      'Citation',
      'Reference',
    ]) {
      expect(html).toContain(`>${command}<`)
    }
  })
})
