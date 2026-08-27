import { readFileSync } from 'node:fs'
import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import { WorkbenchToolbar } from '../src/renderer/workbench/WorkbenchToolbar.js'

const read = (path: string) => readFileSync(new URL(path, import.meta.url), 'utf8')

const toolbarProps = {
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
}

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
    expect(app).toContain('window.latexApi.getCompileDiagnostics({ projectId })')
    expect(app).toContain("next.delete('compile')")
    expect(app).toContain('onToggleFiles={() => setFilesOpen((open) => !open)}')
    expect(app).toContain('onTogglePreview={() => setPreviewOpen((open) => !open)}')
    expect(app).toContain('onToggleAi={() => setAiOpen((open) => !open)}')
    expect(app).toContain('onEditorCommand={(command) =>')
    expect(app).toContain('pdfAvailable={Boolean(editorState.preview?.pdfUrl)}')
    expect(app).toContain('if (editorState.previewStale) setStaleExportOpen(true)')
    expect(app).toContain('<ExportPdfDialog')
    expect(app.match(/mainFile=\{mainFile\}/g)).toHaveLength(2)
    expect(app).toContain("compiling={compiling || bundleStatus.state === 'downloading'}")
    expect(styles).toMatch(/\.latex-workbench-toolbar\s*{[^}]*display:\s*flex/s)
    expect(styles).toMatch(/\.latex-workbench-body\s*{[^}]*min-height:\s*0/s)
    expect(styles).toMatch(/\.latex-toolbar-tabs\s*{[^}]*min-height:\s*39px/s)
    expect(styles).toMatch(/\.latex-toolbar-body\s*{[^}]*height:\s*80px/s)
  })

  it('uses the shared quick-access row and five-tab ribbon contract', () => {
    const html = renderToStaticMarkup(createElement(WorkbenchToolbar, toolbarProps))

    expect(html).toContain('role="toolbar"')
    expect(html).toContain('aria-label="Save"')
    expect(html).toContain('aria-label="Undo"')
    expect(html).toContain('aria-label="Redo"')
    expect(html).toContain('chapters/main.tex')
    expect(html).toContain('aria-label="Unsaved changes"')
    expect(html.match(/role="tab"/g)).toHaveLength(5)
    for (const tab of ['Home', 'Insert', 'Compile', 'PDF', 'View'])
      expect(html).toContain(`>${tab}<`)
    expect(html.indexOf('aria-label="Save"')).toBeLessThan(html.indexOf('>Home<'))
    expect(html.indexOf('aria-label="Undo"')).toBeLessThan(html.indexOf('>Home<'))
    expect(html.indexOf('aria-label="Redo"')).toBeLessThan(html.indexOf('>Home<'))
    expect(html).toContain('M3 4.5C3 3.67158')
    expect(html).toContain('M5.91026 4L2.5 7.14791L5.91026 10.8205')
    expect(html).toContain('M18.0897 4L21.5 7.14791L18.0897 10.8205')
    expect(html).not.toContain('latex-toolbar-quick-glyph')
  })

  it('puts WisWork AI first on Home and separates tab-specific tools', () => {
    const home = renderToStaticMarkup(createElement(WorkbenchToolbar, toolbarProps))
    expect(home.indexOf('>WisWork AI<')).toBeLessThan(home.indexOf('>Bold<'))
    expect(home).toContain('>Section<')
    expect(home).not.toContain('aria-label="Compile"')

    const insert = renderToStaticMarkup(
      createElement(WorkbenchToolbar, { ...toolbarProps, initialTab: 'insert' }),
    )
    for (const command of ['Inline math', 'Equation', 'Figure', 'Table', 'Citation', 'Reference'])
      expect(insert).toContain(`>${command}<`)
    expect(insert).not.toContain('>WisWork AI<')

    const compile = renderToStaticMarkup(
      createElement(WorkbenchToolbar, {
        ...toolbarProps,
        initialTab: 'compile',
        compilePanel: createElement('div', null, 'compile results'),
      }),
    )
    expect(compile).toContain('aria-label="Compile"')
    expect(compile).toContain('>Problems (0)<')
    expect(compile).toContain('data-compile-state="idle"')
    expect(compile).toContain('m9 6 8 6-8 6V6Z')

    const pdf = renderToStaticMarkup(
      createElement(WorkbenchToolbar, {
        ...toolbarProps,
        initialTab: 'pdf',
        pdfAvailable: true,
        onExportPdf: vi.fn(),
      }),
    )
    expect(pdf).toContain('>PDF preview<')
    expect(pdf).toContain('>Export PDF<')

    const view = renderToStaticMarkup(
      createElement(WorkbenchToolbar, { ...toolbarProps, initialTab: 'view' }),
    )
    expect(view).toContain('>Files<')
    expect(view).toContain('>PDF preview<')
    expect(view).toContain('>AI panel<')
  })

  it('turns the sole compile action into cancel while work is busy', () => {
    const html = renderToStaticMarkup(
      createElement(WorkbenchToolbar, {
        activePath: 'main.tex',
        dirty: false,
        disabled: false,
        compiling: true,
        initialTab: 'compile',
        filesOpen: true,
        previewOpen: true,
        aiOpen: true,
        onSave: vi.fn(),
        onCompile: vi.fn(),
        onToggleFiles: vi.fn(),
        onTogglePreview: vi.fn(),
        onToggleAi: vi.fn(),
      }),
    )
    expect(html).toContain('aria-label="Cancel"')
    expect(html).not.toMatch(/<button[^>]+aria-label="Compile"/)
    expect(html).toContain('data-compile-state="running"')
    expect(html).toContain('latex-compile-spinner')
    expect(html).not.toContain('m9 6 8 6-8 6V6Z')
  })

  it('keeps ribbon tools equal height and animates the running compile indicator', () => {
    const styles = read('../src/renderer/styles.css')
    expect(styles).toMatch(
      /\.latex-toolbar-group\s*>\s*\.latex-toolbar-button\s*{[^}]*height:\s*64px/s,
    )
    expect(styles).toMatch(/\.latex-compile-spinner\s*{[^}]*animation:/s)
    expect(styles).toContain('@keyframes latex-compile-spin')
  })
})
