import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'
import { FileActionDialog } from '../src/renderer/project/FileActionDialog.js'

describe('LaTeX file action dialog', () => {
  it('uses an in-app path input for creating files', () => {
    const html = renderToStaticMarkup(
      createElement(FileActionDialog, {
        action: { kind: 'create' },
        busy: false,
        onCancel: () => undefined,
        onSubmit: () => undefined,
      }),
    )
    expect(html).toContain('role="dialog"')
    expect(html).toContain('Project-relative path')
    expect(html).toContain('<input')
    expect(html).not.toContain('window.prompt')
  })

  it('requires an explicit confirmation before deletion', () => {
    const html = renderToStaticMarkup(
      createElement(FileActionDialog, {
        action: { kind: 'delete', path: 'chapter.tex' },
        busy: false,
        onCancel: () => undefined,
        onSubmit: () => undefined,
      }),
    )
    expect(html).toContain('Delete <strong>chapter.tex</strong>?')
    expect(html).toContain('cannot be undone')
    expect(html).toContain('class="danger-button"')
  })
})
