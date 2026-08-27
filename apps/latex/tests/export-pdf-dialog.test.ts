import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import { ExportPdfDialog } from '../src/renderer/pdf/ExportPdfDialog.js'

describe('stale PDF export dialog', () => {
  it('offers compile, last successful PDF, and cancel as explicit choices', () => {
    const html = renderToStaticMarkup(
      createElement(ExportPdfDialog, {
        open: true,
        busy: false,
        onCancel: vi.fn(),
        onCompile: vi.fn(),
        onExportLast: vi.fn(),
      }),
    )
    expect(html).toContain('role="dialog"')
    expect(html).toContain('uncompiled changes')
    expect(html).toContain('>Cancel<')
    expect(html).toContain('>Compile now<')
    expect(html).toContain('>Export last PDF<')
  })

  it('does not render while closed', () => {
    const html = renderToStaticMarkup(
      createElement(ExportPdfDialog, {
        open: false,
        busy: false,
        onCancel: vi.fn(),
        onCompile: vi.fn(),
        onExportLast: vi.fn(),
      }),
    )
    expect(html).toBe('')
  })
})
