// @vitest-environment jsdom

import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it, vi } from 'vitest'
import { PdfPreview } from '../src/renderer/pdf/PdfPreview.js'

describe('LaTeX PDF preview chrome', () => {
  it('offers a close control without changing the shared viewer', () => {
    const html = renderToStaticMarkup(
      createElement(PdfPreview, {
        pdfUrl: null,
        revision: null,
        location: null,
        stale: false,
        onReverseSync: vi.fn(),
        onClose: vi.fn(),
      }),
    )
    expect(html).toContain('aria-label="Close PDF preview"')
    expect(html).toContain('class="readonly-pdf-viewer"')
  })
})
