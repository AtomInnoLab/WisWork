import { renderToStaticMarkup } from 'react-dom/server'
import { createElement } from 'react'
import { describe, expect, it, vi } from 'vitest'
import { CompilePanel } from '../src/renderer/compile/CompilePanel.js'

describe('CompilePanel bundle status', () => {
  it('shows AI-triggered bundle download progress and keeps cancel available', () => {
    const html = renderToStaticMarkup(
      createElement(CompilePanel, {
        compiling: false,
        bundleStatus: { state: 'downloading', receivedBytes: 25, totalBytes: 100 },
        diagnostics: [],
        log: '',
        onCompile: vi.fn(),
        onCancel: vi.fn(),
        onDiagnostic: vi.fn(),
      }),
    )
    expect(html).toContain('Downloading TeX bundle')
    expect(html).toContain('25%')
    expect(html).toContain('Cancel')
    expect(html).toContain('disabled=""')
  })

  it('describes the indexed TeX bundle as online instead of locally installed', () => {
    const html = renderToStaticMarkup(
      createElement(CompilePanel, {
        compiling: false,
        bundleStatus: { state: 'remote' },
        diagnostics: [],
        log: '',
        onCompile: vi.fn(),
        onCancel: vi.fn(),
        onDiagnostic: vi.fn(),
      }),
    )
    expect(html).toContain('Remote TeX bundle configured')
    expect(html).not.toContain('TeX bundle ready')
  })
})
