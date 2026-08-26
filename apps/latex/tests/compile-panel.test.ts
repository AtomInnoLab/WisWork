import { renderToStaticMarkup } from 'react-dom/server'
import { createElement } from 'react'
import { describe, expect, it, vi } from 'vitest'
import { CompilePanel, runCompilePanelAction } from '../src/renderer/compile/CompilePanel.js'

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
        onAskAi: vi.fn(),
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
        onAskAi: vi.fn(),
      }),
    )
    expect(html).toContain('Remote TeX bundle configured')
    expect(html).not.toContain('TeX bundle ready')
  })

  it('disables compile and cancel while the editor transaction is frozen and gates handlers', () => {
    const action = vi.fn()
    runCompilePanelAction(true, action)
    expect(action).not.toHaveBeenCalled()
    const html = renderToStaticMarkup(
      createElement(CompilePanel, {
        compiling: true,
        disabled: true,
        bundleStatus: { state: 'ready', bytes: 1 },
        diagnostics: [],
        log: '',
        onCompile: action,
        onCancel: action,
        onDiagnostic: vi.fn(),
        onAskAi: vi.fn(),
      }),
    )
    expect(html.match(/disabled=""/g)).toHaveLength(2)
  })

  it('offers a separate AI action for each diagnostic', () => {
    const html = renderToStaticMarkup(
      createElement(CompilePanel, {
        compiling: false,
        bundleStatus: { state: 'ready', bytes: 1 },
        diagnostics: [
          {
            path: 'main.tex',
            lineIndex: 4,
            columnIndex: 2,
            severity: 'error',
            message: 'Undefined control sequence',
          },
        ],
        log: '',
        onCompile: vi.fn(),
        onCancel: vi.fn(),
        onDiagnostic: vi.fn(),
        onAskAi: vi.fn(),
      }),
    )
    expect(html).toContain('Open main.tex:5')
    expect(html).toContain('Ask AI about this issue')
  })
})
