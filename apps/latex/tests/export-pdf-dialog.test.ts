// @vitest-environment jsdom

import { createElement } from 'react'
import { act } from 'react'
import { createRoot } from 'react-dom/client'
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

  it('moves focus inside, traps keyboard focus, and restores the trigger', () => {
    ;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const trigger = document.createElement('button')
    const container = document.createElement('div')
    document.body.append(trigger, container)
    trigger.focus()
    const root = createRoot(container)
    act(() => {
      root.render(
        createElement(ExportPdfDialog, {
          open: true,
          busy: false,
          onCancel: vi.fn(),
          onCompile: vi.fn(),
          onExportLast: vi.fn(),
        }),
      )
    })
    const buttons = Array.from(container.querySelectorAll('button'))
    expect(document.activeElement).toBe(buttons[0])
    buttons.at(-1)!.focus()
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }))
    expect(document.activeElement).toBe(buttons[0])

    act(() => root.unmount())
    expect(document.activeElement).toBe(trigger)
    trigger.remove()
    container.remove()
  })
})
