// @vitest-environment jsdom
import React, { act } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { OfficeDesignPanel } from '../src/OfficeDesignPanel.js'

let root: Root
let container: HTMLDivElement
beforeEach(() => {
  Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })
  container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
})
afterEach(() => {
  act(() => root.unmount())
  container.remove()
  vi.useRealTimers()
})
const original = '# DESIGN.md\n\n## Visual direction\n\n**Ocean blue**'
const saved = '# DESIGN.md\n\n## Visual direction\n\n**Volcano black**'
const current = { markdown: original, sourceId: 'plan-1' }
const button = (text: string) =>
  Array.from(container.querySelectorAll('button')).find((b) => b.textContent?.includes(text))!

describe('Office DESIGN.md reading and connected-PC editing', () => {
  it('renders Markdown in a prominent reader and synchronizes PC saves as an applicable draft', async () => {
    vi.useFakeTimers()
    const request = vi.fn(async (body: { action: string; documentId: string }) => ({
      documentId: body.documentId,
      markdown: body.action === 'read' ? saved : original,
      revision: body.action === 'read' ? 'new' : 'old',
    }))
    const apply = vi.fn()
    await act(async () =>
      root.render(
        React.createElement(OfficeDesignPanel, { current, busy: false, request, onApply: apply }),
      ),
    )
    expect(container.querySelector('.design-document-entry')).not.toBeNull()
    await act(async () => button('DESIGN.md').click())
    expect(
      Array.from(container.querySelectorAll('[role="dialog"] .ai-md-h')).some(
        (heading) => heading.textContent === 'Visual direction',
      ),
    ).toBe(true)
    expect(container.querySelector('textarea')).toBeNull()
    await act(async () => button('在 WisWork PC 编辑').click())
    expect(request.mock.calls[0]![0]).toMatchObject({ action: 'open', markdown: original })
    await act(async () => vi.advanceTimersByTimeAsync(5000))
    expect(container.textContent).toContain('Volcano black')
    expect(container.textContent).toContain('已从 PC 同步')
    expect(apply).not.toHaveBeenCalled()
    await act(async () => button('应用设计修改').click())
    expect(apply).toHaveBeenCalledWith(saved)
  })

  it('does not apply a saved draft over a newer authoritative design', async () => {
    vi.useFakeTimers()
    const request = vi.fn(async (body: { documentId: string }) => ({
      documentId: body.documentId,
      markdown: saved,
      revision: 'new',
    }))
    const apply = vi.fn()
    await act(async () =>
      root.render(
        React.createElement(OfficeDesignPanel, { current, busy: false, request, onApply: apply }),
      ),
    )
    await act(async () => button('DESIGN.md').click())
    await act(async () => button('在 WisWork PC 编辑').click())
    await act(async () =>
      root.render(
        React.createElement(OfficeDesignPanel, {
          current: { markdown: '# DESIGN.md\nA newer design', sourceId: 'plan-2' },
          busy: false,
          request,
          onApply: apply,
        }),
      ),
    )
    expect(container.textContent).toContain('设计已更新')
    expect(button('应用设计修改').disabled).toBe(true)
    expect(apply).not.toHaveBeenCalled()
  })

  it('keeps reading available while busy and makes unavailable PC editing explicit', async () => {
    await act(async () =>
      root.render(
        React.createElement(OfficeDesignPanel, { current, busy: true, onApply: vi.fn() }),
      ),
    )
    await act(async () => button('DESIGN.md').click())
    expect(container.querySelector('[role="dialog"]')).not.toBeNull()
    expect(button('重新配对以启用编辑').disabled).toBe(true)
  })

  it('offers an explicit re-pair action when the connected PC lacks file sync', async () => {
    const onRepairConnection = vi.fn()
    await act(async () =>
      root.render(
        React.createElement(OfficeDesignPanel, {
          current,
          busy: false,
          onRepairConnection,
          onApply: vi.fn(),
        }),
      ),
    )
    await act(async () => button('DESIGN.md').click())
    expect(button('重新配对以启用编辑').disabled).toBe(false)
    await act(async () => button('重新配对以启用编辑').click())
    expect(onRepairConnection).toHaveBeenCalledOnce()
  })

  it('does not reopen a dismissed timeline selection when the current design changes', async () => {
    const selection = { markdown: original, editable: true }
    const render = (markdown: string, selected = selection) =>
      root.render(
        React.createElement(OfficeDesignPanel, {
          current: { markdown, sourceId: 'plan' },
          selection: selected,
          busy: false,
          onApply: vi.fn(),
        }),
      )
    await act(async () => render(original))
    expect(container.querySelector('[role="dialog"]')).not.toBeNull()
    await act(async () =>
      container.querySelector<HTMLButtonElement>('[aria-label="关闭设计文档"]')!.click(),
    )
    expect(container.querySelector('[role="dialog"]')).toBeNull()
    await act(async () => render(saved))
    expect(container.querySelector('[role="dialog"]')).toBeNull()
    await act(async () => render(saved, { ...selection }))
    expect(container.querySelector('[role="dialog"]')).not.toBeNull()
    expect(container.textContent).toContain('历史设计快照')
  })

  it('pauses PC polling when the reader is closed and resumes it on reopening', async () => {
    vi.useFakeTimers()
    const request = vi.fn(async (body: { documentId: string }) => ({
      documentId: body.documentId,
      markdown: original,
      revision: 'saved',
    }))
    await act(async () =>
      root.render(
        React.createElement(OfficeDesignPanel, { current, busy: false, request, onApply: vi.fn() }),
      ),
    )
    await act(async () => button('DESIGN.md').click())
    await act(async () => button('在 WisWork PC 编辑').click())
    await act(async () =>
      container.querySelector<HTMLButtonElement>('[aria-label="关闭设计文档"]')!.click(),
    )
    request.mockClear()
    await act(async () => vi.advanceTimersByTimeAsync(60_000))
    expect(request).not.toHaveBeenCalled()
    await act(async () => button('DESIGN.md').click())
    await act(async () => vi.advanceTimersByTimeAsync(5000))
    expect(request).toHaveBeenCalledOnce()
    expect(request).toHaveBeenCalledWith(
      expect.objectContaining({ action: 'read' }),
      expect.any(AbortSignal),
    )
  })
})
