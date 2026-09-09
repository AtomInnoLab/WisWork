// @vitest-environment jsdom

import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ImageSearchSettings } from '../src/renderer/src/Home'
import type { HomeApi } from '../src/shared/home-api'

globalThis.IS_REACT_ACT_ENVIRONMENT = true

function setInputValue(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set
  setter?.call(input, value)
  input.dispatchEvent(new Event('input', { bubbles: true }))
}

describe('image search settings UI', () => {
  let container: HTMLDivElement
  let root: ReturnType<typeof createRoot>

  beforeEach(() => {
    container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
  })

  afterEach(async () => {
    await act(async () => root.unmount())
    container.remove()
  })

  it('shows status and supports password-only save, test, and clear without reading the key', async () => {
    const saveImageSearchKey = vi.fn(async () => ({ configured: true, source: 'stored' as const }))
    const clearImageSearchKey = vi.fn(async () => ({ configured: false, source: 'none' as const }))
    const testImageSearchKey = vi.fn(async () => ({ ok: true, resultCount: 1 }))
    window.aiOffice = {
      imageSearchKeyStatus: vi.fn(async () => ({ configured: false, source: 'none' as const })),
      saveImageSearchKey,
      clearImageSearchKey,
      testImageSearchKey,
    } as unknown as HomeApi

    await act(async () => root.render(createElement(ImageSearchSettings, { language: 'en' })))
    expect(container.textContent).toContain('Not configured')
    const input = container.querySelector('input') as HTMLInputElement
    expect(input.type).toBe('password')
    await act(async () => setInputValue(input, 'secret-key'))
    await act(async () =>
      (container.querySelector('[data-action="save"]') as HTMLButtonElement).click(),
    )
    expect(saveImageSearchKey).toHaveBeenCalledWith('secret-key')
    expect(input.value).toBe('')
    expect(container.textContent).not.toContain('secret-key')

    await act(async () =>
      (container.querySelector('[data-action="test"]') as HTMLButtonElement).click(),
    )
    expect(testImageSearchKey).toHaveBeenCalledOnce()
    expect(container.textContent).toContain('Connection successful')

    await act(async () =>
      (container.querySelector('[data-action="clear"]') as HTMLButtonElement).click(),
    )
    expect(clearImageSearchKey).toHaveBeenCalledOnce()
    expect(container.textContent).toContain('Not configured')
  })
})
