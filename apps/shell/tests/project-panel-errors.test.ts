// @vitest-environment jsdom

import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { ProjectPanel } from '../src/renderer/src/Home'
import type { ProjectHomeApi } from '../src/shared/home-api'

globalThis.IS_REACT_ACT_ENVIRONMENT = true

function setInputValue(input: HTMLInputElement, value: string) {
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set
  setter?.call(input, value)
  input.dispatchEvent(new Event('input', { bubbles: true }))
}

describe('ProjectPanel mutation failures', () => {
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
    delete window.aiOfficeProject
  })

  it('keeps a failed create retryable without rendering the raw IPC error', async () => {
    let rejectFirst!: (reason: unknown) => void
    const firstAttempt = new Promise<void>((_resolve, reject) => {
      rejectFirst = reject
    })
    const createProject = vi
      .fn<(...args: Parameters<ProjectHomeApi['createProject']>) => Promise<void>>()
      .mockReturnValueOnce(firstAttempt)
      .mockResolvedValueOnce(undefined)
    window.aiOfficeProject = { createProject } as ProjectHomeApi
    const onRefresh = vi.fn()

    await act(async () => {
      root.render(
        createElement(ProjectPanel, {
          projects: [],
          selectedId: null,
          onSelect: vi.fn(),
          onRefresh,
        }),
      )
    })
    await act(async () => {
      ;(container.querySelector('.proj-add-btn') as HTMLButtonElement).click()
    })
    const input = container.querySelector('.proj-rename-input') as HTMLInputElement
    await act(async () => setInputValue(input, 'Quarterly plan'))
    act(() => input.dispatchEvent(new FocusEvent('focusout', { bubbles: true })))

    expect(input.disabled).toBe(true)
    expect((container.querySelector('.proj-add-btn') as HTMLButtonElement).disabled).toBe(true)
    await act(async () => rejectFirst(new Error('/Users/alice/private token=secret-token')))

    expect(container.querySelector('[role="alert"]')?.textContent).toContain(
      'PROJECT_CREATE_FAILED',
    )
    expect(container.querySelector('[role="alert"]')?.textContent).toContain(
      '项目操作失败，请重试。',
    )
    expect(container.textContent).not.toContain('/Users/alice/private')
    expect(container.textContent).not.toContain('secret-token')
    expect(input.value).toBe('Quarterly plan')
    expect(input.disabled).toBe(false)
    expect(onRefresh).not.toHaveBeenCalled()

    await act(async () => input.dispatchEvent(new FocusEvent('focusout', { bubbles: true })))

    expect(createProject).toHaveBeenCalledTimes(2)
    expect(container.querySelector('[role="alert"]')).toBeNull()
    expect(container.querySelector('.proj-rename-input')).toBeNull()
    expect(onRefresh).toHaveBeenCalledTimes(1)
  })
})
