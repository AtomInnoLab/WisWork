// @vitest-environment jsdom
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { EnhancedMutationConfirmation, type EnhancedMutationProposal } from '@wiswork/ui'

;(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true

const proposal = (overrides: Partial<EnhancedMutationProposal> = {}): EnhancedMutationProposal => ({
  proposalId: 'opaque-proposal-1',
  documentId: 'docs:document-1',
  generation: 4,
  toolName: 'replace_blocks',
  summary: 'Replace 2 document blocks',
  expiresAt: Date.now() + 60_000,
  ...overrides,
})

function setup() {
  let listener: ((value: EnhancedMutationProposal) => void) | undefined
  const unsubscribe = vi.fn()
  const api = {
    onProposal: vi.fn((next: (value: EnhancedMutationProposal) => void) => {
      listener = next
      return unsubscribe
    }),
    confirmProposal: vi.fn(async () => undefined),
    cancelProposal: vi.fn(async () => undefined),
  }
  const node = document.createElement('div')
  document.body.append(node)
  const root = createRoot(node)
  act(() => root.render(createElement(EnhancedMutationConfirmation, { api })))
  return {
    api,
    node,
    root,
    unsubscribe,
    emit: (value: EnhancedMutationProposal) => act(() => listener?.(value)),
  }
}

afterEach(() => {
  document.body.replaceChildren()
  vi.useRealTimers()
})

describe('Enhanced mutation confirmation', () => {
  it('keeps a proposal pending until one explicit confirmation', async () => {
    const view = setup()
    view.emit(proposal())
    expect(view.node.textContent).toContain('Replace 2 document blocks')
    expect(view.api.confirmProposal).not.toHaveBeenCalled()

    const confirm = view.node.querySelector<HTMLButtonElement>('[data-action="confirm"]')!
    await act(async () => confirm.click())
    await act(async () => confirm.click())

    expect(view.api.confirmProposal).toHaveBeenCalledTimes(1)
    expect(view.api.confirmProposal).toHaveBeenCalledWith('docs:document-1', 4, 'opaque-proposal-1')
  })

  it('cancels on rejection and removes the prompt without confirming', async () => {
    const view = setup()
    view.emit(proposal())
    await act(async () =>
      view.node.querySelector<HTMLButtonElement>('[data-action="cancel"]')!.click(),
    )
    expect(view.api.cancelProposal).toHaveBeenCalledWith('docs:document-1', 4, 'opaque-proposal-1')
    expect(view.api.confirmProposal).not.toHaveBeenCalled()
    expect(view.node.querySelector('[role="alertdialog"]')).toBeNull()
  })

  it('cancels an expired proposal and a pending proposal on unmount', async () => {
    vi.useFakeTimers()
    vi.setSystemTime(new Date('2026-09-02T00:00:00Z'))
    const view = setup()
    view.emit(proposal({ expiresAt: Date.now() + 1_000 }))
    await act(async () => vi.advanceTimersByTimeAsync(1_001))
    expect(view.api.cancelProposal).toHaveBeenCalledTimes(1)

    view.emit(proposal({ proposalId: 'opaque-proposal-2', expiresAt: Date.now() + 5_000 }))
    act(() => view.root.unmount())
    await act(async () => undefined)
    expect(view.api.cancelProposal).toHaveBeenCalledWith('docs:document-1', 4, 'opaque-proposal-2')
    expect(view.unsubscribe).toHaveBeenCalledOnce()
  })

  it('bounds displayed data and ignores malformed, replayed, and stale replacement events', async () => {
    const view = setup()
    view.emit(proposal({ summary: 'x'.repeat(2_000) }))
    expect(view.node.textContent?.length).toBeLessThan(1_000)

    const first = proposal()
    view.emit(first)
    view.emit(first)
    expect(view.node.textContent).toContain('Replace 2 document blocks')

    view.emit(proposal({ proposalId: '', documentId: '../secret' }))
    expect(view.node.textContent).toContain('Replace 2 document blocks')
    expect(view.api.confirmProposal).not.toHaveBeenCalled()
  })
})
