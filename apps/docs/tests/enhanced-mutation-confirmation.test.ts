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
  summary: { operation: 'replace', target: 'blocks', scope: 'bounded-set', count: 2 },
  expiresAt: Date.now() + 60_000,
  ...overrides,
})

function setup() {
  let listener: ((value: EnhancedMutationProposal) => void) | undefined
  const api = {
    onProposal: vi.fn((next: (value: EnhancedMutationProposal) => void) => {
      listener = next
      return vi.fn()
    }),
    confirmProposal: vi.fn(async () => undefined),
    cancelProposal: vi.fn(async () => undefined),
  }
  const node = document.createElement('div')
  document.body.append(node)
  const root = createRoot(node)
  act(() => root.render(createElement(EnhancedMutationConfirmation, { api, locale: 'zh' })))
  return { api, node, root, emit: (value: EnhancedMutationProposal) => act(() => listener?.(value)) }
}

afterEach(() => {
  document.body.replaceChildren()
})

describe('Enhanced mutation confirmation', () => {
  it('automatically confirms a valid bounded proposal without rendering a dialog', async () => {
    const view = setup()
    view.emit(proposal())
    await act(async () => undefined)

    expect(view.api.confirmProposal).toHaveBeenCalledOnce()
    expect(view.api.confirmProposal).toHaveBeenCalledWith('docs:document-1', 4, 'opaque-proposal-1')
    expect(view.api.cancelProposal).not.toHaveBeenCalled()
    expect(view.node.querySelector('[role="alertdialog"]')).toBeNull()
  })

  it('cancels expired and malformed proposals instead of approving them', async () => {
    const view = setup()
    view.emit(proposal({ expiresAt: Date.now() - 1 }))
    view.emit(proposal({ proposalId: 'malformed', summary: 'unsafe' as never }))
    await act(async () => undefined)

    expect(view.api.confirmProposal).not.toHaveBeenCalled()
    expect(view.api.cancelProposal).toHaveBeenCalledTimes(2)
  })

  it('consumes a proposal only once when an event is replayed', async () => {
    const view = setup()
    const candidate = proposal()
    view.emit(candidate)
    view.emit(candidate)
    await act(async () => undefined)

    expect(view.api.confirmProposal).toHaveBeenCalledOnce()
  })
})
