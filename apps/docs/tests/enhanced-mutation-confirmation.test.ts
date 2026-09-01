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

function setup(locale: Parameters<typeof EnhancedMutationConfirmation>[0]['locale'] = 'en') {
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
  act(() => root.render(createElement(EnhancedMutationConfirmation, { api, locale })))
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
  document.documentElement.lang = 'en'
  vi.useRealTimers()
})

describe('Enhanced mutation confirmation', () => {
  it('keeps a proposal pending until one explicit confirmation', async () => {
    const view = setup()
    view.emit(proposal())
    expect(view.node.textContent).toContain('Replace')
    expect(view.node.textContent).toContain('Blocks')
    expect(view.node.textContent).toContain('Bounded set')
    expect(view.node.textContent).toContain('2')
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

  it('cancels malformed or generic summaries and ignores replayed events', async () => {
    const view = setup()
    view.emit(proposal({ summary: 'Review the proposed document change' as never }))
    expect(view.node.querySelector('[role="alertdialog"]')).toBeNull()
    expect(view.api.cancelProposal).toHaveBeenCalledWith('docs:document-1', 4, 'opaque-proposal-1')

    const first = proposal({ proposalId: 'opaque-proposal-2' })
    view.emit(first)
    view.emit(first)
    expect(view.node.textContent).toContain('Replace')

    view.emit(proposal({ proposalId: '', documentId: '../secret', summary: undefined as never }))
    expect(view.node.textContent).toContain('Replace')
    expect(view.api.confirmProposal).not.toHaveBeenCalled()
  })

  it('renders the informed summary and consent controls in Chinese', () => {
    document.documentElement.lang = 'en-US'
    const view = setup('zh')
    view.emit(
      proposal({
        summary: { operation: 'format', target: 'cells', scope: 'selection', count: 12 },
      }),
    )
    expect(view.node.textContent).toContain('确认文档更改')
    expect(view.node.textContent).toContain('格式调整')
    expect(view.node.textContent).toContain('单元格')
    expect(view.node.textContent).toContain('当前选区')
    expect(view.node.textContent).toContain('12')
    expect(view.node.textContent).toContain('拒绝')
  })

  it('uses the explicit locale rather than inferring document.lang', () => {
    document.documentElement.lang = 'en-US'
    const view = setup('ja')
    view.emit(proposal())
    expect(view.node.textContent).toContain('文書の変更を確認')
    expect(view.node.textContent).toContain('置換')
    expect(view.node.textContent).not.toContain('Confirm document change')
  })
})
