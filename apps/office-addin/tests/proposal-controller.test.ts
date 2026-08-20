import { describe, expect, it, vi } from 'vitest'
import { createProposalController } from '../src/agent/proposal-controller.js'
import type { OfficeDocumentClient } from '../src/office-document.js'

function document(selection = 'before') {
  return {
    readSelection: vi.fn().mockResolvedValue(selection),
    replaceSelection: vi.fn().mockResolvedValue(undefined),
    appendText: vi.fn().mockResolvedValue(undefined),
  } as unknown as OfficeDocumentClient
}

describe('proposal controller', () => {
  it('keeps one pending proposal and captures selection state', async () => {
    const controller = createProposalController(document())
    const first = await controller.propose('replace', 'after')
    const second = await controller.propose('append', 'more')
    expect(controller.pending()).toEqual(second)
    expect(second.id).not.toBe(first.id)
    expect(second.before).toBe('before')
  })

  it.each(['reject', 'newTurn', 'logout'] as const)(
    '%s invalidates a pending proposal',
    async (name) => {
      const controller = createProposalController(document())
      await controller.propose('replace', 'after')
      controller[name]()
      expect(controller.pending()).toBeUndefined()
    },
  )

  it('revalidates selection and refuses a stale proposal without mutation', async () => {
    const doc = document()
    const controller = createProposalController(doc)
    const proposal = await controller.propose('replace', 'after')
    vi.mocked(doc.readSelection).mockResolvedValueOnce('changed')
    await expect(controller.confirm(proposal.id)).rejects.toThrow('proposal_stale')
    expect(doc.replaceSelection).not.toHaveBeenCalled()
    expect(controller.pending()).toBeUndefined()
  })

  it('mutates only after explicit confirmation and consumes the proposal once', async () => {
    const doc = document()
    const controller = createProposalController(doc)
    const proposal = await controller.propose('append', 'more')
    expect(doc.appendText).not.toHaveBeenCalled()
    await controller.confirm(proposal.id)
    expect(doc.appendText).toHaveBeenCalledWith('before', 'more')
    await expect(controller.confirm(proposal.id)).rejects.toThrow('proposal_missing')
  })
})
