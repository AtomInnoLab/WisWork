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

  it('does not let mutation of the propose result change the confirmed write', async () => {
    const doc = document()
    const controller = createProposalController(doc)
    const proposal = await controller.propose('replace', 'original preview')
    expect(Object.isFrozen(proposal)).toBe(true)
    expect(() => Object.assign(proposal, { value: 'attacker value', before: 'changed' })).toThrow()
    await controller.confirm(proposal.id)
    expect(doc.replaceSelection).toHaveBeenCalledWith('original preview')
  })

  it('returns an immutable pending snapshot detached from internal confirmation state', async () => {
    const doc = document()
    const controller = createProposalController(doc)
    const proposed = await controller.propose('append', ' original')
    const pending = controller.pending()!
    expect(pending).not.toBe(proposed)
    expect(Object.isFrozen(pending)).toBe(true)
    expect(() => Object.assign(pending, { operation: 'replace', value: ' attacker' })).toThrow()
    await controller.confirm(proposed.id)
    expect(doc.appendText).toHaveBeenCalledWith('before', ' original')
    expect(doc.replaceSelection).not.toHaveBeenCalled()
  })
})
