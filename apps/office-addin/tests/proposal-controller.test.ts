import { describe, expect, it, vi } from 'vitest'
import {
  MAX_PROPOSAL_PREVIEW_BYTES,
  createProposalController,
  createStructuredProposalController,
} from '../src/agent/proposal-controller.js'
import type { OfficeDocumentClient } from '../src/office-document.js'

function document(selection = 'before') {
  let current = selection
  return {
    readSelection: vi.fn(async () => current),
    replaceSelection: vi.fn(async (value: string) => {
      current = value
    }),
    appendText: vi.fn(async (before: string, value: string) => {
      current = `${before}${value}`
    }),
  } as unknown as OfficeDocumentClient
}

describe('proposal controller', () => {
  it('publishes proposal lifecycle changes and exposes the eventual user decision', async () => {
    const controller = createStructuredProposalController()
    const snapshots: Array<string | undefined> = []
    const unsubscribe = controller.subscribe(() => snapshots.push(controller.pending()?.id))
    const proposal = controller.propose({
      operation: 'edit',
      title: 'Edit document',
      preview: {},
      impact: { host: 'word', targets: ['document'], count: 1 },
      fingerprint: 'v1',
      validate: async () => true,
      execute: async () => undefined,
    })
    const decision = controller.waitForDecision(proposal.id)

    expect(snapshots).toEqual([proposal.id])
    await controller.confirm(proposal.id)
    await expect(decision).resolves.toEqual({ status: 'confirmed' })
    expect(snapshots).toEqual([proposal.id, undefined])

    unsubscribe()
  })

  it.each([
    ['reject', 'rejected'],
    ['newTurn', 'cancelled'],
    ['logout', 'cancelled'],
  ] as const)('settles a suspended proposal as %s on %s', async (action, status) => {
    const controller = createStructuredProposalController()
    const proposal = controller.propose({
      operation: 'edit',
      title: 'Edit document',
      preview: {},
      impact: { host: 'word', targets: ['document'], count: 1 },
      fingerprint: 'v1',
      validate: async () => true,
      execute: async () => undefined,
    })
    const decision = controller.waitForDecision(proposal.id)
    controller[action]()
    await expect(decision).resolves.toEqual({ status })
  })

  it('settles a failed confirmation with its stable error code', async () => {
    const controller = createStructuredProposalController()
    const proposal = controller.propose({
      operation: 'edit',
      title: 'Edit document',
      preview: {},
      impact: { host: 'word', targets: ['document'], count: 1 },
      fingerprint: 'v1',
      validate: async () => false,
      execute: async () => undefined,
    })
    const decision = controller.waitForDecision(proposal.id)
    await expect(controller.confirm(proposal.id)).rejects.toThrow('proposal_stale')
    await expect(decision).resolves.toEqual({ status: 'failed', error: 'proposal_stale' })
  })

  it('never exposes an arbitrary Office error through a suspended decision', async () => {
    const record = vi.fn()
    const controller = createStructuredProposalController({ setTool: vi.fn(), record })
    const proposal = controller.propose({
      operation: 'edit',
      title: 'Edit document',
      preview: {},
      impact: { host: 'word', targets: ['document'], count: 1 },
      fingerprint: 'v1',
      validate: async () => true,
      execute: async () => {
        throw new Error('/Users/alice/private.docx access token secret')
      },
    })
    const decision = controller.waitForDecision(proposal.id)
    await expect(controller.confirm(proposal.id)).rejects.toThrow('alice')
    await expect(decision).resolves.toEqual({
      status: 'failed',
      error: 'office_write_failed',
    })
    expect(record).toHaveBeenCalledWith({
      phase: 'write',
      errorCode: 'office_write_failed',
      error: expect.objectContaining({ message: expect.stringContaining('alice') }),
    })
  })

  it('diagnoses validation and verification at their exact safe phases', async () => {
    const record = vi.fn()
    const validation = createStructuredProposalController({ setTool: vi.fn(), record })
    const stale = validation.propose({
      operation: 'write_document',
      toolName: 'write_document',
      title: 'Write',
      preview: {},
      impact: { host: 'word', targets: ['document'], count: 1 },
      fingerprint: 'v1',
      validate: async () => false,
      execute: async () => undefined,
    })
    await expect(validation.confirm(stale.id)).rejects.toThrow('proposal_stale')
    expect(record).toHaveBeenLastCalledWith({
      phase: 'validate',
      errorCode: 'proposal_stale',
      error: expect.any(Error),
    })

    const verification = createStructuredProposalController({ setTool: vi.fn(), record })
    const failed = verification.propose({
      operation: 'set_cell_range',
      toolName: 'set_cell_range',
      title: 'Write cells',
      preview: {},
      impact: { host: 'Excel', targets: ['sheet:1!A1'], count: 1 },
      fingerprint: 'v1',
      validate: async () => true,
      execute: async () => undefined,
      verify: async () => {
        throw new Error('office_verify_failed')
      },
    })
    await expect(verification.confirm(failed.id)).rejects.toThrow('office_verify_failed')
    expect(record).toHaveBeenLastCalledWith({
      phase: 'verify',
      errorCode: 'office_verify_failed',
      error: expect.any(Error),
    })
  })

  it('never lets a broken diagnostic sink replace the document failure', async () => {
    const controller = createStructuredProposalController({
      setTool: () => {
        throw new Error('diagnostic down')
      },
      record: () => {
        throw new Error('diagnostic down')
      },
    })
    const proposal = controller.propose({
      operation: 'edit',
      title: 'Edit',
      preview: {},
      impact: { host: 'word', targets: ['document'], count: 1 },
      fingerprint: 'v1',
      validate: async () => true,
      execute: async () => {
        throw new Error('office_write_failed')
      },
    })
    await expect(controller.confirm(proposal.id)).rejects.toThrow('office_write_failed')
  })

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

  it('rejects proposals and confirmations while another confirmation is active', async () => {
    let finishValidation!: (valid: boolean) => void
    const controller = createStructuredProposalController()
    const request = {
      operation: 'edit',
      title: 'Edit',
      preview: {},
      impact: { host: 'word', targets: [], count: 0 },
      fingerprint: 'v1',
      validate: () =>
        new Promise<boolean>((resolve) => {
          finishValidation = resolve
        }),
      execute: async () => undefined,
    }
    const proposal = controller.propose(request)
    const confirmation = controller.confirm(proposal.id)
    expect(() => controller.propose(request)).toThrow('proposal_confirmation_in_progress')
    await expect(controller.confirm(proposal.id)).rejects.toThrow(
      'proposal_confirmation_in_progress',
    )
    finishValidation(true)
    await confirmation
  })

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

  it('fails verification when a legacy Office write does not produce the previewed state', async () => {
    const doc = document()
    vi.mocked(doc.readSelection)
      .mockResolvedValueOnce('before')
      .mockResolvedValueOnce('before')
      .mockResolvedValueOnce('unchanged')
    const controller = createProposalController(doc)
    const proposal = await controller.propose('replace', 'after')
    await expect(controller.confirm(proposal.id)).rejects.toThrow('office_verify_failed')
    expect(doc.readSelection).toHaveBeenCalledTimes(3)
  })

  it('re-reads but never reports success when logout occurs during a legacy Office callback', async () => {
    const doc = document()
    let finish!: () => void
    vi.mocked(doc.readSelection)
      .mockResolvedValueOnce('before')
      .mockResolvedValueOnce('before')
      .mockResolvedValueOnce('after')
    vi.mocked(doc.replaceSelection).mockImplementation(
      () => new Promise<void>((resolve) => (finish = resolve)),
    )
    const controller = createProposalController(doc)
    const proposal = await controller.propose('replace', 'after')
    const confirmation = controller.confirm(proposal.id)
    const rejected = expect(confirmation).rejects.toThrow('proposal_stale')
    await vi.waitFor(() => expect(doc.replaceSelection).toHaveBeenCalledOnce())
    controller.logout()
    finish()
    await rejected
    expect(doc.readSelection).toHaveBeenCalledTimes(3)
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

  it('supports immutable structured previews, stale checks, and captured execution', async () => {
    const execute = vi.fn().mockResolvedValue(undefined)
    const verify = vi.fn().mockResolvedValue(undefined)
    const controller = createStructuredProposalController()
    const proposal = controller.propose({
      operation: 'set_cell_range',
      title: 'Set Sheet1!A1',
      preview: { range: 'A1', values: [['new']] },
      impact: { host: 'excel', targets: ['Sheet1!A1'], count: 1 },
      fingerprint: 'sheet-v1',
      validate: vi.fn().mockResolvedValue(true),
      execute,
      verify,
    })
    expect(Object.isFrozen(proposal)).toBe(true)
    expect(Object.isFrozen(proposal.preview)).toBe(true)
    expect(() => Object.assign(proposal.preview, { range: 'B2' })).toThrow()
    await controller.confirm(proposal.id)
    expect(execute).toHaveBeenCalledOnce()
    expect(verify).toHaveBeenCalledOnce()
    await expect(controller.confirm(proposal.id)).rejects.toThrow('proposal_missing')
  })

  it('fails closed for stale and concurrent structured confirmation', async () => {
    let finish!: () => void
    const execute = vi.fn(
      () =>
        new Promise<void>((resolve) => {
          finish = resolve
        }),
    )
    const controller = createStructuredProposalController()
    const stale = controller.propose({
      operation: 'edit',
      title: 'Edit',
      preview: {},
      impact: { host: 'word', targets: [], count: 0 },
      fingerprint: 'v1',
      validate: async () => false,
      execute,
    })
    await expect(controller.confirm(stale.id)).rejects.toThrow('proposal_stale')
    expect(execute).not.toHaveBeenCalled()

    const live = controller.propose({
      operation: 'edit',
      title: 'Edit',
      preview: {},
      impact: { host: 'word', targets: [], count: 0 },
      fingerprint: 'v2',
      validate: async () => true,
      execute,
    })
    const first = controller.confirm(live.id)
    await expect(controller.confirm(live.id)).rejects.toThrow('proposal_confirmation_in_progress')
    finish()
    await first
  })

  it('rejects unbounded or non-serializable structured previews', () => {
    const controller = createStructuredProposalController()
    const base = {
      operation: 'edit',
      title: 'Edit',
      impact: { host: 'word', targets: [], count: 0 },
      fingerprint: 'v1',
      validate: async () => true,
      execute: async () => undefined,
    }
    expect(() =>
      controller.propose({ ...base, preview: { text: 'x'.repeat(MAX_PROPOSAL_PREVIEW_BYTES) } }),
    ).toThrow('invalid_tool_input')
    const cyclic: Record<string, unknown> = {}
    cyclic.self = cyclic
    expect(() => controller.propose({ ...base, preview: cyclic })).toThrow('invalid_tool_input')
  })

  it.each(['reject', 'newTurn', 'logout'] as const)(
    '%s cancels confirmation while validation is in flight',
    async (invalidate) => {
      let finishValidation!: (valid: boolean) => void
      const execute = vi.fn()
      const controller = createStructuredProposalController()
      const proposal = controller.propose({
        operation: 'edit',
        title: 'Edit',
        preview: {},
        impact: { host: 'word', targets: [], count: 0 },
        fingerprint: 'v1',
        validate: () =>
          new Promise((resolve) => {
            finishValidation = resolve
          }),
        execute,
      })
      const confirmation = controller.confirm(proposal.id)
      controller[invalidate]()
      finishValidation(true)
      await expect(confirmation).rejects.toThrow('proposal_stale')
      expect(execute).not.toHaveBeenCalled()
    },
  )
})
