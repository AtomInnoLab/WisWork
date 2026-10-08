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
  it('quarantines writes until the matching generation proves stable', () => {
    const controller = createStructuredProposalController()
    const request = {
      operation: 'edit',
      title: 'Edit document',
      preview: {},
      impact: { host: 'word', targets: ['document'], count: 1 },
      fingerprint: 'v1',
      validate: async () => true,
      execute: async () => undefined,
    }
    const lease = controller.quarantine({ sessionId: 'session-a', generation: 1 })

    expect(controller.isQuarantined()).toBe(true)
    expect(() => controller.propose(request)).toThrow('office_state_uncertain')
    controller.newTurn()
    expect(() => controller.propose(request)).toThrow('office_state_uncertain')
    controller.resolveQuarantine(lease, { stable: true })
    expect(controller.isQuarantined()).toBe(false)
    expect(controller.propose(request)).toBeDefined()
  })

  it('does not let an old reconciliation clear a replacement session quarantine', () => {
    const controller = createStructuredProposalController()
    const oldLease = controller.quarantine({ sessionId: 'session-a', generation: 1 })
    controller.logout()
    expect(controller.isQuarantined()).toBe(true)
    controller.resolveQuarantine(oldLease, { stable: false })
    expect(controller.isQuarantined()).toBe(true)

    controller.destroyDocumentContext()
    expect(controller.isQuarantined()).toBe(false)

    const currentLease = controller.quarantine({ sessionId: 'session-b', generation: 2 })
    controller.resolveQuarantine(oldLease, { stable: true })
    expect(controller.isQuarantined()).toBe(true)
    controller.resolveQuarantine(currentLease, { stable: false })
    expect(controller.isQuarantined()).toBe(true)
  })

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
      durationMs: expect.any(Number),
    })
  })

  it.each([
    'https://private.example/token',
    '/Users/alice/private.docx',
    'PowerPoint.operations.0.set_shape_text_style.secret',
    'PowerPoint.operations.1234.set_shape_text_style.fontFamily',
  ])('does not forward an unrecognized verification location: %s', async (errorLocation) => {
    const controller = createStructuredProposalController()
    const proposal = controller.propose({
      operation: 'execute_office_js',
      title: 'Edit',
      preview: {},
      impact: { host: 'powerpoint', targets: ['slide'], count: 1 },
      fingerprint: 'v1',
      validate: () => true,
      execute: () => undefined,
      verify: () => {
        throw Object.assign(new Error('office_verify_failed'), { debugInfo: { errorLocation } })
      },
    })
    const decision = controller.waitForDecision(proposal.id)
    await expect(controller.confirm(proposal.id)).rejects.toThrow('office_verify_failed')
    expect(await decision).toEqual({ status: 'failed', error: 'office_verify_failed' })
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
      durationMs: expect.any(Number),
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
      durationMs: expect.any(Number),
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

  it('preserves the actionable Excel overwrite-required failure', async () => {
    const controller = createStructuredProposalController()
    const proposal = controller.propose({
      operation: 'set_cell_range',
      toolName: 'set_cell_range',
      title: 'Write cells',
      preview: {},
      impact: { host: 'Excel', targets: ['sheet:1!G1'], count: 1 },
      fingerprint: 'v1',
      validate: async () => true,
      execute: async () => {
        throw new Error('office_overwrite_required')
      },
    })

    await expect(controller.confirm(proposal.id)).rejects.toThrow('office_overwrite_required')
    await expect(controller.waitForDecision(proposal.id)).rejects.toThrow('proposal_missing')
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
      const decision = controller.waitForDecision(proposal.id)
      const confirmation = controller.confirm(proposal.id)
      controller[invalidate]()
      finishValidation(true)
      await expect(confirmation).resolves.toBeUndefined()
      await expect(decision).resolves.toEqual({ status: 'cancelled' })
      expect(execute).not.toHaveBeenCalled()
    },
  )

  it('reconciles a write when cancellation races after execute resolves', async () => {
    const controller = createStructuredProposalController()
    const verify = vi.fn(async (signal?: AbortSignal) => {
      expect(signal).toBeUndefined()
    })
    const proposal = controller.propose({
      operation: 'edit',
      title: 'Edit',
      preview: {},
      impact: { host: 'word', targets: [], count: 0 },
      fingerprint: 'v1',
      validate: async () => true,
      execute: async () => controller.newTurn(),
      verify,
    })

    await expect(controller.confirm(proposal.id)).resolves.toBeUndefined()
    expect(verify).toHaveBeenCalledOnce()
  })
})

describe('structured proposal write hooks', () => {
  function request() {
    return {
      operation: 'edit',
      title: 'Edit page',
      preview: {},
      impact: { host: 'powerpoint', targets: ['slide'], count: 1 },
      fingerprint: 'v1',
      validate: vi.fn(async () => true),
      execute: vi.fn(async () => {}),
      verify: vi.fn(async () => {}),
    }
  }
  it('runs hooks around write and verification after successful validation', async () => {
    const order: string[] = []
    const beforeWrite = vi.fn(async () => {
      order.push('before')
    })
    const afterWrite = vi.fn(() => {
      order.push('after')
    })
    const controller = createStructuredProposalController(undefined, { beforeWrite, afterWrite })
    const input = request()
    input.validate.mockImplementation(async () => {
      order.push('validate')
      return true
    })
    input.execute.mockImplementation(async () => {
      order.push('execute')
    })
    input.verify.mockImplementation(async () => {
      order.push('verify')
    })
    const proposal = controller.propose(input),
      decision = controller.waitForDecision(proposal.id)
    await controller.confirm(proposal.id)
    expect(order).toEqual(['validate', 'before', 'validate', 'execute', 'verify', 'after'])
    expect(beforeWrite).toHaveBeenCalledWith(proposal, expect.any(AbortSignal))
    await expect(decision).resolves.toEqual({ status: 'confirmed' })
  })
  it.each(['reject', 'invalid', 'validation_error'] as const)(
    'skips hooks for %s',
    async (outcome) => {
      const hooks = { beforeWrite: vi.fn(async () => {}), afterWrite: vi.fn() }
      const controller = createStructuredProposalController(undefined, hooks),
        input = request()
      if (outcome === 'invalid') input.validate.mockResolvedValue(false)
      if (outcome === 'validation_error')
        input.validate.mockRejectedValue(new Error('office_verify_failed'))
      const proposal = controller.propose(input),
        decision = controller.waitForDecision(proposal.id)
      if (outcome === 'reject') controller.reject()
      else await expect(controller.confirm(proposal.id)).rejects.toThrow()
      await decision
      expect(hooks.beforeWrite).not.toHaveBeenCalled()
      expect(hooks.afterWrite).not.toHaveBeenCalled()
      expect(input.execute).not.toHaveBeenCalled()
    },
  )
  it.each(['before', 'execute', 'verify', 'after'] as const)(
    'releases after %s failure without confirming',
    async (failure) => {
      const input = request()
      const hooks = { beforeWrite: vi.fn(async () => {}), afterWrite: vi.fn() }
      if (failure === 'before')
        hooks.beforeWrite.mockRejectedValue(new Error('office_write_failed'))
      if (failure === 'execute') input.execute.mockRejectedValue(new Error('office_write_failed'))
      if (failure === 'verify') input.verify.mockRejectedValue(new Error('office_verify_failed'))
      if (failure === 'after')
        hooks.afterWrite.mockImplementation(() => {
          throw new Error('release failed')
        })
      const controller = createStructuredProposalController(undefined, hooks)
      const proposal = controller.propose(input),
        decision = controller.waitForDecision(proposal.id)
      await expect(controller.confirm(proposal.id)).rejects.toThrow()
      expect(hooks.afterWrite).toHaveBeenCalledTimes(1)
      if (failure === 'before') expect(input.execute).not.toHaveBeenCalled()
      await expect(decision).resolves.toMatchObject({ status: 'failed' })
      expect(() => controller.propose(request())).not.toThrow()
    },
  )
  it('releases and prevents execution when cancellation races the before-write hook', async () => {
    let entered!: () => void, release!: () => void
    const ready = new Promise<void>((resolve) => {
      entered = resolve
    })
    const pending = new Promise<void>((resolve) => {
      release = resolve
    })
    const hooks = {
      beforeWrite: vi.fn(async () => {
        entered()
        await pending
      }),
      afterWrite: vi.fn(),
    }
    const controller = createStructuredProposalController(undefined, hooks),
      input = request()
    const proposal = controller.propose(input),
      decision = controller.waitForDecision(proposal.id)
    const confirmation = controller.confirm(proposal.id)
    await ready
    controller.newTurn()
    release()
    await expect(confirmation).rejects.toThrow('proposal_stale')
    await expect(decision).resolves.toEqual({ status: 'failed', error: 'proposal_stale' })
    expect(input.execute).not.toHaveBeenCalled()
    expect(hooks.afterWrite).toHaveBeenCalledTimes(1)
  })
  it('revalidates after persistence and prevents stale writes when the document changed', async () => {
    let valid = true
    const input = request()
    input.validate.mockImplementation(async () => valid)
    const hooks = {
      beforeWrite: vi.fn(async () => {
        valid = false
      }),
      afterWrite: vi.fn(),
    }
    const controller = createStructuredProposalController(undefined, hooks)
    const proposal = controller.propose(input),
      decision = controller.waitForDecision(proposal.id)
    await expect(controller.confirm(proposal.id)).rejects.toThrow('proposal_stale')
    expect(input.validate).toHaveBeenCalledTimes(2)
    expect(input.execute).not.toHaveBeenCalled()
    expect(hooks.afterWrite).toHaveBeenCalledTimes(1)
    await expect(decision).resolves.toEqual({ status: 'failed', error: 'proposal_stale' })
  })
  it('still verifies and releases when cancellation races a committed write', async () => {
    const hooks = { beforeWrite: vi.fn(async () => {}), afterWrite: vi.fn() }
    const controller = createStructuredProposalController(undefined, hooks)
    const input = request()
    input.execute.mockImplementation(async () => {
      controller.newTurn()
    })
    const proposal = controller.propose(input),
      decision = controller.waitForDecision(proposal.id)
    await controller.confirm(proposal.id)
    expect(input.verify).toHaveBeenCalledWith(undefined)
    expect(hooks.afterWrite).toHaveBeenCalledTimes(1)
    await expect(decision).resolves.toEqual({ status: 'confirmed' })
  })
  it('preserves the primary write failure when release also throws', async () => {
    const input = request()
    input.execute.mockRejectedValue(new Error('office_state_uncertain'))
    const controller = createStructuredProposalController(undefined, {
      beforeWrite: async () => {},
      afterWrite: () => {
        throw new Error('cleanup')
      },
    })
    const proposal = controller.propose(input),
      decision = controller.waitForDecision(proposal.id)
    await expect(controller.confirm(proposal.id)).rejects.toThrow('office_state_uncertain')
    await expect(decision).resolves.toEqual({ status: 'failed', error: 'office_state_uncertain' })
  })
})

describe('post-write evidence', () => {
  const pngBase64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
  const digest = 'f4b555ad4009f54a1a37dc29e7ccf9f8f4cfe22410ba7769061c0328cdb6db67'
  const request = () => ({
    operation: 'edit',
    title: 'Edit page',
    preview: {},
    impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
    fingerprint: 'v1',
    validate: vi.fn(async () => true),
    execute: vi.fn(async () => {}),
    verify: vi.fn(async () => {}),
    postWrite: vi.fn(async () => ({
      status: 'captured' as const,
      pages: [{ slideId: 'slide-1', pngBase64, digest }],
    })),
  })

  it('runs capture after successful target verification and returns bounded evidence', async () => {
    const order: string[] = []
    const input = request()
    input.execute.mockImplementation(async () => {
      order.push('execute')
    })
    input.verify.mockImplementation(async () => {
      order.push('verify')
    })
    input.postWrite.mockImplementation(async () => {
      order.push('postWrite')
      return { status: 'captured', pages: [{ slideId: 'slide-1', pngBase64, digest }] }
    })
    const controller = createStructuredProposalController()
    const proposal = controller.propose(input)
    const decision = controller.waitForDecision(proposal.id)
    await controller.confirm(proposal.id)
    expect(order).toEqual(['execute', 'verify', 'postWrite'])
    await expect(decision).resolves.toEqual({
      status: 'confirmed',
      postWrite: { status: 'captured', pages: [{ slideId: 'slide-1', pngBase64, digest }] },
    })
  })

  it.each(['reject', 'verify_failed'] as const)('never captures after %s', async (mode) => {
    const input = request()
    if (mode === 'verify_failed') input.verify.mockRejectedValue(new Error('office_verify_failed'))
    const controller = createStructuredProposalController()
    const proposal = controller.propose(input)
    const decision = controller.waitForDecision(proposal.id)
    if (mode === 'reject') controller.reject()
    else await expect(controller.confirm(proposal.id)).rejects.toThrow('office_verify_failed')
    await decision
    expect(input.postWrite).not.toHaveBeenCalled()
  })

  it('keeps a successful write confirmed if capture fails or produces malformed evidence', async () => {
    for (const result of [
      new Error('capture error'),
      { status: 'captured', pages: [{ slideId: 'slide-1', pngBase64: 'bad', digest }] },
      {
        status: 'captured',
        pages: [{ slideId: 'slide-1', pngBase64: 'A'.repeat(3 * 1024 * 1024), digest }],
      },
    ]) {
      const input = request()
      input.postWrite.mockImplementation(async () => {
        if (result instanceof Error) throw result
        return result as Awaited<ReturnType<typeof input.postWrite>>
      })
      const controller = createStructuredProposalController()
      const proposal = controller.propose(input)
      const decision = controller.waitForDecision(proposal.id)
      await expect(controller.confirm(proposal.id)).resolves.toBeUndefined()
      await expect(decision).resolves.toMatchObject({
        status: 'confirmed',
        postWrite: { status: 'unavailable' },
      })
    }
  })
})

it('preserves the stable backup capacity code in a failed proposal decision', async () => {
  const controller = createStructuredProposalController()
  const proposal = controller.propose({
    operation: 'stage_existing_presentation_page_change',
    title: 'Stage page',
    preview: {},
    impact: { host: 'powerpoint', targets: ['slide'], count: 1 },
    fingerprint: 'v1',
    validate: () => true,
    execute: () => {
      throw new Error('presentation_existing_backup_capacity')
    },
  })
  const decision = controller.waitForDecision(proposal.id)
  await expect(controller.confirm(proposal.id)).rejects.toThrow(
    'presentation_existing_backup_capacity',
  )
  await expect(decision).resolves.toEqual({
    status: 'failed',
    error: 'presentation_existing_backup_capacity',
  })
})

it('keeps all 600 mutation targets in a bounded proposal snapshot', () => {
  const controller = createStructuredProposalController()
  const targets = Array.from({ length: 600 }, (_, i) => `page-${i}`)
  const proposal = controller.propose({
    operation: 'edit_slide_master_xml',
    toolName: 'edit_slide_master_xml',
    title: 'Change master with complete page scope',
    preview: { qaScope: { basis: 'master_xml_savepoint', hostSlideIds: targets } },
    impact: { host: 'powerpoint', targets, count: 600 },
    fingerprint: 'master-proof',
    validate: async () => true,
    execute: async () => {},
  })
  targets[599] = 'mutated-alias'
  expect(proposal.impact.targets).toHaveLength(600)
  expect(proposal.impact.targets[599]).toBe('page-599')
  expect(controller.pending()!.impact.targets).toEqual(proposal.impact.targets)
})

it('rejects oversized target snapshots before replacing the pending proposal', () => {
  const controller = createStructuredProposalController()
  expect(() =>
    controller.propose({
      operation: 'edit_slide_master_xml',
      title: 'Oversized master scope',
      preview: {},
      impact: {
        host: 'powerpoint',
        targets: Array.from({ length: 600 }, (_, i) => `${i}-${'x'.repeat(200)}`),
        count: 600,
      },
      fingerprint: 'master-proof',
      validate: async () => true,
      execute: async () => {},
    }),
  ).toThrow('invalid_tool_input')
  expect(controller.pending()).toBeUndefined()
})
