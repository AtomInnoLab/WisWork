import type { AgentStreamCallbacks, AgentTransport } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { bindAuthLoss, createOfficeAgentSession } from '../src/agent/use-office-agent.js'

function transportHarness() {
  let callbacks: AgentStreamCallbacks | undefined
  const cancel = vi.fn()
  const transport: AgentTransport = {
    stream(_request, next) {
      callbacks = next
      return { cancel }
    },
  }
  return { transport, cancel, callbacks: () => callbacks! }
}

function proposalsHarness() {
  let pending:
    | { id: string; operation: 'replace'; before: string; value: string; fingerprint: string }
    | undefined
  const controller = {
    pending: () => pending,
    propose: vi.fn(),
    confirm: vi.fn(async () => {
      pending = undefined
    }),
    reject: vi.fn(() => {
      pending = undefined
    }),
    newTurn: vi.fn(() => {
      pending = undefined
    }),
    logout: vi.fn(() => {
      pending = undefined
    }),
  }
  return {
    controller,
    setPending() {
      pending = {
        id: 'p1',
        operation: 'replace' as const,
        before: 'old',
        value: 'new',
        fingerprint: 'x',
      }
    },
    clearPending() {
      pending = undefined
    },
  }
}

describe('Office agent session', () => {
  it('streams assistant text and reports completion', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    session.send('Summarize this')
    await Promise.resolve()
    harness.callbacks().onDelta('Hello')
    harness.callbacks().onDelta(' world')
    harness.callbacks().onDone()

    expect(session.snapshot()).toMatchObject({
      assistantText: 'Hello world',
      busy: false,
      status: 'done',
    })
  })

  it('invalidates a pending proposal before a new instruction and on logout', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    proposals.setPending()
    session.send('new request')
    await Promise.resolve()
    expect(proposals.controller.newTurn).toHaveBeenCalledOnce()
    harness.callbacks().onDone()
    proposals.setPending()
    session.logout()

    expect(proposals.controller.logout).toHaveBeenCalledOnce()
    expect(session.snapshot()).toMatchObject({
      assistantText: '',
      proposal: undefined,
      busy: false,
    })
  })

  it('stops the active stream and surfaces safe proposal confirmation errors', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    proposals.controller.confirm.mockImplementation(async () => {
      proposals.clearPending()
      throw new Error('proposal_stale')
    })
    proposals.setPending()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    session.send('work')
    await Promise.resolve()
    session.stop()
    expect(harness.cancel).toHaveBeenCalledOnce()
    harness.callbacks().onDone()
    await session.confirm('p1')
    expect(session.snapshot().error).toBe('proposal_stale')
    expect(session.snapshot().proposal).toBeUndefined()
  })

  it('allows only one confirmation and blocks competing actions until the write settles', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    proposals.setPending()
    let settle!: () => void
    proposals.controller.confirm.mockImplementation(
      () =>
        new Promise<void>((resolve) => {
          settle = () => {
            proposals.clearPending()
            resolve()
          }
        }),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    const first = session.confirm('p1')
    const second = session.confirm('p1')
    session.send('race')
    session.reject()
    session.stop()
    session.logout()

    expect(session.snapshot().applying).toBe(true)
    expect(proposals.controller.confirm).toHaveBeenCalledOnce()
    expect(proposals.controller.newTurn).not.toHaveBeenCalled()
    expect(proposals.controller.reject).not.toHaveBeenCalled()
    expect(proposals.controller.logout).not.toHaveBeenCalled()
    expect(harness.cancel).not.toHaveBeenCalled()

    settle()
    await Promise.all([first, second])
    expect(session.snapshot().applying).toBe(false)
    expect(session.snapshot().proposal).toBeUndefined()
    session.logout()
    expect(proposals.controller.logout).toHaveBeenCalledOnce()
  })

  it('does not confirm a visible proposal until the active agent run finishes', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    session.send('prepare an edit')
    await Promise.resolve()
    proposals.setPending()
    harness.callbacks().onDelta('Review this change')

    await session.confirm('p1')
    expect(session.snapshot().busy).toBe(true)
    expect(proposals.controller.confirm).not.toHaveBeenCalled()

    harness.callbacks().onDone()
    await session.confirm('p1')
    expect(proposals.controller.confirm).toHaveBeenCalledOnce()
  })

  it('atomically resets history and proposals when authentication is lost', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    proposals.setPending()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })
    let authLoss: (() => void) | undefined
    const signedOut = vi.fn()
    const disconnect = bindAuthLoss(
      { subscribeAuthLoss: (listener) => ((authLoss = listener), () => (authLoss = undefined)) },
      session,
      signedOut,
    )

    session.send('active request')
    await Promise.resolve()
    authLoss?.()

    expect(harness.cancel).toHaveBeenCalledOnce()
    expect(proposals.controller.logout).toHaveBeenCalledOnce()
    expect(session.snapshot()).toMatchObject({
      busy: false,
      proposal: undefined,
      assistantText: '',
    })
    expect(signedOut).toHaveBeenCalledOnce()
    disconnect()
    expect(authLoss).toBeUndefined()
  })
})
