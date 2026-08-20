import type { AgentStreamCallbacks, AgentTransport } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { createOfficeAgentSession } from '../src/agent/use-office-agent.js'

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
  return {
    controller: {
      pending: () => pending,
      propose: vi.fn(),
      confirm: vi.fn(),
      reject: vi.fn(() => {
        pending = undefined
      }),
      newTurn: vi.fn(() => {
        pending = undefined
      }),
      logout: vi.fn(() => {
        pending = undefined
      }),
    },
    setPending() {
      pending = {
        id: 'p1',
        operation: 'replace' as const,
        before: 'old',
        value: 'new',
        fingerprint: 'x',
      }
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
    proposals.controller.confirm.mockRejectedValue(new Error('proposal_stale'))
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
    await session.confirm('p1')
    expect(session.snapshot().error).toBe('proposal_stale')
  })
})
