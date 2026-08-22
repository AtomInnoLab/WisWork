import type { AgentStreamCallbacks, AgentTransport, ToolExecution } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { bindAuthLoss, createOfficeAgentSession } from '../src/agent/use-office-agent.js'
import type { StructuredProposal } from '../src/agent/proposal-controller.js'

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
  it('keeps a bounded immutable two-turn user and assistant presentation timeline', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [
          {
            name: 'propose_replace_selection',
            description: 'prepare a proposal',
            inputSchema: { type: 'object' },
          },
        ],
        executeTool: vi.fn(async () => ({ output: 'prepared', summary: 'Prepared edit' })),
      },
      proposals: proposals.controller,
    })

    session.send('First question')
    await Promise.resolve()
    harness.callbacks().onDelta('First')
    harness.callbacks().onDelta(' answer')
    harness.callbacks().onDone()
    session.send('Second question')
    await Promise.resolve()
    harness.callbacks().onDelta('Second answer')
    harness.callbacks().onDone()

    const timeline = session.snapshot().timeline
    expect(timeline.map(({ kind }) => kind)).toEqual(['user', 'assistant', 'user', 'assistant'])
    expect(timeline.map((event) => ('text' in event ? event.text : ''))).toEqual([
      'First question',
      'First answer',
      'Second question',
      'Second answer',
    ])
    expect(Object.isFrozen(timeline)).toBe(true)
    expect(Object.isFrozen(timeline[0])).toBe(true)
  })

  it('replaces only the active assistant event while streaming', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Stream')
    await Promise.resolve()
    harness.callbacks().onDelta('A')
    const first = session.snapshot().timeline
    harness.callbacks().onDelta('B')
    const second = session.snapshot().timeline

    expect(first).toHaveLength(2)
    expect(second).toHaveLength(2)
    expect(second[0]).toBe(first[0])
    expect(second[1]).toMatchObject({ id: first[1]?.id, kind: 'assistant', text: 'AB' })
  })

  it('places completed tool work and its proposal inline before the following assistant turn', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [
          {
            name: 'propose_replace_selection',
            description: 'prepare a proposal',
            inputSchema: { type: 'object' },
          },
        ],
        executeTool: vi.fn(async () => ({ output: 'prepared', summary: 'Prepared edit' })),
      },
      proposals: proposals.controller,
    })

    session.send('Edit this')
    await Promise.resolve()
    harness.callbacks().onDelta('I will prepare it.')
    harness.callbacks().onToolCall({ id: 'tool-1', name: 'propose_replace_selection', input: {} })
    proposals.setPending()
    harness.callbacks().onDone()
    await Promise.resolve()
    harness.callbacks().onDelta('Ready for review.')
    harness.callbacks().onDone()

    expect(session.snapshot().timeline.map(({ kind }) => kind)).toEqual([
      'user',
      'assistant',
      'tool',
      'proposal',
      'assistant',
    ])
    expect(session.snapshot().timeline[2]).toMatchObject({
      kind: 'tool',
      callId: 'tool-1',
      state: 'complete',
    })
    expect(session.snapshot().timeline[3]).toMatchObject({
      kind: 'proposal',
      proposal: { id: 'p1' },
      state: 'pending',
    })

    await session.confirm('p1')
    expect(session.snapshot().timeline[3]).toMatchObject({ kind: 'proposal', state: 'applied' })
  })

  it('never exposes internal tool identifiers while a tool is running or fails', async () => {
    const harness = transportHarness()
    let failTool!: () => void
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'bash', description: 'internal', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(
          () =>
            new Promise<ToolExecution>((resolve) => {
              failTool = () => resolve({ output: 'sandbox_denied', isError: true, summary: 'bash' })
            }),
        ),
      },
      proposals: proposalsHarness().controller,
    })

    session.send('Open my attachment')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'tool-private', name: 'bash', input: {} })
    harness.callbacks().onDone()
    await Promise.resolve()
    expect(JSON.stringify(session.snapshot())).not.toContain('Running bash')
    expect(session.snapshot().activity).toBe('正在处理附件…')
    failTool()
    await Promise.resolve()
    await Promise.resolve()
    expect(JSON.stringify(session.snapshot())).not.toContain('"summary":"bash"')
  })

  it('retries the last bounded instruction after a stable run error', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Try this')
    await Promise.resolve()
    harness.callbacks().onError('provider_unavailable')
    session.retry()
    await Promise.resolve()

    expect(session.snapshot().status).toBe('working')
    expect(
      session
        .snapshot()
        .timeline.filter((event) => event.kind === 'user')
        .map((event) => (event.kind === 'user' ? event.text : '')),
    ).toEqual(['Try this', 'Try this'])
  })

  it.each([
    ['authenticationLost', 'resolve'],
    ['newTask', 'reject'],
    ['logout', 'resolve'],
  ] as const)(
    'does not repopulate presentation after %s races an in-flight confirmation that later %s',
    async (reset, outcome) => {
      const harness = transportHarness()
      const proposals = proposalsHarness()
      proposals.setPending()
      let settle!: () => void
      proposals.controller.confirm.mockImplementation(
        () =>
          new Promise<void>((resolve, reject) => {
            settle = () => (outcome === 'resolve' ? resolve() : reject(new Error('proposal_stale')))
          }),
      )
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposals.controller,
      })

      const confirmation = session.confirm('p1')
      expect(session.snapshot().applying).toBe(true)
      session[reset]()
      expect(session.snapshot()).toMatchObject({
        applying: false,
        status: 'idle',
        activity: '',
        timeline: [],
        error: undefined,
      })

      settle()
      await confirmation
      expect(session.snapshot()).toMatchObject({
        applying: false,
        status: 'idle',
        activity: '',
        timeline: [],
        error: undefined,
      })
    },
  )

  it('maps arbitrary transport failures to a stable code, safe copy, and retry policy', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Fail safely')
    await Promise.resolve()
    harness.callbacks().onError('/Users/alice/private token=secret')

    expect(session.snapshot()).toMatchObject({
      error: 'agent_run_failed',
      errorMessage: 'The Agent could not complete this request. Try again.',
      retryable: true,
    })
    expect(JSON.stringify(session.snapshot())).not.toContain('alice')
    expect(JSON.stringify(session.snapshot())).not.toContain('secret')
  })

  it('does not retry a known non-retryable authentication failure', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })
    session.send('Protected request')
    await Promise.resolve()
    harness.callbacks().onError('auth_required')
    session.retry()
    expect(session.snapshot()).toMatchObject({ error: 'auth_required', retryable: false })
    expect(session.snapshot().timeline.filter((event) => event.kind === 'user')).toHaveLength(1)
  })

  it('clears presentation state atomically for new task and logout', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    session.send('Old task')
    await Promise.resolve()
    harness.callbacks().onDelta('Old answer')
    harness.callbacks().onDone()
    session.newTask()
    expect(session.snapshot()).toMatchObject({ timeline: [], status: 'idle', error: undefined })
    expect(proposals.controller.logout).toHaveBeenCalledOnce()

    session.send('Another task')
    await Promise.resolve()
    session.logout()
    expect(harness.cancel).toHaveBeenCalledTimes(2)
    expect(session.snapshot().timeline).toEqual([])
  })
  it('exposes generic structured proposal fields without legacy coercion', () => {
    const harness = transportHarness()
    const proposal: StructuredProposal = Object.freeze({
      id: 'structured',
      operation: 'edit_slide_xml',
      toolName: 'edit_slide_xml',
      title: 'Update slide XML',
      preview: { nodes: 2 },
      impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
      fingerprint: 'fp',
      before: '<old/>',
      after: '<new/>',
      code: 'context.sync()',
    })
    const controller = {
      pending: () => proposal,
      propose: vi.fn(),
      confirm: vi.fn(),
      reject: vi.fn(),
      newTurn: vi.fn(),
      logout: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: controller,
    })
    expect(session.snapshot().proposal).toEqual(proposal)
  })

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
    expect(session.snapshot()).toMatchObject({
      error: 'proposal_stale',
      errorMessage: '文档内容已发生变化，刚才的修改未应用。',
      retryable: true,
    })
    expect(session.snapshot().proposal).toBeUndefined()
  })

  it.each([
    ['office_verify_failed', 'The approved change could not be verified.'],
    ['office_recovery_failed', 'The document could not be restored after the failed change.'],
  ])('preserves the terminal confirmation code %s with safe copy', async (code, message) => {
    const proposals = proposalsHarness()
    proposals.controller.confirm.mockRejectedValue(new Error(code))
    proposals.setPending()
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposals.controller,
    })

    await session.confirm('p1')

    expect(session.snapshot()).toMatchObject({
      error: code,
      errorMessage: message,
      retryable: false,
    })
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

    expect(session.snapshot().applying).toBe(true)
    expect(proposals.controller.confirm).toHaveBeenCalledOnce()
    expect(proposals.controller.newTurn).not.toHaveBeenCalled()
    expect(proposals.controller.reject).not.toHaveBeenCalled()
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
