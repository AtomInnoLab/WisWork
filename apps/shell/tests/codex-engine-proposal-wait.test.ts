import { afterEach, expect, it, vi } from 'vitest'
import type { ToolExecution } from '@wiswork/agent-core'
import { createProductionCodexBootstrap } from '../src/main/codex-engine'
import type { CodexRuntimeEngineEvent } from '../src/main/codex-runtime'

const mock = vi.hoisted(() => ({
  notify: undefined as any,
  document: undefined as any,
  onStreamActivity: undefined as ((turnId?: string) => void) | undefined,
  revoke: vi.fn(),
  startThread: vi.fn(async () => ({ thread: { id: 'thread' } })),
  startTurn: vi.fn(async () => ({ turn: { id: 'turn' } })),
}))
vi.mock('@wiswork/codex-bridge', async (original) => ({
  ...(await original<any>()),
  startResponsesBridge: async (options: { onStreamActivity?: (turnId?: string) => void }) => {
    mock.onStreamActivity = options.onStreamActivity
    return { baseUrl: '', secret: '', close: async () => {} }
  },
  startDynamicMcpGateway: async () => ({
    url: '',
    secret: '',
    register: (document: any) => {
      mock.document = document
      return () => {}
    },
    beginTurn: () => ({ capability: 'capability' }),
    bindTurn: () => {},
    revokeTurn: mock.revoke,
    close: async () => {},
  }),
  CodexProcessManager: class {
    crashed = new Promise(() => {})
    async stop() {}
    async start() {
      return {
        initialize: async () => {},
        onNotification: (listener: any) => {
          mock.notify = listener
          return () => {}
        },
        startThread: mock.startThread,
        startTurn: mock.startTurn,
        interruptTurn: async () => {},
      }
    }
  },
}))
afterEach(() => {
  vi.useRealTimers()
  vi.clearAllMocks()
  mock.startThread.mockImplementation(async () => ({ thread: { id: 'thread' } }))
  mock.startTurn.mockImplementation(async () => ({ turn: { id: 'turn' } }))
})

async function startSlidesTurn() {
  vi.useFakeTimers()
  const events: CodexRuntimeEngineEvent[] = []
  const engine = await createProductionCodexBootstrap({ fetchWithAuth: vi.fn() }).start({
    executablePath: '',
    onCrash: vi.fn(),
  })
  engine.registerDocument!({
    ownerId: 'owner',
    documentId: 'doc',
    host: 'slides',
    generation: 1,
    session: {
      credentials: {},
      listTools: () => [
        { name: 'ask_clarification', annotations: { readOnlyHint: true } },
        { name: 'read_presentation', annotations: { readOnlyHint: true } },
        { name: 'build_deck', annotations: { readOnlyHint: false } },
      ],
      close: () => {},
    } as any,
    onEvent: (event) => events.push(event),
  })
  let result = 'pending'
  const running = engine
    .startTurn({
      documentId: 'doc',
      host: 'slides',
      generation: 1,
      text: 'make slides',
    })
    .then(
      () => {
        result = 'done'
      },
      (error: Error) => {
        result = error.message
      },
    )
  await vi.advanceTimersByTimeAsync(0)
  return {
    engine,
    events,
    running,
    get result() {
      return result
    },
    startTool(toolName: string, callId = toolName) {
      mock.document.onToolEvent({ type: 'tool-start', callId, toolName })
    },
    completeTool(toolName: string, isError = false, callId = toolName) {
      mock.document.onToolEvent({ type: 'tool-complete', callId, toolName, isError })
    },
    completeNativeTurn() {
      mock.notify({
        method: 'turn/completed',
        params: { threadId: 'thread', turn: { id: 'turn', status: 'completed' } },
      })
    },
    emitText(text = 'Continuing after the questionnaire') {
      mock.notify({
        method: 'item/agentMessage/delta',
        params: { threadId: 'thread', turnId: 'turn', delta: text },
      })
    },
    propose() {
      let resolve!: (execution: ToolExecution) => void
      let reject!: (error: Error) => void
      const settled = new Promise<ToolExecution>((onResolve, onReject) => {
        resolve = onResolve
        reject = onReject
      })
      mock.document.onProposal({
        proposalId: 'proposal',
        call: { id: 'build_deck', name: 'build_deck', input: {} },
        expiresAt: Date.now() + 300_000,
        summary: {},
        settled,
      })
      return { resolve, reject }
    },
  }
}

it('keeps buffered stream activity alive beyond the turn idle deadline without emitting events', async () => {
  const turn = await startSlidesTurn()
  try {
    turn.emitText('Preparing the design')
    for (let index = 0; index < 4; index++) {
      await vi.advanceTimersByTimeAsync(25_000)
      mock.onStreamActivity?.('turn')
      expect(turn.result).toBe('pending')
    }
    expect(turn.events).toEqual([{ type: 'text', text: 'Preparing the design' }])
    expect(mock.revoke).not.toHaveBeenCalled()
    turn.completeNativeTurn()
    await turn.running
    expect(turn.result).toBe('done')
    expect(turn.events.at(-1)).toEqual({ type: 'terminal', status: 'completed' })
    expect(mock.startTurn).toHaveBeenCalledTimes(1)
  } finally {
    await turn.engine.close()
  }
})

it.each(['absent', 'missing', 'wrong'])(
  'still expires the active turn when stream activity has an %s turn ID',
  async (activity) => {
    const turn = await startSlidesTurn()
    try {
      await vi.advanceTimersByTimeAsync(45_000)
      if (activity !== 'absent')
        mock.onStreamActivity?.(activity === 'missing' ? undefined : 'another-turn')
      await vi.advanceTimersByTimeAsync(15_001)
      await turn.running
      expect(turn.result).toBe('enhanced_turn_timeout')
      expect(turn.events).toEqual([{ type: 'terminal', status: 'failed' }])
    } finally {
      await turn.engine.close()
    }
  },
)

it('expires after matching buffered stream activity stops', async () => {
  const turn = await startSlidesTurn()
  try {
    await vi.advanceTimersByTimeAsync(45_000)
    mock.onStreamActivity?.('turn')
    await vi.advanceTimersByTimeAsync(45_000)
    expect(turn.result).toBe('pending')
    await vi.advanceTimersByTimeAsync(15_001)
    await turn.running
    expect(turn.result).toBe('enhanced_turn_timeout')
    expect(turn.events).toEqual([{ type: 'terminal', status: 'failed' }])
  } finally {
    await turn.engine.close()
  }
})

it('does not let activity from a completed turn extend its replacement turn', async () => {
  const turn = await startSlidesTurn()
  try {
    turn.completeNativeTurn()
    await turn.running
    mock.startTurn.mockResolvedValueOnce({ turn: { id: 'replacement-turn' } })
    let nextResult = 'pending'
    const nextRunning = turn.engine
      .startTurn({ documentId: 'doc', host: 'slides', generation: 1, text: 'continue' })
      .then(
        () => {
          nextResult = 'done'
        },
        (error: Error) => {
          nextResult = error.message
        },
      )
    await vi.advanceTimersByTimeAsync(45_000)
    mock.onStreamActivity?.('turn')
    await vi.advanceTimersByTimeAsync(15_001)
    await nextRunning
    expect(nextResult).toBe('enhanced_turn_timeout')
    expect(turn.events).toEqual([
      { type: 'terminal', status: 'completed' },
      { type: 'terminal', status: 'failed' },
    ])
  } finally {
    await turn.engine.close()
  }
})

it.each(['cancelled', 'closed', 'completed'])(
  'does not revive a %s turn with buffered stream activity',
  async (status) => {
    const turn = await startSlidesTurn()
    try {
      if (status === 'cancelled') await turn.engine.cancelTurn('doc')
      else if (status === 'closed') await turn.engine.closeDocument!('doc')
      else turn.completeNativeTurn()
      await turn.running
      const terminalEvents = [...turn.events]
      const revocations = mock.revoke.mock.calls.length
      const result = turn.result
      for (let index = 0; index < 4; index++) {
        await vi.advanceTimersByTimeAsync(25_000)
        mock.onStreamActivity?.('turn')
      }
      expect(turn.result).toBe(result)
      expect(turn.events).toEqual(terminalEvents)
      expect(mock.revoke).toHaveBeenCalledTimes(revocations)
      expect(mock.startTurn).toHaveBeenCalledTimes(1)
      expect(vi.getTimerCount()).toBe(0)
    } finally {
      await turn.engine.close()
    }
  },
)

it.each(['applied', 'tool_error', 'rejected'])(
  'publishes proposal tool completion before the deferred terminal; outcome=%s',
  async (outcome) => {
    const turn = await startSlidesTurn()
    try {
      turn.startTool('build_deck')
      const proposal = turn.propose()
      turn.completeNativeTurn()
      if (outcome === 'rejected') proposal.reject(new Error('private execution failure'))
      else
        proposal.resolve({
          output: outcome,
          summary: outcome,
          isError: outcome === 'tool_error',
          mutated: outcome === 'applied',
        })
      // The gateway resumes after these proposal promise listeners and then
      // publishes tool-complete. Native completion must stay deferred until then.
      await vi.advanceTimersByTimeAsync(0)
      expect(turn.result).toBe('pending')
      expect(mock.revoke).not.toHaveBeenCalled()
      expect(turn.events.some((event) => event.type === 'terminal')).toBe(false)
      turn.completeTool('build_deck', outcome !== 'applied')
      await turn.running
      expect(turn.result).toBe(outcome === 'rejected' ? 'enhanced_proposal_failed' : 'done')
      expect(turn.events.slice(-2)).toEqual([
        {
          type: 'tool-complete',
          callId: 'build_deck',
          toolName: 'build_deck',
          isError: outcome !== 'applied',
        },
        { type: 'terminal', status: outcome === 'rejected' ? 'failed' : 'completed' },
      ])
    } finally {
      await turn.engine.close()
    }
  },
)

it.each(
  [
    { output: 'applied', isError: false, mutated: true },
    { output: 'partial_build', isError: true, mutated: true },
    { output: 'tool_failed', isError: true, mutated: false },
  ].flatMap((execution) => ['before', 'after'].map((completion) => ({ execution, completion }))),
)(
  'recognizes the native questionnaire continuation when $execution.output finishes $completion completion',
  async ({ execution, completion }) => {
    const turn = await startSlidesTurn()
    turn.startTool('ask_clarification')
    turn.completeTool('ask_clarification')
    turn.startTool('build_deck')
    const proposal = turn.propose()
    if (completion === 'after') {
      turn.completeNativeTurn()
      await vi.advanceTimersByTimeAsync(120_000)
      expect(turn.result).toBe('pending')
      expect(mock.revoke).not.toHaveBeenCalled()
    }
    proposal.resolve({ ...execution, summary: execution.output })
    // Proposal promise listeners run before the gateway emits tool-complete.
    await vi.advanceTimersByTimeAsync(0)
    turn.completeTool('build_deck', execution.isError)
    if (completion === 'before') turn.completeNativeTurn()
    await turn.running
    expect(turn.result).toBe('done')
    expect(turn.events.filter((event) => event.type === 'terminal')).toEqual([
      { type: 'terminal', status: 'completed' },
    ])
    expect(mock.startTurn).toHaveBeenCalledExactlyOnceWith(
      'thread',
      '<wiswork_turn_capability>capability</wiswork_turn_capability>\n\nmake slides',
    )
    await turn.engine.close()
  },
)

it.each([false, true])(
  'recognizes a post-answer read as continuation; isError=%s',
  async (isError) => {
    const turn = await startSlidesTurn()
    turn.startTool('ask_clarification')
    turn.completeTool('ask_clarification')
    turn.startTool('read_presentation')
    turn.completeTool('read_presentation', isError)
    turn.completeNativeTurn()
    await turn.running
    expect(turn.result).toBe('done')
    expect(mock.startTurn).toHaveBeenCalledTimes(1)
    await turn.engine.close()
  },
)

it.each(['before', 'after'])(
  'recovers a failed questionnaire after a successful retry; proposal finishes %s native completion',
  async (completion) => {
    const turn = await startSlidesTurn()
    turn.startTool('ask_clarification', 'malformed-survey')
    turn.completeTool('ask_clarification', true, 'malformed-survey')
    turn.startTool('ask_clarification', 'corrected-survey')
    turn.completeTool('ask_clarification', false, 'corrected-survey')
    turn.startTool('build_deck')
    const proposal = turn.propose()
    if (completion === 'after') {
      turn.completeNativeTurn()
      await vi.advanceTimersByTimeAsync(120_000)
      expect(turn.result).toBe('pending')
      expect(mock.revoke).not.toHaveBeenCalled()
    }
    proposal.resolve({
      output: 'partial_build',
      summary: 'Partial build',
      isError: true,
      mutated: true,
    })
    await vi.advanceTimersByTimeAsync(0)
    turn.completeTool('build_deck', true)
    if (completion === 'before') turn.completeNativeTurn()
    await turn.running
    expect(turn.result).toBe('done')
    expect(turn.events.filter((event) => event.type === 'terminal')).toEqual([
      { type: 'terminal', status: 'completed' },
    ])
    expect(mock.startTurn).toHaveBeenCalledTimes(1)
    await turn.engine.close()
  },
)

it('treats a rejected questionnaire as a recoverable tool error', async () => {
  const turn = await startSlidesTurn()
  turn.startTool('ask_clarification')
  turn.completeTool('ask_clarification', true)
  turn.startTool('read_presentation')
  turn.completeTool('read_presentation')
  turn.completeNativeTurn()
  await turn.running
  expect(turn.result).toBe('done')
  expect(mock.startTurn).toHaveBeenCalledTimes(1)
  await turn.engine.close()
})

it('recognizes post-answer assistant text as native questionnaire continuation', async () => {
  const turn = await startSlidesTurn()
  turn.startTool('ask_clarification')
  turn.completeTool('ask_clarification')
  turn.emitText()
  turn.completeNativeTurn()
  await turn.running
  expect(turn.result).toBe('done')
  expect(turn.events).toContainEqual({ type: 'text', text: 'Continuing after the questionnaire' })
  await turn.engine.close()
})

it('releases an orphaned questionnaire when the model continues with final text', async () => {
  const turn = await startSlidesTurn()
  turn.startTool('ask_clarification')
  turn.emitText('Continuing with professional defaults')
  turn.completeNativeTurn()
  await turn.running
  expect(turn.result).toBe('done')
  expect(turn.events.filter((event) => event.type === 'terminal')).toEqual([
    { type: 'terminal', status: 'completed' },
  ])
  await turn.engine.close()
})

it.each(['completed', 'cancelled'])(
  'waits for a questionnaire retry and preserves %s without continuation',
  async (status) => {
    const turn = await startSlidesTurn()
    turn.startTool('ask_clarification', 'malformed-survey')
    turn.completeTool('ask_clarification', true, 'malformed-survey')
    turn.startTool('ask_clarification', 'corrected-survey')
    turn.completeNativeTurn()
    await vi.advanceTimersByTimeAsync(120_000)
    expect(turn.result).toBe('pending')
    expect(mock.revoke).not.toHaveBeenCalled()
    if (status === 'cancelled') await turn.engine.cancelTurn('doc')
    turn.completeTool('ask_clarification', false, 'corrected-survey')
    await turn.running
    expect(turn.result).toBe(status === 'cancelled' ? 'done' : 'enhanced_questionnaire_incomplete')
    expect(turn.events.filter((event) => event.type === 'terminal')).toEqual([
      { type: 'terminal', status: status === 'cancelled' ? 'cancelled' : 'failed' },
    ])
    expect(mock.startTurn).toHaveBeenCalledTimes(1)
    await turn.engine.close()
  },
)

it.each(['none', 'unknown_tool', 'before_answer', 'write_before_answer'])(
  'still detects an uncontinued questionnaire with activity=%s',
  async (activity) => {
    const turn = await startSlidesTurn()
    turn.startTool('ask_clarification')
    if (activity === 'before_answer') turn.startTool('read_presentation')
    if (activity === 'write_before_answer') turn.startTool('build_deck')
    turn.completeTool('ask_clarification')
    if (activity === 'unknown_tool') {
      turn.startTool('unknown_tool')
      turn.completeTool('unknown_tool')
    }
    if (activity === 'before_answer') turn.completeTool('read_presentation')
    if (activity === 'write_before_answer') turn.completeTool('build_deck')
    turn.completeNativeTurn()
    await turn.running
    expect(turn.result).toBe('enhanced_questionnaire_incomplete')
    expect(mock.startTurn).toHaveBeenCalledTimes(1)
    await turn.engine.close()
  },
)

it.each([
  { output: 'applied', expected: 'enhanced_questionnaire_incomplete', status: 'failed' },
  { output: 'tool_failed', expected: 'enhanced_questionnaire_incomplete', status: 'failed' },
  { output: 'mutation_expired', expected: 'enhanced_questionnaire_incomplete', status: 'failed' },
  { output: 'mutation_cancelled', expected: 'enhanced_questionnaire_incomplete', status: 'failed' },
  { output: 'rejected', expected: 'enhanced_proposal_failed', status: 'failed' },
])(
  'waits for the pending questionnaire after a deferred proposal $output',
  async ({ output, expected, status }) => {
    const turn = await startSlidesTurn()
    turn.startTool('ask_clarification')
    turn.startTool('build_deck')
    const proposal = turn.propose()
    turn.completeNativeTurn()
    if (output === 'rejected') proposal.reject(new Error('private execution failure'))
    else
      proposal.resolve({
        output,
        summary: output,
        isError: output !== 'applied',
        mutated: output === 'applied',
      })
    await vi.advanceTimersByTimeAsync(120_000)
    expect(turn.result).toBe('pending')
    expect(mock.revoke).not.toHaveBeenCalled()
    expect(turn.events.some((event) => event.type === 'terminal')).toBe(false)
    turn.completeTool('build_deck', output !== 'applied')
    turn.completeTool('ask_clarification')
    await turn.running
    expect(turn.result).toBe(expected)
    expect(turn.events.at(-1)).toMatchObject({ type: 'terminal', status })
    expect(mock.startTurn).toHaveBeenCalledTimes(1)
    await turn.engine.close()
  },
)

it.each(['before', 'after'])(
  'waits for a proposal and preserves completion when the questionnaire fails %s native completion',
  async (completion) => {
    const turn = await startSlidesTurn()
    turn.startTool('ask_clarification')
    turn.startTool('build_deck')
    const proposal = turn.propose()
    if (completion === 'before') turn.completeTool('ask_clarification', true)
    turn.completeNativeTurn()
    if (completion === 'after') turn.completeTool('ask_clarification', true)
    await vi.advanceTimersByTimeAsync(120_000)
    expect(turn.result).toBe('pending')
    expect(mock.revoke).not.toHaveBeenCalled()
    proposal.resolve({
      output: 'tool_failed',
      summary: 'Tool failed',
      isError: true,
      mutated: false,
    })
    await vi.advanceTimersByTimeAsync(0)
    turn.completeTool('build_deck', true)
    await turn.running
    expect(turn.result).toBe('done')
    expect(mock.startTurn).toHaveBeenCalledTimes(1)
    await turn.engine.close()
  },
)

it('diagnoses thread and turn start boundaries without retaining request content', async () => {
  const diagnostics: string[] = []
  mock.startThread.mockRejectedValueOnce(new Error('private thread failure'))
  const first = await createProductionCodexBootstrap({
    fetchWithAuth: vi.fn(),
    diagnostics: (code) => diagnostics.push(code),
  }).start({ executablePath: '', onCrash: vi.fn() })
  first.registerDocument!({
    ownerId: 'owner',
    documentId: 'doc',
    host: 'slides',
    generation: 1,
    session: { credentials: {}, listTools: () => [], close: () => {} } as any,
  })
  await expect(
    first.startTurn({ documentId: 'doc', host: 'slides', generation: 1, text: 'PRIVATE PROMPT' }),
  ).rejects.toThrow()
  expect(diagnostics).toEqual(
    expect.arrayContaining(['enhanced_thread_starting', 'enhanced_thread_start_failed']),
  )
  expect(diagnostics.join(',')).not.toContain('PRIVATE')
  await first.close()

  diagnostics.length = 0
  mock.startTurn.mockRejectedValueOnce(new Error('private turn failure'))
  const second = await createProductionCodexBootstrap({
    fetchWithAuth: vi.fn(),
    diagnostics: (code) => diagnostics.push(code),
  }).start({ executablePath: '', onCrash: vi.fn() })
  second.registerDocument!({
    ownerId: 'owner',
    documentId: 'doc',
    host: 'slides',
    generation: 1,
    session: { credentials: {}, listTools: () => [], close: () => {} } as any,
  })
  await expect(
    second.startTurn({ documentId: 'doc', host: 'slides', generation: 1, text: 'PRIVATE PROMPT' }),
  ).rejects.toThrow()
  expect(diagnostics).toEqual(
    expect.arrayContaining([
      'enhanced_thread_starting',
      'enhanced_thread_started',
      'enhanced_turn_starting',
      'enhanced_turn_start_failed',
    ]),
  )
  expect(diagnostics.join(',')).not.toContain('PRIVATE')
  await second.close()
})

it.each(
  ['slides', 'office-powerpoint'].flatMap((host) =>
    ['answered', 'cancelled', 'failed'].map((outcome) => ({ host, outcome })),
  ),
)(
  'keeps a $host questionnaire inside one host run; outcome=$outcome',
  async ({ host, outcome }) => {
    const cancelled = outcome === 'cancelled'
    const engine = await createProductionCodexBootstrap({ fetchWithAuth: vi.fn() }).start({
      executablePath: '',
      onCrash: vi.fn(),
    })
    engine.registerDocument!({
      ownerId: 'owner',
      documentId: 'doc',
      host: host as 'slides' | 'office-powerpoint',
      generation: 1,
      session: {
        credentials: {},
        listTools: () => [{ name: 'build_deck', annotations: { readOnlyHint: false } }],
        close: () => {},
      } as any,
    })
    let done = false
    const running = engine
      .startTurn({
        documentId: 'doc',
        host: host as 'slides' | 'office-powerpoint',
        generation: 1,
        text: 'make slides',
      })
      .then(
        () => {
          done = true
        },
        (error: Error) => error,
      )
    await new Promise((r) => setTimeout(r, 0))
    mock.document.onToolEvent({
      type: 'tool-start',
      callId: 'survey',
      toolName: 'ask_clarification',
    })
    if (outcome !== 'answered')
      mock.notify({
        method: 'turn/completed',
        params: { threadId: 'thread', turn: { id: 'turn', status: 'completed' } },
      })
    await new Promise((r) => setTimeout(r, 0))
    expect(done).toBe(false)
    if (cancelled) await engine.cancelTurn('doc')
    mock.document.onToolEvent({
      type: 'tool-complete',
      callId: 'survey',
      toolName: 'ask_clarification',
      isError: outcome === 'failed',
    })
    await new Promise((r) => setTimeout(r, 0))
    if (outcome === 'failed') {
      expect(await running).toBeUndefined()
      expect(done).toBe(true)
      await engine.close()
      return
    }
    if (cancelled) {
      await running
      expect(mock.startTurn).toHaveBeenCalledTimes(1)
      await engine.close()
      return
    }
    expect(mock.startTurn).toHaveBeenCalledTimes(1)
    expect(mock.revoke).not.toHaveBeenCalled()
    mock.document.onToolEvent({
      type: 'tool-start',
      callId: 'build',
      toolName: 'build_deck',
    })
    mock.document.onToolEvent({
      type: 'tool-complete',
      callId: 'build',
      toolName: 'build_deck',
      isError: false,
    })
    mock.notify({
      method: 'turn/completed',
      params: { threadId: 'thread', turn: { id: 'turn', status: 'completed' } },
    })
    await running
    expect(done).toBe(true)
    await engine.close()
  },
)

it.each(['applied', 'tool_failed', 'mutation_expired', 'mutation_cancelled', 'cancel', 'close'])(
  'waits past idle timeout after model completion until proposal %s',
  async (output) => {
    vi.useFakeTimers()
    const engine = await createProductionCodexBootstrap({ fetchWithAuth: vi.fn() }).start({
      executablePath: '',
      onCrash: vi.fn(),
    })
    const events: any[] = []
    engine.registerDocument!({
      ownerId: 'owner',
      documentId: 'doc',
      host: 'docs',
      generation: 1,
      session: { credentials: {}, listTools: () => [], close: () => {} } as any,
      onEvent: (event) => events.push(event),
    })
    let settle!: (value: any) => void
    const settled = new Promise<any>((resolve) => {
      settle = resolve
    })
    let result: unknown = 'pending'
    const running = engine
      .startTurn({ documentId: 'doc', host: 'docs', generation: 1, text: 'edit' })
      .then(
        () => {
          result = 'done'
        },
        (error) => {
          result = error.message
        },
      )
    await vi.advanceTimersByTimeAsync(0)
    mock.document.onProposal({
      proposalId: 'proposal',
      call: {},
      expiresAt: Date.now() + 300_000,
      summary: {},
      settled,
    })
    await vi.advanceTimersByTimeAsync(120_000)
    expect(result).toBe('pending')
    mock.notify({
      method: 'turn/completed',
      params: { threadId: 'thread', turn: { id: 'turn', status: 'completed' } },
    })
    await vi.advanceTimersByTimeAsync(120_000)
    expect(result).toBe('pending')
    expect(events.some((event) => event.type === 'terminal')).toBe(false)
    if (output === 'cancel' || output === 'close') {
      if (output === 'cancel') await engine.cancelTurn('doc')
      else await engine.closeDocument!('doc')
      await running
      expect(result).toBe(output === 'cancel' ? 'done' : 'document_session_unavailable')
      expect(mock.revoke).toHaveBeenCalled()
      await engine.close()
      return
    }
    settle({
      output,
      summary: output,
      isError: output !== 'applied',
      mutated: output === 'applied',
    })
    await vi.advanceTimersByTimeAsync(0)
    mock.document.onToolEvent({
      type: 'tool-complete',
      callId: 'proposal-call',
      toolName: 'replace_blocks',
      isError: output !== 'applied',
    })
    await running
    expect(result).toBe('done')
    expect(events.at(-1)).toMatchObject({
      type: 'terminal',
      status: 'completed',
    })
    await engine.close()
  },
)
