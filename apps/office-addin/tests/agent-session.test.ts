import {
  suspendToolExecution,
  type AgentStreamCallbacks,
  type AgentTransport,
  type ToolExecution,
} from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { bindAuthLoss, createOfficeAgentSession } from '../src/agent/use-office-agent.js'
import type { ProposalDecision, StructuredProposal } from '../src/agent/proposal-controller.js'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'

function transportHarness() {
  let callbacks: AgentStreamCallbacks | undefined
  const cancel = vi.fn()
  const stream = vi.fn((_request: unknown, next: AgentStreamCallbacks) => {
    callbacks = next
    return { cancel }
  })
  const transport: AgentTransport = {
    stream,
  }
  return { transport, cancel, stream, callbacks: () => callbacks! }
}

function proposalsHarness() {
  let pending:
    | { id: string; operation: 'replace'; before: string; value: string; fingerprint: string }
    | undefined
  const listeners = new Set<() => void>()
  let settleDecision: ((value: ProposalDecision) => void) | undefined
  let decision = Promise.resolve<ProposalDecision>({ status: 'cancelled' })
  const clear = (value: ProposalDecision) => {
    pending = undefined
    settleDecision?.(value)
    settleDecision = undefined
    listeners.forEach((listener) => listener())
  }
  const controller = {
    pending: () => pending,
    subscribe(listener: () => void) {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    waitForDecision: () => decision,
    propose: vi.fn(),
    confirm: vi.fn(async () => {
      clear({ status: 'confirmed' })
    }),
    reject: vi.fn(() => {
      clear({ status: 'rejected' })
    }),
    newTurn: vi.fn(() => {
      clear({ status: 'cancelled' })
    }),
    logout: vi.fn(() => {
      clear({ status: 'cancelled' })
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
      decision = new Promise((resolve) => {
        settleDecision = resolve
      })
      listeners.forEach((listener) => listener())
    },
    clearPending() {
      pending = undefined
      listeners.forEach((listener) => listener())
    },
  }
}

describe('Office agent session', () => {
  it('opens after legacy prompt scrub failure but disables recovery and warns', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        scrubFailed: true,
        validateDocument: vi.fn(async () => true),
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(false)
    expect(session.snapshot().timeline[0]).toMatchObject({
      kind: 'system',
      text: expect.stringContaining('请求原文仍保留在本 PPTX'),
    })
    await session.resumeInterrupted?.()
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it('resumes a pre-tool run only on explicit action after document validation', async () => {
    const harness = transportHarness()
    const validateDocument = vi.fn(async () => true)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: { instruction: 'Create deck', phase: 'running' },
        validateDocument,
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(true)
    expect(harness.stream).not.toHaveBeenCalled()
    await session.resumeInterrupted?.()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    expect(validateDocument).toHaveBeenCalledOnce()
  })

  it('never streams a recovered request when the deck switches after validation', async () => {
    const harness = transportHarness()
    let activeDocument = 'original'
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: { instruction: 'Create deck', phase: 'running' },
        validateDocument: async () => {
          activeDocument = 'copy'
          return true
        },
        begin: async () => {
          if (activeDocument !== 'original') throw new Error('presentation_document_changed')
        },
        finish: vi.fn(async () => undefined),
      },
    })
    await session.resumeInterrupted?.()
    await vi.waitFor(() => expect(session.snapshot().status).toBe('error'))
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it('does not resume a run after a tool boundary', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: { instruction: 'Create deck', phase: 'tool_pending', toolName: 'write_page' },
        validateDocument: vi.fn(async () => true),
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(false)
    await session.resumeInterrupted?.()
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it.each([
    ['complete', '导入回执：2 页完成'],
    ['partial', '导入回执：1/2 页完成'],
    ['uncertain', '下一页结果不确定'],
  ] as const)(
    'shows the matched %s import receipt without replaying the tool',
    (status, detail) => {
      const harness = transportHarness()
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposalsHarness().controller,
        runCheckpoint: {
          interrupted: true,
          recovery: {
            instruction: '',
            phase: 'tool_pending',
            toolName: 'import_presentation_production',
            toolCallId: 'call-1',
            restartSafe: false,
            importReceipt: { state: status, completed: status === 'complete' ? 2 : 1, total: 2 },
          },
          begin: vi.fn(async () => undefined),
          finish: vi.fn(async () => undefined),
        },
      })
      expect(session.snapshot().recoveryAvailable).toBe(false)
      expect(session.snapshot().timeline[0]).toMatchObject({
        kind: 'system',
        text: expect.stringContaining(detail),
      })
      if (status === 'uncertain')
        expect(session.snapshot().timeline[0]).toMatchObject({
          text: expect.stringContaining('reconcile_presentation_production_import'),
        })
      expect(harness.stream).not.toHaveBeenCalled()
    },
  )

  it('shows an interrupted change savepoint without offering write replay', () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: {
          instruction: '',
          phase: 'tool_pending',
          toolName: 'edit_existing_presentation_text',
          toolCallId: 'call-1',
          restartSafe: false,
          changeReceipt: { total: 2, unresolved: 1 },
        },
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(false)
    expect(session.snapshot().timeline[0]).toMatchObject({
      kind: 'system',
      text: expect.stringContaining('对应修改历史 2 项，其中 1 项未结算'),
    })
    expect(harness.stream).not.toHaveBeenCalled()
  })

  it('restarts an interrupted read-only run only on explicit action after document validation', async () => {
    const harness = transportHarness()
    const validateDocument = vi.fn(async () => true)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: true,
        recovery: {
          instruction: 'Inspect this presentation',
          phase: 'tool_pending',
          toolName: 'read_presentation_plan',
          restartSafe: true,
        },
        validateDocument,
        begin: vi.fn(async () => undefined),
        finish: vi.fn(async () => undefined),
      },
    })
    expect(session.snapshot().recoveryAvailable).toBe(true)
    expect(harness.stream).not.toHaveBeenCalled()
    await session.resumeInterrupted?.()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    expect(validateDocument).toHaveBeenCalledOnce()
  })

  it.each(['newTask', 'logout'] as const)(
    'does not restart an interrupted request after %s during document validation',
    async (action) => {
      const harness = transportHarness()
      let releaseValidation!: (valid: boolean) => void
      const begin = vi.fn(async () => undefined)
      const validateDocument = vi.fn(
        () =>
          new Promise<boolean>((resolve) => {
            releaseValidation = resolve
          }),
      )
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposalsHarness().controller,
        runCheckpoint: {
          interrupted: true,
          recovery: {
            instruction: 'Inspect this presentation',
            phase: 'tool_pending',
            toolName: 'read_presentation_plan',
            restartSafe: true,
          },
          validateDocument,
          begin,
          finish: vi.fn(async () => undefined),
        },
      })
      const pending = session.resumeInterrupted?.()
      expect(validateDocument).toHaveBeenCalledOnce()
      session[action]()
      releaseValidation(true)
      await pending
      expect(begin).not.toHaveBeenCalled()
      expect(harness.stream).not.toHaveBeenCalled()
    },
  )

  it('saves a checkpoint before starting and clears it after completion', async () => {
    const harness = transportHarness()
    const begin = vi.fn(async (_runId: string) => undefined)
    const finish = vi.fn(async (_runId: string) => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: { interrupted: true, begin, finish },
    })
    expect(session.snapshot().timeline[0]).toMatchObject({ kind: 'system' })
    session.send('continue')
    expect(harness.stream).not.toHaveBeenCalled()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    expect(begin).toHaveBeenCalledOnce()
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(finish).toHaveBeenCalledOnce())
    expect(finish.mock.calls[0]?.[0]).toBe(begin.mock.calls[0]?.[0])
  })

  it('records a content-free run completion after a successful stream', async () => {
    const harness = transportHarness()
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      diagnostics,
    })
    session.send('private presentation brief')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(diagnostics.record).toHaveBeenCalledWith({
        phase: 'run',
        errorCode: 'agent_run_completed',
        durationMs: expect.any(Number),
      }),
    )
    expect(diagnostics.setTool).toHaveBeenCalledWith('agent_run')
    expect(JSON.stringify(diagnostics.record.mock.calls)).not.toContain(
      'private presentation brief',
    )
  })

  it('waits for an ordinary tool completion checkpoint before the next request', async () => {
    const harness = transportHarness()
    let savePending!: () => void
    let saveCompleted!: () => void
    const executeTool = vi.fn(async () => ({ output: 'read', summary: 'Read document' }))
    const begin = vi.fn(async (_runId: string) => undefined)
    const tool = vi.fn(
      (_runId: string, phase: 'tool_pending' | 'tool_completed') =>
        new Promise<void>((resolve) => {
          if (phase === 'tool_pending') savePending = resolve
          else saveCompleted = resolve
        }),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'read_document', description: 'read', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals: proposalsHarness().controller,
      runCheckpoint: { interrupted: false, begin, tool, finish: vi.fn(async () => undefined) },
    })

    session.send('read')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'read-1', name: 'read_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_pending',
        'read_document',
        false,
        'read-1',
      ),
    )
    expect(executeTool).not.toHaveBeenCalled()

    savePending()
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_completed',
        'read_document',
        false,
        'read-1',
      ),
    )
    expect(executeTool).toHaveBeenCalledOnce()
    expect(tool.mock.calls[0]?.[0]).toBe(begin.mock.calls[0]?.[0])
    expect(tool.mock.calls[1]?.[0]).toBe(begin.mock.calls[0]?.[0])
    expect(harness.stream).toHaveBeenCalledOnce()

    saveCompleted()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
  })

  it('halts after a completed tool when its final checkpoint cannot be saved', async () => {
    const harness = transportHarness()
    const finish = vi.fn(async () => undefined)
    const executeTool = vi.fn(async () => ({
      output: 'changed',
      summary: 'Changed',
      mutated: true,
    }))
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'write', description: 'write', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool: vi.fn(async (_id, phase) => {
          if (phase === 'tool_completed') throw new Error('save failed')
        }),
        finish,
      },
    })
    session.send('change')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'write-1', name: 'write', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable'),
    )
    expect(executeTool).toHaveBeenCalledOnce()
    expect(harness.stream).toHaveBeenCalledOnce()
    expect(finish).not.toHaveBeenCalled()
    expect(session.snapshot().retryable).toBe(false)
  })

  it.each(['newTask', 'logout'] as const)(
    'keeps the pending write checkpoint when %s resets a tool still executing',
    async (action) => {
      const harness = transportHarness()
      let finishWrite!: (result: ToolExecution) => void
      const executeTool = vi.fn(
        () =>
          new Promise<ToolExecution>((resolve) => {
            finishWrite = resolve
          }),
      )
      const begin = vi.fn(async () => undefined)
      const tool = vi.fn(async () => undefined)
      const finish = vi.fn(async () => undefined)
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: {
          id: 'test',
          systemPrompt: 'test',
          tools: [{ name: 'write', description: 'write', inputSchema: { type: 'object' } }],
          executeTool,
        },
        proposals: proposalsHarness().controller,
        runCheckpoint: { interrupted: false, begin, tool, finish },
      })
      session.send('change')
      await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
      harness.callbacks().onToolCall({ id: 'write-1', name: 'write', input: {} })
      harness.callbacks().onDone()
      await vi.waitFor(() => expect(executeTool).toHaveBeenCalledOnce())
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_pending',
        'write',
        false,
        'write-1',
      )

      session[action]()
      expect(finish).not.toHaveBeenCalled()
      finishWrite({ output: 'changed', summary: 'Changed', mutated: true })
      await Promise.resolve()
      expect(tool).not.toHaveBeenCalledWith(
        expect.any(String),
        'tool_completed',
        'write',
        true,
        'write-1',
      )
      expect(finish).not.toHaveBeenCalled()
    },
  )

  it('does not execute an old tool after its pending checkpoint outlives the run', async () => {
    const harness = transportHarness()
    let releasePending!: () => void
    const executeTool = vi.fn(async () => ({ output: 'done', summary: 'Done' }))
    const tool = vi.fn((_runId: string, phase: 'tool_pending' | 'tool_completed') =>
      phase === 'tool_pending'
        ? new Promise<void>((resolve) => {
            releasePending = resolve
          })
        : Promise.resolve(),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'read', description: 'read', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool,
        finish: vi.fn(async () => undefined),
      },
    })
    session.send('old')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'old-tool', name: 'read', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_pending',
        'read',
        false,
        'old-tool',
      ),
    )
    session.newTask()
    session.send('new')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    releasePending()
    await Promise.resolve()
    expect(executeTool).not.toHaveBeenCalled()
    expect(tool).toHaveBeenCalledTimes(1)
  })

  it('waits for proposal decision and completion checkpoint before continuing', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    let releaseCompleted!: () => void
    const tool = vi.fn((_runId: string, phase: 'tool_pending' | 'tool_completed') =>
      phase === 'tool_completed'
        ? new Promise<void>((resolve) => {
            releaseCompleted = resolve
          })
        : Promise.resolve(),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'propose', description: 'propose', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => {
          proposals.setPending()
          return { output: 'prepared', summary: 'Prepared' }
        }),
      },
      proposals: proposals.controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool,
        finish: vi.fn(async () => undefined),
      },
    })
    session.send('edit')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'proposal-tool', name: 'propose', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal?.id).toBe('p1'))
    expect(tool).toHaveBeenCalledTimes(1)
    await session.confirm('p1')
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_completed',
        'propose',
        true,
        'proposal-tool',
      ),
    )
    expect(harness.stream).toHaveBeenCalledOnce()
    releaseCompleted()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
  })

  it('halts after approved work if its completion checkpoint fails', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const finish = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'propose', description: 'propose', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => {
          proposals.setPending()
          return { output: 'prepared', summary: 'Prepared' }
        }),
      },
      proposals: proposals.controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool: vi.fn(async (_id, phase) => {
          if (phase === 'tool_completed') throw new Error('save failed')
        }),
        finish,
      },
    })
    session.send('edit')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'proposal-tool', name: 'propose', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal?.id).toBe('p1'))
    await session.confirm('p1')
    await vi.waitFor(() =>
      expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable'),
    )
    expect(harness.stream).toHaveBeenCalledOnce()
    expect(finish).not.toHaveBeenCalled()
  })

  it('waits for a suspended tool result and its completion checkpoint', async () => {
    const harness = transportHarness()
    let releaseResult!: (result: ToolExecution) => void
    let releaseCompleted!: () => void
    const tool = vi.fn((_runId: string, phase: 'tool_pending' | 'tool_completed') =>
      phase === 'tool_completed'
        ? new Promise<void>((resolve) => {
            releaseCompleted = resolve
          })
        : Promise.resolve(),
    )
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'wait', description: 'wait', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () =>
          suspendToolExecution(
            new Promise<ToolExecution>((resolve) => {
              releaseResult = resolve
            }),
          ),
        ),
      },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin: vi.fn(async () => undefined),
        tool,
        finish: vi.fn(async () => undefined),
      },
    })
    session.send('wait')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledOnce())
    harness.callbacks().onToolCall({ id: 'wait-tool', name: 'wait', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(releaseResult).toBeTypeOf('function'))
    expect(tool).toHaveBeenCalledTimes(1)
    releaseResult({ output: 'done', summary: 'Done' })
    await vi.waitFor(() =>
      expect(tool).toHaveBeenCalledWith(
        expect.any(String),
        'tool_completed',
        'wait',
        false,
        'wait-tool',
      ),
    )
    expect(harness.stream).toHaveBeenCalledOnce()
    releaseCompleted()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
  })

  it('does not start a run when its checkpoint cannot be saved', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: {
        interrupted: false,
        begin: async () => {
          throw new Error('save failed')
        },
        finish: vi.fn(async () => undefined),
      },
    })
    session.send('continue')
    await vi.waitFor(() => expect(session.snapshot().status).toBe('error'))
    expect(session.snapshot().error).toBe('presentation_run_checkpoint_unavailable')
    expect(harness.stream).not.toHaveBeenCalled()
  })
  it('passes a longer presentation brief to the bounded local recovery checkpoint', async () => {
    const harness = transportHarness()
    const begin = vi.fn(async () => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: { interrupted: false, begin, finish: vi.fn(async () => undefined) },
    })
    const brief = '页面、来源和样式要求。'.repeat(150)
    expect(brief.length).toBeGreaterThan(1000)
    session.send(brief)
    await vi.waitFor(() => expect(begin).toHaveBeenCalledOnce())
    expect(begin).toHaveBeenCalledWith(expect.any(String), brief)
    session.stop()
  })
  it('does not start a cancelled run after a delayed checkpoint save', async () => {
    const harness = transportHarness()
    let release!: () => void
    const begin = vi.fn(
      () =>
        new Promise<void>((resolve) => {
          release = resolve
        }),
    )
    const finish = vi.fn(async (_runId: string) => undefined)
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      runCheckpoint: { interrupted: false, begin, finish },
    })
    session.send('do work')
    session.stop()
    release()
    await vi.waitFor(() => expect(finish).toHaveBeenCalledOnce())
    expect(harness.stream).not.toHaveBeenCalled()
    expect(session.snapshot().status).toBe('cancelled')
  })
  it('preserves bounded local diagnostics when Relay authentication is lost', () => {
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: transportHarness().transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
      diagnostics,
    })

    session.authenticationLost()

    expect(diagnostics.clear).not.toHaveBeenCalled()
    session.logout()
    expect(diagnostics.clear).toHaveBeenCalledOnce()
  })

  it('correlates a run and records a stable tool failure without retaining output', async () => {
    const harness = transportHarness()
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'write_document', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => ({
          output: 'office_api_unsupported',
          isError: true,
          summary: 'failed',
        })),
      },
      proposals: proposalsHarness().controller,
      diagnostics,
    })
    session.send('write my secret article')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'call', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(diagnostics.record).toHaveBeenCalled())
    expect(diagnostics.startTrace).toHaveBeenCalledOnce()
    expect(diagnostics.setTool).toHaveBeenCalledWith('write_document')
    expect(diagnostics.record).toHaveBeenCalledWith({
      phase: 'tool',
      errorCode: 'office_api_unsupported',
      durationMs: expect.any(Number),
    })
    expect(JSON.stringify(diagnostics.record.mock.calls)).not.toContain('write my secret article')
  })

  it('forwards an in-memory Office diagnostic cause without adding it to model output', async () => {
    const harness = transportHarness()
    const officeError = Object.assign(new Error('secret workbook value'), {
      name: 'RichApi.Error',
      code: 'InvalidArgument',
      debugInfo: { errorLocation: 'Worksheet.getRange' },
    })
    const diagnostics = {
      startTrace: vi.fn(() => 'trace'),
      setTool: vi.fn(),
      record: vi.fn(),
      clear: vi.fn(),
    }
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'excel',
        systemPrompt: 'test',
        tools: [{ name: 'get_cell_ranges', description: 'read', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(async () => ({
          output: 'office_read_failed',
          isError: true,
          summary: 'failed',
          diagnosticError: officeError,
        })),
      },
      proposals: proposalsHarness().controller,
      diagnostics,
    })
    session.send('read cells')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'call', name: 'get_cell_ranges', input: {} })
    harness.callbacks().onDone()

    await vi.waitFor(() =>
      expect(diagnostics.record).toHaveBeenCalledWith({
        phase: 'tool',
        errorCode: 'office_read_failed',
        error: officeError,
        durationMs: expect.any(Number),
      }),
    )
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    expect(JSON.stringify(harness.stream.mock.calls[1])).not.toContain('secret workbook value')
  })
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
        executeTool: vi.fn(async () => {
          proposals.setPending()
          return { output: 'prepared', summary: 'Prepared edit' }
        }),
      },
      proposals: proposals.controller,
    })

    session.send('Edit this')
    await Promise.resolve()
    harness.callbacks().onDelta('I will prepare it.')
    harness.callbacks().onToolCall({ id: 'tool-1', name: 'propose_replace_selection', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal?.id).toBe('p1'))
    await session.confirm('p1')
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
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
      state: 'applied',
    })
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

  it('shows PowerPoint production as a stage activity through ACP updates', async () => {
    const harness = transportHarness()
    let finish!: () => void
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [
          {
            name: 'run_presentation_production',
            description: 'run',
            inputSchema: { type: 'object' },
          },
        ],
        executeTool: vi.fn(
          () =>
            new Promise<ToolExecution>((resolve) => {
              finish = () => resolve({ output: '{}', summary: 'internal', mutated: false })
            }),
        ),
      },
      proposals: proposalsHarness().controller,
    })
    session.send('继续制作')
    await Promise.resolve()
    harness
      .callbacks()
      .onToolCall({ id: 'production-1', name: 'run_presentation_production', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() =>
      expect(session.snapshot().timeline).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ kind: 'tool', summary: '正在处理逐页制作…', state: 'running' }),
        ]),
      ),
    )
    finish()
    await vi.waitFor(() =>
      expect(session.snapshot().timeline).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            kind: 'tool',
            summary: '逐页制作操作已结束',
            state: 'complete',
          }),
        ]),
      ),
    )
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

  it('reports screenshot metadata without sending bulk images into model history', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const pngBase64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aPioAAAAASUVORK5CYII='
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'powerpoint',
        systemPrompt: 'test',
        tools: [{ name: 'edit_slide_text', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: () => {
          const proposal = proposals.propose({
            operation: 'edit_slide_text',
            title: 'Edit',
            preview: {},
            impact: { host: 'powerpoint', targets: ['slide-1'], count: 1 },
            fingerprint: 'v1',
            validate: () => true,
            execute: () => {},
            postWrite: () => ({
              status: 'captured',
              pages: [
                {
                  slideId: 'slide-1',
                  pngBase64,
                  digest: 'f4b555ad4009f54a1a37dc29e7ccf9f8f4cfe22410ba7769061c0328cdb6db67',
                },
              ],
            }),
          })
          return {
            output: JSON.stringify({ proposalId: proposal.id }),
            mutated: false,
            summary: 'Awaiting confirmation',
          }
        },
      },
      proposals,
    })
    session.send('edit')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'write', name: 'edit_slide_text', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())
    await session.confirm(session.snapshot().proposal!.id)
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    const resumed = harness.stream.mock.calls[1]?.[0] as {
      messages: Array<{ results?: Array<{ output: string; content?: unknown }> }>
    }
    const result = resumed.messages.at(-1)?.results?.[0]
    expect(JSON.parse(result!.output)).toMatchObject({
      status: 'applied',
      qaPassed: false,
      visualReview: 'pending',
      postWrite: { status: 'captured', pages: [{ slideId: 'slide-1' }] },
    })
    expect(result!.content).toBeUndefined()
  })

  it('rejects an in-loop proposal immediately and resumes without executing the write', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const execute = vi.fn(async () => undefined)
    const executeTool = vi.fn(() => {
      const proposal = proposals.propose({
        operation: 'write_document',
        title: 'Write document',
        preview: {},
        impact: { host: 'word', targets: ['document'], count: 1 },
        fingerprint: 'v1',
        validate: async () => true,
        execute,
      })
      return {
        output: JSON.stringify({ proposalId: proposal.id }),
        mutated: false,
        summary: 'Awaiting confirmation',
      }
    })
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'word',
        systemPrompt: 'test',
        tools: [{ name: 'write_document', description: 'write', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals,
    })

    session.send('write')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'write-1', name: 'write_document', input: {} })
    harness.callbacks().onToolCall({ id: 'write-2', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())

    session.reject()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))

    expect(execute).not.toHaveBeenCalled()
    expect(executeTool).toHaveBeenCalledOnce()
    expect(session.snapshot().timeline).toEqual(
      expect.arrayContaining([expect.objectContaining({ kind: 'proposal', state: 'rejected' })]),
    )
  })

  it('does not create another Word write proposal after stale validation in the same run', async () => {
    const harness = transportHarness()
    const proposals = createStructuredProposalController()
    const executeTool = vi.fn(() => {
      const proposal = proposals.propose({
        operation: 'write_document',
        title: 'Write document',
        preview: {},
        impact: { host: 'word', targets: ['document'], count: 1 },
        fingerprint: 'v1',
        validate: async () => false,
        execute: vi.fn(),
      })
      return {
        output: JSON.stringify({ proposalId: proposal.id }),
        mutated: false,
        summary: 'Awaiting confirmation',
      }
    })
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'word',
        systemPrompt: 'test',
        tools: [{ name: 'write_document', description: 'write', inputSchema: { type: 'object' } }],
        executeTool,
      },
      proposals,
    })

    session.send('write')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'write-1', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())
    await session.confirm(session.snapshot().proposal!.id)
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))

    harness.callbacks().onToolCall({ id: 'write-2', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(3))
    expect(executeTool).toHaveBeenCalledOnce()
    expect(session.snapshot().proposal).toBeUndefined()

    harness.callbacks().onDelta('The document changed before the edit could be applied.')
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().busy).toBe(false))
    expect(executeTool).toHaveBeenCalledOnce()
    expect(session.snapshot().proposal).toBeUndefined()
  })

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

  it('reports the bounded transport deadline as a retryable request timeout', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Build a complex presentation')
    await Promise.resolve()
    harness.callbacks().onError('transport_timeout')

    expect(session.snapshot()).toMatchObject({
      error: 'request_timeout',
      errorMessage: 'The Agent took too long to respond. Try again.',
      retryable: true,
    })
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

  it('preserves Agent history after stop but clears it for a new task', async () => {
    const harness = transportHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
      proposals: proposalsHarness().controller,
    })

    session.send('Keep this context')
    await Promise.resolve()
    session.stop()
    harness.callbacks().onDone()
    await Promise.resolve()
    session.send('Continue')
    await Promise.resolve()

    expect(harness.stream.mock.calls[1]?.[0]).toMatchObject({
      messages: expect.arrayContaining([
        expect.objectContaining({ role: 'user', text: 'Keep this context' }),
        expect.objectContaining({ role: 'user', text: 'Continue' }),
      ]),
    })

    harness.callbacks().onDone()
    await Promise.resolve()
    session.newTask()
    session.send('Fresh context')
    await Promise.resolve()

    const freshRequest = harness.stream.mock.calls[2]?.[0] as {
      messages: Array<{ role: string; text?: string }>
    }
    expect(freshRequest.messages).toEqual([
      expect.objectContaining({ role: 'user', text: 'Fresh context' }),
    ])
  })

  it.each(['logout', 'dispose'] as const)(
    'suppresses late transport and proposal callbacks after %s',
    async (endSession) => {
      const harness = transportHarness()
      const proposals = proposalsHarness()
      const listener = vi.fn()
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: { id: 'test', systemPrompt: 'test', tools: [], executeTool: vi.fn() },
        proposals: proposals.controller,
      })
      session.subscribe(listener)
      session.send('Old request')
      await Promise.resolve()
      const callbacks = harness.callbacks()

      session[endSession]()
      listener.mockClear()
      callbacks.onDelta('Late answer')
      callbacks.onDone()
      if (endSession === 'dispose') proposals.setPending()
      await Promise.resolve()

      expect(session.snapshot()).toMatchObject({
        assistantText: '',
        busy: false,
        timeline: [],
        proposal: undefined,
      })
      expect(listener).not.toHaveBeenCalled()
      if (endSession === 'dispose') {
        session.send('Must not run')
        session.stop()
        session.reject()
        session.newTask()
        session.retry()
        session.logout()
        session.authenticationLost()
        await Promise.resolve()
        expect(harness.stream).toHaveBeenCalledOnce()
        expect(proposals.controller.logout).toHaveBeenCalledOnce()
        expect(proposals.controller.reject).not.toHaveBeenCalled()
      }
    },
  )
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
      subscribe: () => () => undefined,
      waitForDecision: vi.fn(async () => ({ status: 'cancelled' as const })),
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
    [
      'office_overwrite_required',
      'The target cells contain data. Choose an empty range or explicitly allow overwrite.',
    ],
    ['office_recovery_failed', 'The document could not be restored after the failed change.'],
    [
      'office_concurrent_change',
      'The document changed during the operation. Inspect it before trying again.',
    ],
    [
      'office_state_uncertain',
      'The change may be partially applied. Inspect the document before trying again.',
    ],
    [
      'office_recovery_failed:word_body_shape',
      'The document could not be restored after the failed change (word_body_shape).',
    ],
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

  it('allows only one confirmation, blocks competing writes, and lets Stop abort applying', async () => {
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

    expect(session.snapshot()).toMatchObject({ applying: false, status: 'cancelled' })
    expect(proposals.controller.confirm).toHaveBeenCalledOnce()
    expect(proposals.controller.newTurn).toHaveBeenCalledOnce()
    expect(proposals.controller.reject).not.toHaveBeenCalled()
    expect(harness.cancel).not.toHaveBeenCalled()

    settle()
    await Promise.all([first, second])
    expect(session.snapshot().applying).toBe(false)
    expect(session.snapshot().proposal).toBeUndefined()
    session.logout()
    expect(proposals.controller.logout).toHaveBeenCalledOnce()
  })

  it('pauses the same agent turn for approval, applies immediately, then resumes once', async () => {
    const harness = transportHarness()
    const proposals = proposalsHarness()
    const session = createOfficeAgentSession({
      transport: harness.transport,
      skill: {
        id: 'test',
        systemPrompt: 'test',
        tools: [{ name: 'write_document', description: 'write', inputSchema: { type: 'object' } }],
        executeTool: vi.fn(() => {
          proposals.setPending()
          return {
            output: JSON.stringify({ status: 'awaiting_user_confirmation' }),
            mutated: false,
            summary: 'Awaiting confirmation',
          }
        }),
      },
      proposals: proposals.controller,
    })

    session.send('prepare an edit')
    await Promise.resolve()
    harness.callbacks().onToolCall({ id: 'write-1', name: 'write_document', input: {} })
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().proposal?.id).toBe('p1'))

    expect(session.snapshot()).toMatchObject({ busy: true, applying: false })
    expect(harness.stream).toHaveBeenCalledTimes(1)
    await session.confirm('p1')
    expect(proposals.controller.confirm).toHaveBeenCalledOnce()
    await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))
    expect(session.snapshot().timeline).toEqual(
      expect.arrayContaining([expect.objectContaining({ kind: 'proposal', state: 'applied' })]),
    )

    harness.callbacks().onDelta('The approved change is now applied.')
    harness.callbacks().onDone()
    await vi.waitFor(() => expect(session.snapshot().busy).toBe(false))
    expect(harness.stream).toHaveBeenCalledTimes(2)
  })

  it.each([
    ['word', 'write_document'],
    ['excel', 'set_cell_range'],
    ['powerpoint', 'edit_slide_text'],
  ] as const)(
    'returns the confirmed %s mutation to the same AgentLoop tool call',
    async (host, toolName) => {
      const harness = transportHarness()
      const proposals = createStructuredProposalController()
      const execute = vi.fn(async () => undefined)
      const session = createOfficeAgentSession({
        transport: harness.transport,
        skill: {
          id: host,
          systemPrompt: 'test',
          tools: [{ name: toolName, description: 'write', inputSchema: { type: 'object' } }],
          executeTool: vi.fn(() => {
            const proposal = proposals.propose({
              operation: toolName,
              toolName,
              title: `Confirm ${toolName}`,
              preview: { operation: toolName },
              impact: { host, targets: ['target-1'], count: 1 },
              fingerprint: 'v1',
              validate: async () => true,
              execute,
            })
            return {
              output: JSON.stringify({
                proposalId: proposal.id,
                status: 'awaiting_user_confirmation',
              }),
              mutated: false,
              summary: 'Awaiting confirmation',
            }
          }),
        },
        proposals,
      })

      session.send(`change ${host}`)
      await Promise.resolve()
      harness.callbacks().onToolCall({ id: `${host}-write`, name: toolName, input: {} })
      harness.callbacks().onDone()
      await vi.waitFor(() => expect(session.snapshot().proposal).toBeDefined())
      const id = session.snapshot().proposal!.id

      await session.confirm(id)
      await vi.waitFor(() => expect(harness.stream).toHaveBeenCalledTimes(2))

      expect(execute).toHaveBeenCalledOnce()
      const resumed = harness.stream.mock.calls[1]?.[0] as {
        messages: Array<{ role: string; results?: Array<{ output: string }> }>
      }
      expect(resumed.messages.at(-1)).toMatchObject({
        role: 'tool',
        results: [
          {
            output: JSON.stringify({ proposalId: id, status: 'applied' }),
          },
        ],
      })
    },
  )

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
