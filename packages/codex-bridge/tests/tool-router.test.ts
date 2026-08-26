import type { AgentSkill, AgentToolCall, ToolExecution } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import {
  DocumentToolRouter,
  type MutationExecutionGuard,
  type ToolSessionRegistration,
} from '../src/tool-router.js'

const readTool = {
  name: 'read_document',
  description: 'Read the current document.',
  inputSchema: { type: 'object', additionalProperties: false },
}
const mutateTool = {
  name: 'replace_text',
  description: 'Replace selected text.',
  inputSchema: {
    type: 'object',
    properties: { text: { type: 'string' } },
    required: ['text'],
    additionalProperties: false,
  },
}

function fixture(overrides: Partial<ToolSessionRegistration> = {}) {
  let revision = 'rev-1'
  let open = true
  const events: string[] = []
  const executeTool = vi.fn(async (call: AgentToolCall): Promise<ToolExecution> => {
    events.push(`read:${call.name}`)
    return { output: 'document text', summary: 'Read document' }
  })
  const skill: AgentSkill = {
    id: 'docs',
    systemPrompt: 'Use document tools.',
    tools: [readTool, mutateTool],
    executeTool,
  }
  const registration: ToolSessionRegistration = {
    skill,
    policy: { read_document: 'read', replace_text: 'mutate' },
    isOpen: () => open,
    getRevision: vi.fn(async () => {
      events.push('revision')
      return revision
    }),
    requestApproval: vi.fn(async () => {
      events.push('approval')
      return true
    }),
    captureSnapshot: vi.fn(async () => {
      events.push('snapshot')
      return 'snapshot-1'
    }),
    executeMutation: vi.fn(
      async (_call: AgentToolCall, guard: MutationExecutionGuard): Promise<ToolExecution> => {
        events.push(`mutate:${guard.expectedRevision}:${guard.snapshotId}`)
        revision = 'rev-2'
        return { output: 'updated', summary: 'Replace text', mutated: true }
      },
    ),
    validateMutation: vi.fn(async () => {
      events.push('validate')
    }),
    ...overrides,
  }
  const tokens = [
    Buffer.alloc(32, 1),
    Buffer.alloc(32, 2),
    Buffer.alloc(32, 3),
    Buffer.alloc(32, 4),
  ]
  const router = new DocumentToolRouter({ randomBytes: () => tokens.shift()! })
  return {
    router,
    registration,
    executeTool,
    events,
    setRevision: (value: string) => (revision = value),
    closeRenderer: () => (open = false),
  }
}

describe('document-scoped tool router', () => {
  it('creates high-entropy isolated credentials and scopes tool listing', () => {
    const a = fixture()
    const first = a.router.register(a.registration)
    const second = a.router.register({
      ...a.registration,
      skill: { ...a.registration.skill, id: 'sheets', tools: [readTool] },
      policy: { read_document: 'read' },
    })

    expect(first.sessionId).not.toBe(second.sessionId)
    expect(first.token).not.toBe(second.token)
    expect(Buffer.from(first.sessionId, 'base64url')).toHaveLength(32)
    expect(Buffer.from(first.token, 'base64url')).toHaveLength(32)
    expect(a.router.listTools(first)).toEqual([
      { ...readTool, annotations: { readOnlyHint: true, destructiveHint: false } },
      { ...mutateTool, annotations: { readOnlyHint: false, destructiveHint: true } },
    ])
    expect(() => a.router.listTools({ sessionId: first.sessionId, token: second.token })).toThrow(
      'mcp_unauthorized',
    )
    expect(() => a.router.listTools({ sessionId: second.sessionId, token: first.token })).toThrow(
      'mcp_unauthorized',
    )
  })

  it('executes a read tool without approval or snapshot', async () => {
    const f = fixture()
    const session = f.router.register(f.registration)

    await expect(
      f.router.callTool(session, { id: 'call-1', name: 'read_document', input: {} }),
    ).resolves.toEqual({ output: 'document text', summary: 'Read document' })
    expect(f.events).toEqual(['read:read_document'])
  })

  it('executes mutation only after approval, revision recheck, and snapshot', async () => {
    const f = fixture()
    const session = f.router.register(f.registration)

    await expect(
      f.router.callTool(session, {
        id: 'call-2',
        name: 'replace_text',
        input: { text: 'new' },
      }),
    ).resolves.toMatchObject({ output: 'updated', mutated: true })
    expect(f.events).toEqual([
      'revision',
      'approval',
      'revision',
      'snapshot',
      'revision',
      'mutate:rev-1:snapshot-1',
      'validate',
    ])
  })

  it('does not mutate after denial or a revision mismatch', async () => {
    const denied = fixture({ requestApproval: vi.fn(async () => false) })
    const deniedSession = denied.router.register(denied.registration)
    await expect(
      denied.router.callTool(deniedSession, { id: 'denied', name: 'replace_text', input: {} }),
    ).resolves.toMatchObject({ output: 'tool_denied', isError: false, mutated: false })
    expect(denied.registration.captureSnapshot).not.toHaveBeenCalled()
    expect(denied.registration.executeMutation).not.toHaveBeenCalled()

    const stale = fixture()
    vi.mocked(stale.registration.requestApproval).mockImplementation(async () => {
      stale.setRevision('rev-other')
      return true
    })
    const staleSession = stale.router.register(stale.registration)
    await expect(
      stale.router.callTool(staleSession, { id: 'stale', name: 'replace_text', input: {} }),
    ).resolves.toMatchObject({ output: 'document_changed', isError: true, mutated: false })
    expect(stale.registration.captureSnapshot).not.toHaveBeenCalled()
    expect(stale.registration.executeMutation).not.toHaveBeenCalled()
  })

  it('does not mutate when cancelled or when the renderer closes at any boundary', async () => {
    const cancelled = fixture()
    const controller = new AbortController()
    vi.mocked(cancelled.registration.captureSnapshot).mockImplementation(async () => {
      controller.abort()
      return 'snapshot-1'
    })
    const cancelledSession = cancelled.router.register(cancelled.registration)
    await expect(
      cancelled.router.callTool(
        cancelledSession,
        { id: 'cancelled', name: 'replace_text', input: {} },
        controller.signal,
      ),
    ).resolves.toMatchObject({ output: 'tool_cancelled', isError: true, mutated: false })
    expect(cancelled.registration.executeMutation).not.toHaveBeenCalled()

    const closed = fixture()
    vi.mocked(closed.registration.captureSnapshot).mockImplementation(async () => {
      closed.closeRenderer()
      return 'snapshot-1'
    })
    const closedSession = closed.router.register(closed.registration)
    await expect(
      closed.router.callTool(closedSession, { id: 'closed', name: 'replace_text', input: {} }),
    ).resolves.toMatchObject({ output: 'tool_session_closed', isError: true, mutated: false })
    expect(closed.registration.executeMutation).not.toHaveBeenCalled()
  })

  it('aborts pending work on cancellation and teardown, and refuses duplicate calls', async () => {
    let observedSignal: AbortSignal | undefined
    const f = fixture({
      skill: {
        id: 'slow',
        systemPrompt: '',
        tools: [readTool],
        executeTool: (_call, signal) =>
          new Promise((resolve) => {
            observedSignal = signal
            signal?.addEventListener('abort', () =>
              resolve({ output: 'tool_cancelled', isError: true, summary: 'Cancelled' }),
            )
          }),
      },
      policy: { read_document: 'read' },
    })
    const session = f.router.register(f.registration)
    const pending = f.router.callTool(session, { id: 'same', name: 'read_document', input: {} })
    await expect(
      f.router.callTool(session, { id: 'same', name: 'read_document', input: {} }),
    ).resolves.toMatchObject({ output: 'tool_call_in_progress', isError: true })
    expect(f.router.cancel(session, 'same')).toBe(true)
    expect(observedSignal?.aborted).toBe(true)
    await expect(pending).resolves.toMatchObject({ output: 'tool_cancelled' })

    const again = f.router.callTool(session, { id: 'next', name: 'read_document', input: {} })
    f.router.close(session)
    expect(observedSignal?.aborted).toBe(true)
    await expect(again).resolves.toMatchObject({ output: 'tool_cancelled' })
    expect(() => f.router.listTools(session)).toThrow('mcp_unauthorized')
  })

  it('settles cancellation even when a read adapter ignores AbortSignal', async () => {
    const f = fixture({
      skill: {
        id: 'ignores-abort',
        systemPrompt: '',
        tools: [readTool],
        executeTool: () => new Promise(() => undefined),
      },
      policy: { read_document: 'read' },
    })
    const session = f.router.register(f.registration)
    const pending = f.router.callTool(session, { id: 'ignored', name: 'read_document', input: {} })
    f.router.close(session)

    await expect(
      Promise.race([
        pending,
        new Promise((resolve) => setTimeout(() => resolve({ output: 'did_not_settle' }), 50)),
      ]),
    ).resolves.toMatchObject({ output: 'tool_cancelled' })
  })

  it('holds the one-call lock until an abort-ignoring mutation definitively settles', async () => {
    let settleMutation = (_execution: ToolExecution): void => undefined
    const f = fixture({
      executeMutation: vi.fn(
        () =>
          new Promise<ToolExecution>((resolve) => {
            settleMutation = resolve
          }),
      ),
    })
    const session = f.router.register(f.registration)
    const controller = new AbortController()
    const mutation = f.router.callTool(
      session,
      { id: 'mutation', name: 'replace_text', input: {} },
      controller.signal,
    )
    await vi.waitFor(() => expect(f.registration.executeMutation).toHaveBeenCalledOnce())
    controller.abort()

    await expect(
      f.router.callTool(session, { id: 'overlap', name: 'read_document', input: {} }),
    ).resolves.toMatchObject({ output: 'tool_call_in_progress', isError: true })
    settleMutation({ output: 'committed', summary: 'Replace', mutated: true })
    await expect(mutation).resolves.toMatchObject({ output: 'committed', mutated: true })
    expect(f.registration.validateMutation).toHaveBeenCalledOnce()
    await expect(
      f.router.callTool(session, { id: 'after', name: 'read_document', input: {} }),
    ).resolves.toMatchObject({ output: 'document text' })
  })

  it('rejects an aggregate registration larger than the document bridge budget', () => {
    const f = fixture()
    const tools = Array.from({ length: 17 }, (_, index) => ({
      name: `tool_${index}`,
      description: 'd'.repeat(250_000),
      inputSchema: { type: 'object' },
    }))
    expect(() =>
      f.router.register({
        ...f.registration,
        skill: { ...f.registration.skill, tools },
        policy: Object.fromEntries(tools.map((tool) => [tool.name, 'read'])),
      }),
    ).toThrow('tool_registration_limit')
  })

  it('requires an explicit policy and mutation lifecycle for every advertised tool', () => {
    const f = fixture()
    expect(() =>
      f.router.register({ ...f.registration, policy: { read_document: 'read' } }),
    ).toThrow('invalid_tool_policy')
    expect(() => f.router.register({ ...f.registration, executeMutation: undefined })).toThrow(
      'invalid_mutation_lifecycle',
    )
  })
})
