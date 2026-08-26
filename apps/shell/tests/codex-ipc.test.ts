import { EventEmitter } from 'node:events'
import type { ToolSessionRegistration } from '@wiswork/codex-bridge'
import { describe, expect, it, vi } from 'vitest'
import { registerCodexToolIpc } from '../src/main/codex-ipc'
import { CODEX_TOOL_CHANNELS, type CodexToolRequest } from '../src/shared/codex-api'

class FakeSender extends EventEmitter {
  readonly sent: Array<{ channel: string; payload: unknown }> = []
  destroyed = false

  isDestroyed(): boolean {
    return this.destroyed
  }

  send(channel: string, payload: unknown): void {
    if (this.destroyed) throw new Error('destroyed')
    this.sent.push({ channel, payload })
  }

  destroy(): void {
    this.destroyed = true
    this.emit('destroyed')
  }
}

function fixture() {
  const handlers = new Map<string, (event: { sender: FakeSender }, ...args: unknown[]) => unknown>()
  const sessions: ToolSessionRegistration[] = []
  const close = vi.fn()
  const onRegister = vi.fn((registration: ToolSessionRegistration) => {
    sessions.push(registration)
    return { close }
  })
  const bridge = registerCodexToolIpc({
    ipcMain: { handle: (channel, handler) => handlers.set(channel, handler as never) },
    ownsDocument: () => true,
    onRegister: ({ registration }) => onRegister(registration),
    requestTimeoutMs: 100,
    randomId: (() => {
      let value = 0
      return () => `request-${++value}`
    })(),
  })
  const sender = new FakeSender()
  const registration = {
    documentId: 'document-1',
    skill: {
      id: 'docs',
      systemPrompt: 'Use document tools.',
      tools: [
        {
          name: 'read_document',
          description: 'Read document.',
          inputSchema: { type: 'object', additionalProperties: false },
        },
        {
          name: 'replace_text',
          description: 'Replace text.',
          inputSchema: { type: 'object', additionalProperties: false },
        },
      ],
    },
    policy: { read_document: 'read', replace_text: 'mutate' },
  } as const
  return { handlers, sessions, close, onRegister, bridge, sender, registration }
}

function lastRequest(sender: FakeSender): CodexToolRequest {
  return sender.sent.at(-1)!.payload as CodexToolRequest
}

describe('Codex renderer tool IPC', () => {
  it('registers a bounded remote AgentSkill without returning MCP identity', async () => {
    const f = fixture()
    const ack = await f.handlers.get(CODEX_TOOL_CHANNELS.register)!(
      { sender: f.sender },
      f.registration,
    )
    expect(ack).toEqual({ registered: true })
    expect(f.onRegister).toHaveBeenCalledOnce()
    expect(f.sessions[0]).toMatchObject({
      skill: { id: 'docs', systemPrompt: 'Use document tools.', tools: f.registration.skill.tools },
      policy: f.registration.policy,
    })
    expect(JSON.stringify(ack)).not.toContain('token')
    expect(JSON.stringify(ack)).not.toContain('url')
  })

  it('routes execution only to the exact owning sender and validates its response shape', async () => {
    const f = fixture()
    await f.handlers.get(CODEX_TOOL_CHANNELS.register)!({ sender: f.sender }, f.registration)
    const call = { id: 'call-1', name: 'read_document', input: {} }
    const execution = f.sessions[0]!.skill.executeTool(call)
    expect(f.sender.sent).toHaveLength(1)
    expect(lastRequest(f.sender)).toMatchObject({
      type: 'execute',
      requestId: 'request-1',
      documentId: 'document-1',
      call,
    })

    const attacker = new FakeSender()
    await expect(
      f.handlers.get(CODEX_TOOL_CHANNELS.response)!(
        { sender: attacker },
        {
          requestId: 'request-1',
          ok: true,
          type: 'execution',
          execution: { output: 'forged', summary: 'Forged' },
        },
      ),
    ).rejects.toThrow('untrusted_codex_tool_response')
    expect(
      await f.handlers.get(CODEX_TOOL_CHANNELS.response)!(
        { sender: f.sender },
        {
          requestId: 'request-1',
          ok: true,
          type: 'execution',
          execution: { output: 'text', summary: 'Read' },
        },
      ),
    ).toBe(true)
    await expect(execution).resolves.toEqual({ output: 'text', summary: 'Read' })
  })

  it('exposes revision, approval, snapshot, and atomic guarded mutation requests', async () => {
    const f = fixture()
    await f.handlers.get(CODEX_TOOL_CHANNELS.register)!({ sender: f.sender }, f.registration)
    const remote = f.sessions[0]!
    const call = { id: 'call-2', name: 'replace_text', input: { text: 'new' } }

    const revision = remote.getRevision()
    expect(lastRequest(f.sender).type).toBe('revision')
    await f.handlers.get(CODEX_TOOL_CHANNELS.response)!(
      { sender: f.sender },
      {
        requestId: lastRequest(f.sender).requestId,
        ok: true,
        type: 'revision',
        revision: 'rev-1',
      },
    )
    await expect(revision).resolves.toBe('rev-1')

    const approval = remote.requestApproval(call, 'rev-1')
    expect(lastRequest(f.sender)).toMatchObject({ type: 'approval', expectedRevision: 'rev-1' })
    await f.handlers.get(CODEX_TOOL_CHANNELS.response)!(
      { sender: f.sender },
      {
        requestId: lastRequest(f.sender).requestId,
        ok: true,
        type: 'approval',
        approved: true,
      },
    )
    await expect(approval).resolves.toBe(true)

    const snapshot = remote.captureSnapshot(call, 'rev-1')
    expect(lastRequest(f.sender)).toMatchObject({ type: 'snapshot', expectedRevision: 'rev-1' })
    await f.handlers.get(CODEX_TOOL_CHANNELS.response)!(
      { sender: f.sender },
      {
        requestId: lastRequest(f.sender).requestId,
        ok: true,
        type: 'snapshot',
        snapshotId: 'snapshot-1',
      },
    )
    await expect(snapshot).resolves.toBe('snapshot-1')

    const mutation = remote.executeMutation!(call, {
      expectedRevision: 'rev-1',
      snapshotId: 'snapshot-1',
    })
    expect(lastRequest(f.sender)).toMatchObject({
      type: 'executeMutation',
      guard: { expectedRevision: 'rev-1', snapshotId: 'snapshot-1' },
    })
    await f.handlers.get(CODEX_TOOL_CHANNELS.response)!(
      { sender: f.sender },
      {
        requestId: lastRequest(f.sender).requestId,
        ok: true,
        type: 'execution',
        execution: { output: 'updated', summary: 'Replace', mutated: true },
      },
    )
    await mutation
    const requestsBeforeValidation = f.sender.sent.length
    await remote.validateMutation!(
      call,
      { output: 'updated', summary: 'Replace', mutated: true },
      {
        expectedRevision: 'rev-1',
        snapshotId: 'snapshot-1',
      },
    )
    expect(f.sender.sent).toHaveLength(requestsBeforeValidation)
  })

  it('cancels renderer work and ignores late responses after abort, unregister, or destruction', async () => {
    const f = fixture()
    await f.handlers.get(CODEX_TOOL_CHANNELS.register)!({ sender: f.sender }, f.registration)
    const controller = new AbortController()
    const pending = f.sessions[0]!.skill.executeTool(
      { id: 'call-1', name: 'read_document', input: {} },
      controller.signal,
    )
    controller.abort()
    expect(f.sender.sent.at(-1)).toMatchObject({
      channel: CODEX_TOOL_CHANNELS.cancel,
      payload: { requestId: 'request-1', documentId: 'document-1' },
    })
    await expect(pending).rejects.toThrow('tool_cancelled')
    expect(
      await f.handlers.get(CODEX_TOOL_CHANNELS.response)!(
        { sender: f.sender },
        {
          requestId: 'request-1',
          ok: false,
          code: 'tool_cancelled',
        },
      ),
    ).toBe(false)

    const next = f.sessions[0]!.getRevision()
    await f.handlers.get(CODEX_TOOL_CHANNELS.unregister)!({ sender: f.sender }, 'document-1')
    await expect(next).rejects.toThrow('tool_session_closed')
    expect(f.sender.sent.at(-1)).toMatchObject({
      channel: CODEX_TOOL_CHANNELS.cancel,
      payload: { requestId: 'request-2', documentId: 'document-1' },
    })
    expect(f.close).toHaveBeenCalledOnce()

    await f.handlers.get(CODEX_TOOL_CHANNELS.register)!({ sender: f.sender }, f.registration)
    const destroyed = f.sessions[1]!.getRevision()
    f.sender.destroy()
    await expect(destroyed).rejects.toThrow('tool_session_closed')
    expect(f.close).toHaveBeenCalledTimes(2)
  })

  it('waits for guarded mutation cancellation acknowledgement before releasing IPC work', async () => {
    const f = fixture()
    await f.handlers.get(CODEX_TOOL_CHANNELS.register)!({ sender: f.sender }, f.registration)
    const controller = new AbortController()
    const pending = f.sessions[0]!.executeMutation!(
      { id: 'mutation', name: 'replace_text', input: {} },
      { expectedRevision: 'rev-1', snapshotId: 'snapshot-1' },
      controller.signal,
    )
    controller.abort()
    expect(f.sender.sent.at(-1)).toMatchObject({
      channel: CODEX_TOOL_CHANNELS.cancel,
      payload: { requestId: 'request-1' },
    })
    let settled = false
    void pending.then(
      () => {
        settled = true
      },
      () => {
        settled = true
      },
    )
    await Promise.resolve()
    expect(settled).toBe(false)
    await f.handlers.get(CODEX_TOOL_CHANNELS.response)!(
      { sender: f.sender },
      { requestId: 'request-1', ok: false, code: 'tool_cancelled' },
    )
    await expect(pending).rejects.toThrow('tool_cancelled')
  })

  it('rejects aggregate tool registrations beyond the IPC budget', async () => {
    const f = fixture()
    const tools = Array.from({ length: 17 }, (_, index) => ({
      name: `tool_${index}`,
      description: 'd'.repeat(250_000),
      inputSchema: { type: 'object' },
    }))
    await expect(
      f.handlers.get(CODEX_TOOL_CHANNELS.register)!(
        { sender: f.sender },
        {
          documentId: 'large',
          skill: { id: 'large', systemPrompt: '', tools },
          policy: Object.fromEntries(tools.map((tool) => [tool.name, 'read'])),
        },
      ),
    ).rejects.toThrow('invalid_codex_tool_registration')
    expect(f.onRegister).not.toHaveBeenCalled()
  })

  it('uses authoritative document ownership and global active identity across renderers', async () => {
    const f = fixture()
    const other = new FakeSender()
    await f.handlers.get(CODEX_TOOL_CHANNELS.register)!({ sender: f.sender }, f.registration)
    await expect(
      f.handlers.get(CODEX_TOOL_CHANNELS.register)!({ sender: other }, f.registration),
    ).rejects.toThrow('codex_tool_session_exists')

    const handlers = new Map<
      string,
      (event: { sender: FakeSender }, ...args: unknown[]) => unknown
    >()
    const onRegister = vi.fn(() => ({ close: vi.fn() }))
    registerCodexToolIpc({
      ipcMain: { handle: (channel, handler) => handlers.set(channel, handler as never) },
      ownsDocument: (owner, documentId) => owner === f.sender && documentId === 'document-1',
      onRegister,
    })
    await expect(
      handlers.get(CODEX_TOOL_CHANNELS.register)!({ sender: other }, f.registration),
    ).rejects.toThrow('untrusted_codex_tool_registration')
    expect(onRegister).not.toHaveBeenCalled()
  })

  it('rejects malformed registration and bounds failed/late responses without diagnostics content', async () => {
    vi.useFakeTimers()
    try {
      const diagnostics = vi.fn()
      const f = fixture()
      f.bridge.close()
      const handlers = new Map<
        string,
        (event: { sender: FakeSender }, ...args: unknown[]) => unknown
      >()
      registerCodexToolIpc({
        ipcMain: { handle: (channel, handler) => handlers.set(channel, handler as never) },
        ownsDocument: () => true,
        onRegister: ({ registration }) => f.onRegister(registration),
        requestTimeoutMs: 10,
        randomId: () => 'timeout-request',
        diagnostics,
      })
      await expect(
        handlers.get(CODEX_TOOL_CHANNELS.register)!(
          { sender: f.sender },
          {
            ...f.registration,
            policy: { read_document: 'read' },
          },
        ),
      ).rejects.toThrow('invalid_codex_tool_registration')
      expect(f.onRegister).not.toHaveBeenCalled()

      await handlers.get(CODEX_TOOL_CHANNELS.register)!({ sender: f.sender }, f.registration)
      const pending = f.sessions[0]!.getRevision()
      const rejection = expect(pending).rejects.toThrow('tool_ipc_timeout')
      await vi.advanceTimersByTimeAsync(10)
      await rejection
      expect(f.sender.sent.at(-1)).toMatchObject({
        channel: CODEX_TOOL_CHANNELS.cancel,
        payload: { requestId: 'timeout-request' },
      })
      expect(diagnostics).toHaveBeenCalledWith({ code: 'tool_ipc_timeout' })

      const mutation = f.sessions[0]!.executeMutation!(
        { id: 'mutation', name: 'replace_text', input: {} },
        { expectedRevision: 'rev-1', snapshotId: 'snapshot-1' },
      )
      const mutationRejection = expect(mutation).rejects.toThrow('tool_session_closed')
      await vi.advanceTimersByTimeAsync(10)
      await mutationRejection
      expect(f.close).toHaveBeenCalledOnce()
      const replacementRenderer = new FakeSender()
      await expect(
        handlers.get(CODEX_TOOL_CHANNELS.register)!(
          { sender: replacementRenderer },
          f.registration,
        ),
      ).rejects.toThrow('codex_tool_session_poisoned')
      expect(JSON.stringify(diagnostics.mock.calls)).not.toContain('Use document tools')
    } finally {
      vi.useRealTimers()
    }
  })
})
