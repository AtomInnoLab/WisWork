import type { AgentSkill, AgentToolCall, ToolExecution } from '@wiswork/agent-core'
import { suspendToolExecution } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { createDocumentToolSession, ToolRouterError } from '../src/tool-router.js'
import { createDocumentCarrierIssuer } from '../src/index.js'
import captured from './fixtures/codex-0147-request.json'

const readTool = { name: 'read_document', description: 'Read.', inputSchema: { type: 'object' } }
const writeTool = { name: 'replace_text', description: 'Write.', inputSchema: { type: 'object' } }

function fixture(executeTool?: AgentSkill['executeTool']) {
  const execute = vi.fn(
    executeTool ??
      (async (call: AgentToolCall): Promise<ToolExecution> => ({
        output: call.name === 'read_document' ? 'text' : 'changed',
        summary: 'done',
      })),
  )
  let open = true
  const session = createDocumentToolSession({
    identity: {
      ownerId: 'owner-1',
      host: 'docs',
      documentId: 'doc-1',
      sessionId: 'host-session-1',
      generation: 4,
    },
    skill: {
      id: 'docs',
      systemPrompt: '',
      tools: [readTool, writeTool],
      executeTool: execute,
    },
    policy: { read_document: 'read', replace_text: 'mutate' },
    isOpen: () => open,
  })
  return { session, execute, closeHost: () => (open = false) }
}

describe('document tool session', () => {
  it('uses cryptographic credentials and exact immutable document authority', () => {
    const f = fixture()
    expect(Buffer.from(f.session.credentials.sessionId, 'base64url')).toHaveLength(32)
    expect(Buffer.from(f.session.credentials.secret, 'base64url')).toHaveLength(32)
    expect(f.session.identity).toEqual({
      ownerId: 'owner-1',
      host: 'docs',
      documentId: 'doc-1',
      sessionId: 'host-session-1',
      generation: 4,
    })
    expect(Object.isFrozen(f.session.identity)).toBe(true)
    expect(() => f.session.listTools({ ...f.session.credentials, secret: 'A'.repeat(43) })).toThrow(
      'tool_unauthorized',
    )
  })

  it('runs reads but rejects a read that mutates or suspends', async () => {
    const f = fixture()
    await expect(
      f.session.callTool(f.session.credentials, { id: 'r1', name: 'read_document', input: {} }),
    ).resolves.toMatchObject({ output: 'text' })

    const mutating = fixture(async () => ({ output: 'bad', summary: 'bad', mutated: true }))
    await expect(
      mutating.session.callTool(mutating.session.credentials, {
        id: 'r2',
        name: 'read_document',
        input: {},
      }),
    ).resolves.toMatchObject({ output: 'tool_policy_violation', isError: true })

    const suspended = fixture(() =>
      suspendToolExecution(Promise.resolve({ output: 'bad', summary: 'bad' })),
    )
    await expect(
      suspended.session.callTool(suspended.session.credentials, {
        id: 'r3',
        name: 'read_document',
        input: {},
      }),
    ).resolves.toMatchObject({ output: 'tool_policy_violation', isError: true })
  })

  it('never approves, snapshots, or writes; mutations require a host suspension', async () => {
    const direct = fixture(async () => ({ output: 'changed', summary: 'changed', mutated: true }))
    await expect(
      direct.session.callTool(direct.session.credentials, {
        id: 'w1',
        name: 'replace_text',
        input: {},
      }),
    ).resolves.toMatchObject({
      output: 'mutation_authority_required',
      isError: true,
      mutated: false,
    })

    const authorized = fixture(() =>
      suspendToolExecution(
        Promise.resolve({ output: 'changed', summary: 'changed', mutated: true }),
      ),
    )
    await expect(
      authorized.session.callTool(authorized.session.credentials, {
        id: 'w2',
        name: 'replace_text',
        input: {},
      }),
    ).resolves.toMatchObject({ output: 'changed', mutated: true })
  })

  it('closes the Task 2 validator over a router-issued one-use carrier', () => {
    const validate = vi.fn((capability: unknown) => capability === 'opaque-host-proof')
    const issuer = createDocumentCarrierIssuer(
      { host: 'docs', documentId: 'doc-1', sessionId: 'session_1', generation: 4 },
      validate,
    )
    const execute = vi.fn(async () => ({ output: 'text', summary: 'read' }))
    const session = createDocumentToolSession({
      identity: {
        ownerId: 'owner-1',
        host: 'docs',
        documentId: 'doc-1',
        sessionId: 'session_1',
        generation: 4,
      },
      skill: { id: 'docs', systemPrompt: '', tools: [readTool], executeTool: execute },
      policy: { read_document: 'read' },
      isOpen: () => true,
      carrier: { issuer, capability: 'opaque-host-proof' },
    })
    const handle = session.issueCarrier(session.credentials, {
      turnId: 'turn_1',
      sourceNonce: 'N'.repeat(43),
      method: 'mcp__wiswork__wiswork_read_document',
      toolName: 'wiswork_read_document',
      schemaDigest: 'a'.repeat(64),
    })
    expect(validate).toHaveBeenCalledWith(
      'opaque-host-proof',
      expect.objectContaining({ documentId: 'doc-1', generation: 4 }),
    )
    expect(Object.keys(handle)).toEqual([])
    expect(() => issuer.prepareTurn(structuredClone(captured), {}, handle)).not.toThrow()
    expect(() => issuer.prepareTurn(structuredClone(captured), {}, handle)).toThrow(
      'carrier_authorization_consumed',
    )
  })

  it('fails closed for unknown/denied tools, cross-session credentials, cancellation and teardown', async () => {
    const a = fixture()
    const b = fixture()
    await expect(
      a.session.callTool(b.session.credentials, { id: 'x', name: 'read_document', input: {} }),
    ).rejects.toThrow('tool_unauthorized')
    await expect(
      a.session.callTool(a.session.credentials, { id: 'x', name: 'shell', input: {} }),
    ).resolves.toMatchObject({ output: 'unknown_tool', isError: true })
    const controller = new AbortController()
    controller.abort()
    await expect(
      a.session.callTool(
        a.session.credentials,
        { id: 'x2', name: 'read_document', input: {} },
        controller.signal,
      ),
    ).resolves.toMatchObject({ output: 'tool_cancelled', isError: true })
    a.session.close()
    expect(() => a.session.listTools(a.session.credentials)).toThrow('tool_session_closed')
  })

  it('enforces one active call and bounded definitions, input and output', async () => {
    let release!: (value: ToolExecution) => void
    const pending = new Promise<ToolExecution>((resolve) => (release = resolve))
    const f = fixture(() => pending)
    const first = f.session.callTool(f.session.credentials, {
      id: 'one',
      name: 'read_document',
      input: {},
    })
    await expect(
      f.session.callTool(f.session.credentials, { id: 'two', name: 'read_document', input: {} }),
    ).resolves.toMatchObject({ output: 'tool_call_in_progress', isError: true })
    release({ output: 'ok', summary: 'ok' })
    await first
    await expect(
      f.session.callTool(f.session.credentials, {
        id: 'big',
        name: 'read_document',
        input: { text: 'x'.repeat(1_000_001) },
      }),
    ).resolves.toMatchObject({ output: 'invalid_tool_call', isError: true })

    expect(() =>
      createDocumentToolSession({
        identity: f.session.identity,
        skill: {
          id: 'bad',
          systemPrompt: '',
          tools: [{ ...readTool, name: 'exec_command' }],
          executeTool: async () => ({ output: '', summary: '' }),
        },
        policy: { exec_command: 'read' },
        isOpen: () => true,
      }),
    ).toThrowError(ToolRouterError)
  })

  it('bounds read execution time and cancels pending host work on teardown', async () => {
    const signalSeen = vi.fn()
    const session = createDocumentToolSession({
      identity: {
        ownerId: 'owner-1',
        host: 'docs',
        documentId: 'doc-1',
        sessionId: 'session-1',
        generation: 1,
      },
      skill: {
        id: 'docs',
        systemPrompt: '',
        tools: [readTool],
        executeTool: async (_call, signal) => {
          signal?.addEventListener('abort', signalSeen, { once: true })
          return await new Promise<ToolExecution>(() => undefined)
        },
      },
      policy: { read_document: 'read' },
      isOpen: () => true,
      maxCallMs: 10,
    })
    await expect(
      session.callTool(session.credentials, { id: 'slow', name: 'read_document', input: {} }),
    ).resolves.toMatchObject({ output: 'tool_timeout', isError: true })

    const pending = session.callTool(session.credentials, {
      id: 'closing',
      name: 'read_document',
      input: {},
    })
    session.close()
    await expect(pending).resolves.toMatchObject({ output: 'tool_cancelled', isError: true })
    expect(signalSeen).toHaveBeenCalled()
  })
})
