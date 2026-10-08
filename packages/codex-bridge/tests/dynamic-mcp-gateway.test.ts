import { describe, expect, it, vi } from 'vitest'
import { suspendToolExecution, type AgentToolCall, type ToolExecution } from '@wiswork/agent-core'
import { startDynamicMcpGateway } from '../src/dynamic-mcp-gateway.js'

const initialized = new Map<string, Promise<void>>()
const clientIds = new Map<string, string>()

function ensureInitialized(url: string, secret: string): Promise<void> {
  let pending = initialized.get(url)
  if (pending) return pending
  pending = (async () => {
    const headers: Record<string, string> = {
      authorization: `Bearer ${secret}`,
      'content-type': 'application/json',
    }
    const response = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        jsonrpc: '2.0',
        id: 'init',
        method: 'initialize',
        params: {
          protocolVersion: '2025-06-18',
          capabilities: { elicitation: { form: {}, url: {} } },
          clientInfo: { name: 'codex-mcp-client', title: 'Codex', version: '0.147.0' },
        },
      }),
    })
    headers['Mcp-Session-Id'] = response.headers.get('mcp-session-id')!
    clientIds.set(url, headers['Mcp-Session-Id'])
    await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} }),
    })
  })()
  initialized.set(url, pending)
  return pending
}

async function rpc(url: string, secret: string, id: number, method: string, params: object) {
  await ensureInitialized(url, secret)
  return fetch(url, {
    method: 'POST',
    headers: {
      authorization: `Bearer ${secret}`,
      'content-type': 'application/json',
      'Mcp-Session-Id': clientIds.get(url)!,
    },
    body: JSON.stringify({ jsonrpc: '2.0', id, method, params }),
  })
}

describe('fixed dynamic MCP gateway', () => {
  it('starts before documents, advertises only one generic tool, and denies empty calls', async () => {
    const gateway = await startDynamicMcpGateway()
    try {
      const listed = await rpc(gateway.url, gateway.secret, 1, 'tools/list', {})
      const body = await listed.json()
      expect(body.result.tools.map((tool: { name: string }) => tool.name)).toEqual([
        'wiswork_read',
        'wiswork_propose',
      ])
      expect(body.result.tools[0].annotations).toEqual({
        readOnlyHint: true,
        destructiveHint: false,
      })
      expect(JSON.stringify(body)).not.toMatch(/doc-|session_|owner/i)
      const denied = await rpc(gateway.url, gateway.secret, 2, 'tools/call', {
        name: 'wiswork_read',
        arguments: { capability: 'A'.repeat(43), callId: 'x', toolName: 'read_blocks', input: {} },
      })
      expect(denied.status).toBe(403)
    } finally {
      await gateway.close()
    }
  })

  it('consumes the exact turn capability before dispatch and revokes it with the document', async () => {
    const execute = vi.fn(async (_call: unknown) => ({ output: 'ok', summary: 'ok' }))
    const session = {
      identity: { ownerId: 'owner', documentId: 'doc', generation: 7 },
      credentials: { sessionId: 's', secret: 'k' },
      listTools: () => [
        { name: 'read_blocks', annotations: { readOnlyHint: true, destructiveHint: false } },
      ],
      callTool: (_credentials: unknown, call: unknown) => execute(call),
    } as any
    const gateway = await startDynamicMcpGateway()
    const unregister = gateway.register({
      ownerId: 'owner',
      documentId: 'doc',
      generation: 7,
      session,
    })
    const { capability } = gateway.beginTurn({
      documentId: 'doc',
      generation: 7,
      threadId: 'thread',
    })
    const args = { capability, callId: 'call-1', toolName: 'read_blocks', input: { index: 1 } }
    try {
      expect(
        (
          await rpc(gateway.url, gateway.secret, 1, 'tools/call', {
            name: 'wiswork_read',
            arguments: args,
          })
        ).status,
      ).toBe(200)
      expect(execute).toHaveBeenCalledOnce()
      expect(
        (
          await rpc(gateway.url, gateway.secret, 2, 'tools/call', {
            name: 'wiswork_read',
            arguments: args,
          })
        ).status,
      ).toBe(403)
      unregister()
      expect(
        (
          await rpc(gateway.url, gateway.secret, 3, 'tools/call', {
            name: 'wiswork_read',
            arguments: { ...args, callId: 'call-2' },
          })
        ).status,
      ).toBe(403)
      expect(execute).toHaveBeenCalledOnce()
    } finally {
      await gateway.close()
    }
  })

  it('keeps concurrent document grants isolated and rejects replayed call IDs', async () => {
    const gateway = await startDynamicMcpGateway()
    const calls = [
      vi.fn(async (_call: unknown) => ({ output: 'a', summary: 'a' })),
      vi.fn(async (_call: unknown) => ({ output: 'b', summary: 'b' })),
    ]
    const registrations = calls.map((execute, index) => ({
      ownerId: `owner-${index}`,
      documentId: `doc-${index}`,
      generation: 1,
      session: {
        credentials: { sessionId: `s-${index}`, secret: `k-${index}` },
        listTools: () => [
          { name: 'read_blocks', annotations: { readOnlyHint: true, destructiveHint: false } },
        ],
        callTool: (_credentials: unknown, call: unknown) => execute(call),
      } as any,
    }))
    const closes = registrations.map((registration) => gateway.register(registration))
    try {
      const grants = registrations.map((registration, index) =>
        gateway.beginTurn({
          documentId: registration.documentId,
          generation: 1,
          threadId: `thread-${index}`,
        }),
      )
      const first = await rpc(gateway.url, gateway.secret, 1, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: grants[0]!.capability,
          callId: 'a',
          toolName: 'read_blocks',
          input: {},
        },
      })
      const second = await rpc(gateway.url, gateway.secret, 2, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: grants[1]!.capability,
          callId: 'b',
          toolName: 'read_blocks',
          input: {},
        },
      })
      expect([first.status, second.status]).toEqual([200, 200])
      expect(calls[0]).toHaveBeenCalledWith(expect.objectContaining({ id: 'a' }))
      expect(calls[1]).toHaveBeenCalledWith(expect.objectContaining({ id: 'b' }))
      const replay = await rpc(gateway.url, gateway.secret, 3, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: grants[0]!.capability,
          callId: 'a',
          toolName: 'read_blocks',
          input: {},
        },
      })
      expect(replay.status).toBe(403)
      expect(calls[1]).toHaveBeenCalledTimes(1)
    } finally {
      closes.forEach((close) => close())
      await gateway.close()
    }
  })

  it('allows long visual review workflows without a per-turn call quota', async () => {
    const execute = vi.fn(async (_call: unknown) => ({ output: 'ok', summary: 'ok' }))
    const gateway = await startDynamicMcpGateway()
    const close = gateway.register({
      ownerId: 'owner',
      documentId: 'doc-budget',
      generation: 1,
      session: {
        credentials: { sessionId: 'session', secret: 'secret' },
        listTools: () => [
          { name: 'read_blocks', annotations: { readOnlyHint: true, destructiveHint: false } },
        ],
        callTool: (_credentials: unknown, call: unknown) => execute(call),
      } as any,
    })
    try {
      const grant = gateway.beginTurn({
        documentId: 'doc-budget',
        generation: 1,
        threadId: 'thread',
      })
      for (let index = 0; index < 80; index += 1) {
        const response = await rpc(gateway.url, gateway.secret, index + 1, 'tools/call', {
          name: 'wiswork_read',
          arguments: {
            capability: grant.capability,
            callId: `call-${index}`,
            toolName: 'read_blocks',
            input: {},
          },
        })
        expect(response.status).toBe(200)
      }
      expect(execute).toHaveBeenCalledTimes(80)
    } finally {
      close()
      await gateway.close()
    }
  })

  it('keeps an active turn capability alive across its original TTL', async () => {
    let now = 1_000
    const clock = vi.spyOn(Date, 'now').mockImplementation(() => now)
    const execute = vi.fn(async (_call: unknown) => ({ output: 'ok', summary: 'ok' }))
    const gateway = await startDynamicMcpGateway()
    const close = gateway.register({
      ownerId: 'owner',
      documentId: 'doc-active',
      generation: 1,
      session: {
        credentials: { sessionId: 'session', secret: 'secret' },
        listTools: () => [
          { name: 'read_blocks', annotations: { readOnlyHint: true, destructiveHint: false } },
        ],
        callTool: (_credentials: unknown, call: unknown) => execute(call),
      } as any,
    })
    try {
      const grant = gateway.beginTurn({
        documentId: 'doc-active',
        generation: 1,
        threadId: 'thread',
        ttlMs: 1_000,
      })
      now = 1_900
      const first = await rpc(gateway.url, gateway.secret, 40, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: grant.capability,
          callId: 'call-1',
          toolName: 'read_blocks',
          input: {},
        },
      })
      now = 2_800
      const second = await rpc(gateway.url, gateway.secret, 41, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: grant.capability,
          callId: 'call-2',
          toolName: 'read_blocks',
          input: {},
        },
      })
      now = 3_801
      const expired = await rpc(gateway.url, gateway.secret, 42, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: grant.capability,
          callId: 'call-3',
          toolName: 'read_blocks',
          input: {},
        },
      })
      expect([first.status, second.status]).toEqual([200, 200])
      expect(expired.status).toBe(403)
      expect(execute).toHaveBeenCalledTimes(2)
    } finally {
      clock.mockRestore()
      close()
      await gateway.close()
    }
  })

  it.each([600_000, 90_000])(
    'completes a pending mutation with consent capped by the %i ms grant',
    async (ttlMs) => {
      const onProposal = vi.fn()
      const writer = vi.fn()
      let settle!: (value: any) => void
      const never = new Promise<any>((resolve) => {
        settle = resolve
      })
      const diagnostics: string[] = []
      const gateway = await startDynamicMcpGateway((code) => diagnostics.push(code))
      const close = gateway.register({
        ownerId: 'owner',
        documentId: 'doc-proposal',
        generation: 1,
        onProposal,
        summarizeProposal: () => ({
          operation: 'replace',
          target: 'blocks',
          scope: 'bounded-set',
          count: 1,
        }),
        session: {
          credentials: { sessionId: 'session', secret: 'secret' },
          listTools: () => [
            {
              name: 'replace_text',
              annotations: { readOnlyHint: false, destructiveHint: true },
            },
          ],
          callTool: () => suspendToolExecution(never),
          cancelAll: writer,
        } as any,
      })
      try {
        const grant = gateway.beginTurn({
          documentId: 'doc-proposal',
          generation: 1,
          threadId: 'thread',
          ttlMs,
        })
        let returned = false
        const pendingResponse = rpc(gateway.url, gateway.secret, 50, 'tools/call', {
          name: 'wiswork_propose',
          arguments: {
            capability: grant.capability,
            callId: 'proposal-call',
            toolName: 'replace_text',
            input: { text: 'pending' },
          },
        }).then((response) => {
          returned = true
          return response
        })
        await vi.waitFor(() => expect(onProposal).toHaveBeenCalledOnce())
        expect(returned).toBe(false)
        settle({
          output: 'Images are supported only by cover and split_image layouts',
          summary: 'Build failed',
          isError: true,
          mutated: false,
        })
        const response = await pendingResponse
        expect(response.status).toBe(200)
        expect(await response.json()).toEqual(
          expect.objectContaining({
            result: expect.objectContaining({
              content: [
                expect.objectContaining({
                  type: 'text',
                  text: 'Images are supported only by cover and split_image layouts',
                }),
              ],
              isError: true,
            }),
          }),
        )
        expect(onProposal).toHaveBeenCalledOnce()
        expect(diagnostics).toContain('gateway_tool_call_failed')
        expect(diagnostics).not.toContain('gateway_tool_call_completed')
        expect(onProposal.mock.calls[0]![0].expiresAt - Date.now()).toBeGreaterThan(
          Math.min(ttlMs, 300_000) - 10_000,
        )
        expect(onProposal.mock.calls[0]![0].expiresAt - Date.now()).toBeLessThanOrEqual(
          Math.min(ttlMs, 300_000),
        )
        expect(onProposal).toHaveBeenCalledWith(
          expect.objectContaining({
            call: expect.objectContaining({ name: 'replace_text' }),
            summary: { operation: 'replace', target: 'blocks', scope: 'bounded-set', count: 1 },
            settled: never,
          }),
        )
        expect(writer).not.toHaveBeenCalled()
      } finally {
        close()
        await gateway.close()
      }
    },
  )

  it('fails closed before suspending a mutation that has no safe host summary', async () => {
    const callTool = vi.fn()
    const gateway = await startDynamicMcpGateway()
    gateway.register({
      ownerId: 'owner',
      documentId: 'doc-summary',
      generation: 1,
      session: {
        credentials: { sessionId: 's', secret: 'k' },
        listTools: () => [{ name: 'replace_text', annotations: { destructiveHint: true } }],
        callTool,
      } as any,
    })
    try {
      const grant = gateway.beginTurn({ documentId: 'doc-summary', generation: 1, threadId: 't' })
      const response = await rpc(gateway.url, gateway.secret, 51, 'tools/call', {
        name: 'wiswork_propose',
        arguments: {
          capability: grant.capability,
          callId: 'c',
          toolName: 'replace_text',
          input: {},
        },
      })
      expect(response.status).toBe(403)
      expect(callTool).not.toHaveBeenCalled()
    } finally {
      await gateway.close()
    }
  })

  it('revokes a turn capability explicitly and bounds outstanding grants', async () => {
    const execute = vi.fn(async () => ({ output: 'ok', summary: 'ok' }))
    const gateway = await startDynamicMcpGateway()
    gateway.register({
      ownerId: 'owner',
      documentId: 'doc',
      generation: 1,
      session: {
        credentials: { sessionId: 's', secret: 'k' },
        listTools: () => [
          { name: 'read_blocks', annotations: { readOnlyHint: true, destructiveHint: false } },
        ],
        callTool: execute,
      } as any,
    })
    try {
      const grant = gateway.beginTurn({ documentId: 'doc', generation: 1, threadId: 'thread' })
      gateway.revokeTurn(grant.capability)
      const denied = await rpc(gateway.url, gateway.secret, 90, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: grant.capability,
          callId: 'call',
          toolName: 'read_blocks',
          input: {},
        },
      })
      expect(denied.status).toBe(403)
      expect(execute).not.toHaveBeenCalled()
      for (let index = 0; index < 64; index += 1) {
        gateway.beginTurn({ documentId: 'doc', generation: 1, threadId: `thread-${index}` })
      }
      expect(() =>
        gateway.beginTurn({ documentId: 'doc', generation: 1, threadId: 'overflow' }),
      ).toThrow('turn_capability_limit')
    } finally {
      await gateway.close()
    }
  })

  it('balances tool events exactly once when document execution fails', async () => {
    const events: unknown[] = []
    const gateway = await startDynamicMcpGateway()
    gateway.register({
      ownerId: 'owner',
      documentId: 'doc',
      generation: 1,
      onToolEvent: (event) => events.push(event),
      session: {
        credentials: { sessionId: 's', secret: 'k' },
        listTools: () => [
          { name: 'read_blocks', annotations: { readOnlyHint: true, destructiveHint: false } },
        ],
        callTool: () => {
          throw new Error('private')
        },
        cancelAll: () => 0,
      } as any,
    })
    const grant = gateway.beginTurn({ documentId: 'doc', generation: 1, threadId: 'thread' })
    try {
      const response = await rpc(gateway.url, gateway.secret, 91, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: grant.capability,
          callId: 'failed-call',
          toolName: 'read_blocks',
          input: {},
        },
      })
      expect(response.status).toBe(403)
      expect(events).toEqual([
        {
          type: 'tool-start',
          callId: 'failed-call',
          toolName: 'read_blocks',
          turnId: expect.any(String),
        },
        {
          type: 'tool-complete',
          callId: 'failed-call',
          toolName: 'read_blocks',
          isError: true,
          turnId: expect.any(String),
          errorCode: 'tool_execution_failed',
        },
      ])
    } finally {
      await gateway.close()
    }
  })

  it.each(['tool_call_in_progress', 'invalid_tool_call', 'mutation_queue_full'])(
    'publishes a safe semantic receipt for a router rejection: %s',
    async (errorCode) => {
      const events: unknown[] = []
      const gateway = await startDynamicMcpGateway()
      const mutation = errorCode === 'mutation_queue_full'
      const execute = vi.fn((_credentials: unknown, _call: AgentToolCall) => ({
        output: errorCode,
        summary: 'Rejected',
        isError: true,
      }))
      gateway.register({
        ownerId: 'owner',
        documentId: 'doc',
        generation: 1,
        onToolEvent: (event) => events.push(event),
        summarizeProposal: () => ({ title: 'Change', details: [] }) as any,
        session: {
          credentials: { sessionId: 's', secret: 'k' },
          listTools: () => [
            {
              name: 'known_tool',
              annotations: { readOnlyHint: !mutation, destructiveHint: mutation },
            },
          ],
          callTool: execute,
          cancelAll: () => 0,
        } as any,
      })
      const grant = gateway.beginTurn({ documentId: 'doc', generation: 1, threadId: 'thread' })
      try {
        const response = await rpc(gateway.url, gateway.secret, 101, 'tools/call', {
          name: mutation ? 'wiswork_propose' : 'wiswork_read',
          arguments: {
            capability: grant.capability,
            callId: 'rejected',
            toolName: 'known_tool',
            input: {},
          },
        })
        expect(response.status).toBe(200)
        expect((await response.json()).result).toMatchObject({
          isError: true,
          content: [{ type: 'text', text: errorCode }],
        })
        expect(events).toEqual([
          {
            type: 'tool-start',
            callId: 'rejected',
            toolName: 'known_tool',
            turnId: expect.any(String),
          },
          {
            type: 'tool-complete',
            callId: 'rejected',
            toolName: 'known_tool',
            turnId: expect.any(String),
            isError: true,
            errorCode,
          },
        ])
        const start = events[0] as { turnId: string }
        expect(start.turnId).not.toBe(grant.capability)
        expect(execute.mock.calls[0]?.[1].invocationId).toBe(`${start.turnId}:rejected`)
        expect((events[1] as { turnId: string }).turnId).toBe(start.turnId)
      } finally {
        await gateway.close()
      }
    },
  )

  it('keeps delayed receipts bound to their revoked grant and redacts arbitrary failure output', async () => {
    const events: Array<{ type: string; turnId?: string; errorCode?: string }> = []
    let resolveOld!: (value: ToolExecution) => void
    const oldResult = new Promise<ToolExecution>((resolve) => {
      resolveOld = resolve
    })
    const gateway = await startDynamicMcpGateway()
    gateway.register({
      ownerId: 'owner',
      documentId: 'doc',
      generation: 1,
      onToolEvent: (event) => events.push(event),
      session: {
        credentials: { sessionId: 's', secret: 'k' },
        listTools: () => [
          { name: 'read_blocks', annotations: { readOnlyHint: true, destructiveHint: false } },
        ],
        callTool: (_credentials: unknown, call: AgentToolCall) =>
          call.id === 'old' ? oldResult : { output: 'ok' },
        cancelAll: () => 0,
      } as any,
    })
    const first = gateway.beginTurn({ documentId: 'doc', generation: 1, threadId: 'thread' })
    try {
      const pending = rpc(gateway.url, gateway.secret, 102, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: first.capability,
          callId: 'old',
          toolName: 'read_blocks',
          input: {},
        },
      })
      await vi.waitFor(() => expect(events).toHaveLength(1))
      gateway.revokeTurn(first.capability)
      const second = gateway.beginTurn({ documentId: 'doc', generation: 1, threadId: 'thread' })
      await rpc(gateway.url, gateway.secret, 103, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: second.capability,
          callId: 'new',
          toolName: 'read_blocks',
          input: {},
        },
      })
      resolveOld({
        output: 'private provider data https://secret.example/token',
        summary: 'Failed',
        isError: true,
      })
      await pending
      expect(events).toHaveLength(4)
      expect(events[0]?.turnId).toEqual(expect.any(String))
      expect(events[1]?.turnId).not.toBe(events[0]?.turnId)
      expect(events[3]).toMatchObject({
        type: 'tool-complete',
        turnId: events[0]?.turnId,
        errorCode: 'tool_execution_failed',
      })
      expect(JSON.stringify(events)).not.toMatch(/private|secret|token|capability/)
    } finally {
      resolveOld({ output: 'cancelled', summary: 'Cancelled', isError: true })
      await gateway.close()
    }
  })

  it('returns an authorized Office execution failure as a tool error instead of a capability denial', async () => {
    const diagnostics: string[] = []
    const gateway = await startDynamicMcpGateway((code) => diagnostics.push(code))
    gateway.register({
      ownerId: 'owner',
      documentId: 'doc-office-failure',
      generation: 1,
      session: {
        credentials: { sessionId: 's', secret: 'k' },
        listTools: () => [
          { name: 'read_slide', annotations: { readOnlyHint: true, destructiveHint: false } },
        ],
        callTool: async () => {
          throw new Error('office_write_failed')
        },
        cancelAll: () => 0,
      } as any,
    })
    const grant = gateway.beginTurn({
      documentId: 'doc-office-failure',
      generation: 1,
      threadId: 'thread',
    })
    try {
      const response = await rpc(gateway.url, gateway.secret, 92, 'tools/call', {
        name: 'wiswork_read',
        arguments: {
          capability: grant.capability,
          callId: 'office-failed-call',
          toolName: 'read_slide',
          input: {},
        },
      })
      expect(response.status).toBe(200)
      expect(await response.json()).toEqual(
        expect.objectContaining({
          result: expect.objectContaining({
            content: [{ type: 'text', text: 'office_write_failed' }],
            isError: true,
          }),
        }),
      )
      expect(diagnostics).toContain('gateway_tool_call_failed')
      expect(diagnostics).not.toContain('gateway_tool_call_denied')
    } finally {
      await gateway.close()
    }
  })
})
