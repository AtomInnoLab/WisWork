import { describe, expect, it, vi } from 'vitest'
import { suspendToolExecution } from '@wiswork/agent-core'
import { startDynamicMcpGateway } from '../src/dynamic-mcp-gateway.js'

const initialized = new Map<string, Promise<void>>()

function ensureInitialized(url: string, secret: string): Promise<void> {
  let pending = initialized.get(url)
  if (pending) return pending
  pending = (async () => {
    const headers = { authorization: `Bearer ${secret}`, 'content-type': 'application/json' }
    await fetch(url, {
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
    headers: { authorization: `Bearer ${secret}`, 'content-type': 'application/json' },
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

  it('allows eight read-to-proposal calls per turn and denies the ninth', async () => {
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
      for (let index = 0; index < 9; index += 1) {
        const response = await rpc(gateway.url, gateway.secret, index + 1, 'tools/call', {
          name: 'wiswork_read',
          arguments: {
            capability: grant.capability,
            callId: `call-${index}`,
            toolName: 'read_blocks',
            input: {},
          },
        })
        expect(response.status).toBe(index < 8 ? 200 : 403)
      }
      expect(execute).toHaveBeenCalledTimes(8)
    } finally {
      close()
      await gateway.close()
    }
  })

  it('completes a mutation proposal without awaiting or executing the suspended writer', async () => {
    const onProposal = vi.fn()
    const writer = vi.fn()
    const never = new Promise<never>(() => undefined)
    const gateway = await startDynamicMcpGateway()
    const close = gateway.register({
      ownerId: 'owner',
      documentId: 'doc-proposal',
      generation: 1,
      onProposal,
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
      })
      const response = await rpc(gateway.url, gateway.secret, 50, 'tools/call', {
        name: 'wiswork_propose',
        arguments: {
          capability: grant.capability,
          callId: 'proposal-call',
          toolName: 'replace_text',
          input: { text: 'pending' },
        },
      })
      expect(response.status).toBe(200)
      expect(await response.json()).toEqual(
        expect.objectContaining({
          result: expect.objectContaining({
            content: [
              expect.objectContaining({
                type: 'text',
                text: expect.stringContaining('proposalId'),
              }),
            ],
            isError: false,
          }),
        }),
      )
      expect(onProposal).toHaveBeenCalledOnce()
      expect(onProposal).toHaveBeenCalledWith(
        expect.objectContaining({ call: expect.objectContaining({ name: 'replace_text' }) }),
      )
      expect(writer).not.toHaveBeenCalled()
    } finally {
      close()
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
        { type: 'tool-start', callId: 'failed-call', toolName: 'read_blocks' },
        {
          type: 'tool-complete',
          callId: 'failed-call',
          toolName: 'read_blocks',
          isError: true,
        },
      ])
    } finally {
      await gateway.close()
    }
  })
})
