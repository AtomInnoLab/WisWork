import { describe, expect, it, vi } from 'vitest'
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
      expect(body.result.tools.map((tool: { name: string }) => tool.name)).toEqual(['wiswork_call'])
      expect(JSON.stringify(body)).not.toMatch(/doc-|session_|owner/i)
      const denied = await rpc(gateway.url, gateway.secret, 2, 'tools/call', {
        name: 'wiswork_call',
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
            name: 'wiswork_call',
            arguments: args,
          })
        ).status,
      ).toBe(200)
      expect(execute).toHaveBeenCalledOnce()
      expect(
        (
          await rpc(gateway.url, gateway.secret, 2, 'tools/call', {
            name: 'wiswork_call',
            arguments: args,
          })
        ).status,
      ).toBe(403)
      unregister()
      expect(
        (
          await rpc(gateway.url, gateway.secret, 3, 'tools/call', {
            name: 'wiswork_call',
            arguments: { ...args, callId: 'call-2' },
          })
        ).status,
      ).toBe(403)
      expect(execute).toHaveBeenCalledOnce()
    } finally {
      await gateway.close()
    }
  })

  it('keeps concurrent document grants isolated and allows only one carrier call per turn', async () => {
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
        name: 'wiswork_call',
        arguments: {
          capability: grants[0]!.capability,
          callId: 'a',
          toolName: 'read_blocks',
          input: {},
        },
      })
      const second = await rpc(gateway.url, gateway.secret, 2, 'tools/call', {
        name: 'wiswork_call',
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
        name: 'wiswork_call',
        arguments: {
          capability: grants[0]!.capability,
          callId: 'cross',
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
})
