import { describe, expect, it, vi } from 'vitest'
import { startDynamicMcpGateway } from '../src/dynamic-mcp-gateway.js'

async function rpc(url: string, secret: string, id: number, method: string, params: object) {
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
})
