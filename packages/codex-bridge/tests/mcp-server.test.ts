import { request } from 'node:http'
import type { AgentSkill } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { startDocumentMcpServer } from '../src/mcp-server.js'

const skill: AgentSkill = {
  id: 'docs',
  systemPrompt: '',
  tools: [{ name: 'read_document', description: 'Read.', inputSchema: { type: 'object' } }],
  executeTool: async () => ({ output: 'document', summary: 'read' }),
}
const registration = () => ({
  identity: { ownerId: 'o', host: 'docs' as const, documentId: 'd', sessionId: 's', generation: 1 },
  skill,
  policy: { read_document: 'read' as const },
  isOpen: () => true,
})
function post(url: string, secret: string, value: unknown) {
  const data = JSON.stringify(value)
  return new Promise<{ status: number; json: any }>((resolve, reject) => {
    const req = request(
      url,
      {
        method: 'POST',
        headers: {
          authorization: `Bearer ${secret}`,
          'content-type': 'application/json',
          'content-length': String(Buffer.byteLength(data)),
        },
      },
      (response) => {
        const chunks: Buffer[] = []
        response.on('data', (c) => chunks.push(Buffer.from(c)))
        response.on('end', () =>
          resolve({
            status: response.statusCode ?? 0,
            json: JSON.parse(Buffer.concat(chunks).toString()),
          }),
        )
      },
    )
    req.on('error', reject)
    req.end(data)
  })
}
describe('document MCP server', () => {
  it('isolates sessions, lists bounded tools and calls reads', async () => {
    const server = await startDocumentMcpServer()
    const a = server.register(registration())
    const b = server.register(registration())
    try {
      expect(
        (await post(a.url, b.secret, { jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} }))
          .status,
      ).toBe(401)
      const listed = await post(a.url, a.secret, {
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/list',
        params: {},
      })
      expect(listed.json.result.tools[0].name).toBe('read_document')
      const called = await post(a.url, a.secret, {
        jsonrpc: '2.0',
        id: 2,
        method: 'tools/call',
        params: { name: 'read_document', arguments: {} },
      })
      expect(called.json.result.content[0].text).toBe('document')
    } finally {
      await server.close()
    }
  })
  it('rejects unknown methods and redacts diagnostics', async () => {
    const diagnostics = vi.fn()
    const server = await startDocumentMcpServer({ diagnostics })
    const session = server.register(registration())
    try {
      const result = await post(session.url, session.secret, {
        jsonrpc: '2.0',
        id: 'x',
        method: 'shell',
        params: { secret: 'sensitive' },
      })
      expect(result.json.error.message).toBe('method_not_found')
      expect(JSON.stringify(result)).not.toContain('sensitive')
    } finally {
      await server.close()
    }
  })
})
