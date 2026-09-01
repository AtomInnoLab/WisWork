import { request } from 'node:http'
import { suspendToolExecution } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { startDocumentMcpServer } from '../src/mcp-server.js'
import { createDocumentToolManifest } from '../src/tool-router.js'

const rollout = {
  globalEnabled: true,
  rawOfficeEnabled: false,
  hosts: {
    latex: true,
    slides: true,
    docs: true,
    sheets: true,
    'office-word': true,
    'office-excel': true,
    'office-powerpoint': true,
  },
}
const tool = { name: 'get_document_context', description: 'Read.', inputSchema: { type: 'object' } }
const policyGrant = () => {
  const grant = Object.freeze({})
  const snapshot = {
    generation: 1,
    host: 'docs',
    policy: rollout,
    capabilities: ['semantic-read'],
  } as const
  return {
    policyGrant: grant,
    consumePolicyGrant: (candidate: unknown) => {
      if (candidate !== grant) throw new Error('invalid_enhanced_policy_handle')
      return snapshot
    },
  }
}
const registration = () => ({
  identity: { ownerId: 'o', host: 'docs' as const, documentId: 'd', sessionId: 's', generation: 1 },
  manifest: createDocumentToolManifest({
    ...policyGrant(),
    tools: [tool],
    policy: { get_document_context: 'read' },
  }),
  isOpen: () => true,
  executeRead: async () => ({ output: 'document', summary: 'read' }),
  suspendMutation: suspendToolExecution,
})
function post(url: string, secret: string, value: unknown, authorization = `Bearer ${secret}`) {
  const data = JSON.stringify(value)
  return new Promise<{ status: number; json: any }>((resolve, reject) => {
    const req = request(
      url,
      {
        method: 'POST',
        headers: {
          authorization,
          'content-type': 'application/json',
          'content-length': String(Buffer.byteLength(data)),
        },
      },
      (response) => {
        const chunks: Buffer[] = []
        response.on('data', (c) => chunks.push(Buffer.from(c)))
        response.on('end', () => {
          const text = Buffer.concat(chunks).toString()
          resolve({ status: response.statusCode ?? 0, json: text ? JSON.parse(text) : undefined })
        })
      },
    )
    req.on('error', reject)
    req.end(data)
  })
}
function methodStatus(url: string, method: string, authorization?: string) {
  return new Promise<number>((resolve, reject) => {
    const req = request(
      url,
      { method, headers: authorization ? { authorization } : {} },
      (response) => {
        response.resume()
        response.on('end', () => resolve(response.statusCode ?? 0))
      },
    )
    req.on('error', reject)
    req.end()
  })
}
async function initialize(url: string, secret: string) {
  const result = await post(url, secret, {
    jsonrpc: '2.0',
    id: 'init',
    method: 'initialize',
    params: {
      protocolVersion: '2025-06-18',
      capabilities: { elicitation: { form: {}, url: {} } },
      clientInfo: { name: 'codex-mcp-client', title: 'Codex', version: '0.147.0' },
    },
  })
  expect(result.json.result.protocolVersion).toBe('2025-06-18')
  expect(
    (await post(url, secret, { jsonrpc: '2.0', method: 'notifications/initialized', params: {} }))
      .status,
  ).toBe(202)
}

describe('document MCP server', () => {
  it('authenticates first, enforces exact state order and isolates sessions', async () => {
    const server = await startDocumentMcpServer()
    const a = server.register(registration())
    const b = server.register(registration())
    try {
      expect(
        (await post(a.url, b.secret, { jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} }))
          .status,
      ).toBe(401)
      expect(
        (await post(a.url, a.secret, { jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} }))
          .json.error.message,
      ).toBe('not_initialized')
      await initialize(a.url, a.secret)
      expect(
        (
          await post(a.url, a.secret, {
            jsonrpc: '2.0',
            id: 'again',
            method: 'initialize',
            params: {},
          })
        ).json.error,
      ).toBeDefined()
      expect(
        (
          await post(a.url, a.secret, {
            jsonrpc: '2.0',
            method: 'notifications/initialized',
            params: {},
          })
        ).json.error,
      ).toBeDefined()
      const listed = await post(a.url, a.secret, {
        jsonrpc: '2.0',
        id: 2,
        method: 'tools/list',
        params: {},
      })
      expect(listed.json.result.tools[0].name).toBe(tool.name)
      const called = await post(a.url, a.secret, {
        jsonrpc: '2.0',
        id: 3,
        method: 'tools/call',
        params: { name: tool.name, arguments: {} },
      })
      expect(called.json.result.content[0].text).toBe('document')
      expect(
        (
          await post(a.url, a.secret, {
            jsonrpc: '2.0',
            id: 3,
            method: 'tools/call',
            params: { name: tool.name, arguments: {} },
          })
        ).json.error.message,
      ).toBe('request_id_consumed')
    } finally {
      await server.close()
    }
  })

  it('requires canonical bearer and redacts unknown method input', async () => {
    const diagnostics = vi.fn()
    const server = await startDocumentMcpServer({ diagnostics })
    const session = server.register(registration())
    try {
      expect(
        (
          await post(
            session.url,
            session.secret,
            { jsonrpc: '2.0', id: 1, method: 'initialize', params: {} },
            `Bearer ${session.secret}=`,
          )
        ).status,
      ).toBe(401)
      await initialize(session.url, session.secret)
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
      await server.close()
    }
  })

  it('closes instead of evicting consumed RPC ids at the total-call bound', async () => {
    const server = await startDocumentMcpServer({ maxRpcCalls: 2 })
    const session = server.register(registration())
    try {
      await initialize(session.url, session.secret)
      await post(session.url, session.secret, {
        jsonrpc: '2.0',
        id: 'last',
        method: 'tools/list',
        params: {},
      })
      const limited = await post(session.url, session.secret, {
        jsonrpc: '2.0',
        id: 'over',
        method: 'tools/list',
        params: {},
      })
      expect(limited.json.error.message).toBe('session_call_limit')
      expect(
        (
          await post(session.url, session.secret, {
            jsonrpc: '2.0',
            id: 'init',
            method: 'initialize',
            params: {},
          })
        ).status,
      ).toBe(401)
    } finally {
      await server.close()
    }
  })

  it('authenticates path/session/bearer before method dispatch', async () => {
    const server = await startDocumentMcpServer()
    const session = server.register(registration())
    try {
      expect(await methodStatus(session.url, 'GET')).toBe(401)
      expect(await methodStatus(`${server.baseUrl}/unknown`, 'GET', 'Bearer invalid')).toBe(401)
      expect(await methodStatus(session.url, 'GET', `Bearer ${session.secret}`)).toBe(405)
    } finally {
      await server.close()
    }
  })

  it('enforces the compiled active-session cap and releases on close', async () => {
    const server = await startDocumentMcpServer({ maxActiveSessions: 1 })
    const first = server.register(registration())
    try {
      expect(() => server.register(registration())).toThrow('mcp_session_limit')
      first.close()
      expect(() => server.register(registration())).not.toThrow()
    } finally {
      await server.close()
    }
  })
})
