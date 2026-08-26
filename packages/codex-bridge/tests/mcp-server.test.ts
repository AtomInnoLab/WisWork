import type { AgentSkill } from '@wiswork/agent-core'
import { request as httpRequest } from 'node:http'
import { describe, expect, it, vi } from 'vitest'
import { startDocumentMcpServer } from '../src/mcp-server.js'
import type { ToolSessionRegistration } from '../src/tool-router.js'

function registration(
  name: string,
  executeTool?: AgentSkill['executeTool'],
): ToolSessionRegistration {
  const toolName = `${name}_read`
  return {
    skill: {
      id: name,
      systemPrompt: '',
      tools: [
        {
          name: toolName,
          description: `Read ${name}.`,
          inputSchema: { type: 'object', additionalProperties: false },
        },
      ],
      executeTool:
        executeTool ??
        (() => ({ output: `${name} private text`, summary: `Read ${name}`, mutated: false })),
    },
    policy: { [toolName]: 'read' },
    isOpen: () => true,
    getRevision: () => '1',
    requestApproval: () => false,
    captureSnapshot: () => 'unused',
  }
}

async function rpc(
  url: string,
  token: string,
  body: unknown,
): Promise<{ status: number; json?: Record<string, unknown> }> {
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      authorization: `Bearer ${token}`,
      'content-type': 'application/json',
      accept: 'application/json, text/event-stream',
    },
    body: JSON.stringify(body),
  })
  return {
    status: response.status,
    ...(response.status === 202
      ? {}
      : { json: (await response.json()) as Record<string, unknown> }),
  }
}

describe('document MCP server', () => {
  it('binds loopback and authenticates each opaque document endpoint independently', async () => {
    const server = await startDocumentMcpServer()
    try {
      const docs = server.register(registration('docs'))
      const sheets = server.register(registration('sheets'))
      expect(new URL(docs.url).hostname).toBe('127.0.0.1')
      expect(docs.url).not.toBe(sheets.url)
      expect(docs.secret).not.toBe(sheets.secret)

      const initialized = await rpc(docs.url, docs.secret, {
        jsonrpc: '2.0',
        id: 0,
        method: 'initialize',
        params: {
          protocolVersion: '2025-06-18',
          capabilities: { elicitation: { form: {}, url: {} } },
          clientInfo: { name: 'codex-mcp-client', title: 'Codex', version: '0.147.0' },
        },
      })
      expect(initialized).toEqual({
        status: 200,
        json: {
          jsonrpc: '2.0',
          id: 0,
          result: {
            protocolVersion: '2025-06-18',
            capabilities: { tools: {} },
            serverInfo: { name: 'wiswork', version: '0.1.0' },
          },
        },
      })

      await expect(
        rpc(docs.url, sheets.secret, { jsonrpc: '2.0', id: 1, method: 'tools/list' }),
      ).resolves.toMatchObject({ status: 401 })
      await expect(
        rpc(sheets.url, docs.secret, { jsonrpc: '2.0', id: 1, method: 'tools/list' }),
      ).resolves.toMatchObject({ status: 401 })
      await expect(
        rpc(docs.url, sheets.secret, {
          jsonrpc: '2.0',
          method: 'notifications/initialized',
          params: {},
        }),
      ).resolves.toMatchObject({ status: 401 })
      await expect(
        rpc(docs.url, sheets.secret, { jsonrpc: '2.0', id: 1, method: 'unknown/private' }),
      ).resolves.toMatchObject({ status: 401 })

      const listed = await rpc(docs.url, docs.secret, {
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/list',
        params: { _meta: { progressToken: 0 } },
      })
      expect(listed.json).toMatchObject({
        result: {
          tools: [{ name: 'docs_read', annotations: { readOnlyHint: true } }],
        },
      })
      expect(JSON.stringify(listed)).not.toContain('sheets_read')

      const wrongVersion = await rpc(docs.url, docs.secret, {
        jsonrpc: '2.0',
        id: 2,
        method: 'initialize',
        params: {
          protocolVersion: '2025-06-18',
          capabilities: { elicitation: { form: {}, url: {} } },
          clientInfo: { name: 'codex-mcp-client', title: 'Codex', version: '0.148.0' },
        },
      })
      expect(wrongVersion.json).toMatchObject({
        error: { code: -32602, message: 'invalid_params' },
      })
    } finally {
      await server.close()
    }
  })

  it('maps tools/call to MCP content without leaking side-channel display data', async () => {
    const server = await startDocumentMcpServer()
    try {
      const session = server.register({
        ...registration('docs'),
        skill: {
          ...registration('docs').skill,
          executeTool: () => ({
            output: 'safe result',
            summary: 'Read',
            display: { kind: 'text', text: 'renderer-only private display' },
          }),
        },
      })
      const response = await rpc(session.url, session.secret, {
        jsonrpc: '2.0',
        id: 'call-1',
        method: 'tools/call',
        params: { name: 'docs_read', arguments: {}, _meta: { progressToken: 1 } },
      })
      expect(response).toEqual({
        status: 200,
        json: {
          jsonrpc: '2.0',
          id: 'call-1',
          result: { content: [{ type: 'text', text: 'safe result' }], isError: false },
        },
      })
      expect(JSON.stringify(response)).not.toContain('renderer-only')
    } finally {
      await server.close()
    }
  })

  it('forwards request cancellation and aborts calls during session/server teardown', async () => {
    let abortCount = 0
    let notifyStarted = (): void => undefined
    const started = (): Promise<void> =>
      new Promise((resolve) => {
        notifyStarted = resolve
      })
    let callStarted = started()
    const slow = registration(
      'slow',
      (_call, signal) =>
        new Promise((resolve) => {
          notifyStarted()
          signal?.addEventListener(
            'abort',
            () => {
              abortCount += 1
              resolve({ output: 'cancelled', summary: 'Cancelled', isError: true })
            },
            { once: true },
          )
        }),
    )
    const server = await startDocumentMcpServer()
    const session = server.register(slow)
    const call = rpc(session.url, session.secret, {
      jsonrpc: '2.0',
      id: 9,
      method: 'tools/call',
      params: { name: 'slow_read', arguments: {} },
    })
    await callStarted
    await expect(
      rpc(session.url, session.secret, {
        jsonrpc: '2.0',
        method: 'notifications/cancelled',
        params: { requestId: 9, reason: 'private user text' },
      }),
    ).resolves.toEqual({ status: 202 })
    await expect(call).resolves.toMatchObject({
      json: { result: { content: [{ text: 'tool_cancelled' }], isError: true } },
    })
    expect(abortCount).toBe(1)

    callStarted = started()
    const next = rpc(session.url, session.secret, {
      jsonrpc: '2.0',
      id: 10,
      method: 'tools/call',
      params: { name: 'slow_read', arguments: {} },
    })
    await callStarted
    session.close()
    await expect(next).resolves.toMatchObject({
      json: { result: { content: [{ text: 'tool_cancelled' }], isError: true } },
    })
    await server.close()
    expect(abortCount).toBe(2)
    await expect(fetch(session.url)).rejects.toThrow()
  })

  it('aborts renderer work when the authenticated MCP client disconnects', async () => {
    let notifyStarted = (): void => undefined
    const started = new Promise<void>((resolve) => {
      notifyStarted = resolve
    })
    let resolveAborted = (_value: 'aborted'): void => undefined
    const aborted = new Promise<'aborted'>((resolve) => {
      resolveAborted = resolve
    })
    let finishTool = (): void => undefined
    const server = await startDocumentMcpServer()
    const session = server.register(
      registration(
        'disconnect',
        (_call, signal) =>
          new Promise((resolve) => {
            notifyStarted()
            const finish = (): void => {
              resolveAborted('aborted')
              resolve({ output: 'cancelled', summary: 'Cancelled', isError: true })
            }
            signal?.addEventListener('abort', finish, { once: true })
            finishTool = finish
          }),
      ),
    )
    const target = new URL(session.url)
    const request = httpRequest({
      hostname: target.hostname,
      port: target.port,
      path: target.pathname,
      method: 'POST',
      headers: {
        authorization: `Bearer ${session.secret}`,
        'content-type': 'application/json',
      },
    })
    request.on('error', () => undefined)
    request.end(
      JSON.stringify({
        jsonrpc: '2.0',
        id: 11,
        method: 'tools/call',
        params: { name: 'disconnect_read', arguments: {} },
      }),
    )
    await started
    request.destroy()
    let outcome: 'aborted' | 'timeout'
    try {
      outcome = await Promise.race([
        aborted,
        new Promise<'timeout'>((resolve) => setTimeout(() => resolve('timeout'), 50)),
      ])
    } finally {
      finishTool()
      session.close()
      await server.close()
    }
    expect(outcome).toBe('aborted')
  })

  it('fails closed on malformed, oversized, and unknown protocol input with redacted diagnostics', async () => {
    const diagnostics = vi.fn()
    const server = await startDocumentMcpServer({ maxBodyBytes: 128, diagnostics })
    try {
      const session = server.register(registration('docs'))
      const oversized = await fetch(session.url, {
        method: 'POST',
        headers: {
          authorization: `Bearer ${session.secret}`,
          'content-type': 'application/json',
        },
        body: JSON.stringify({ privatePrompt: 'x'.repeat(200) }),
      })
      expect(oversized.status).toBe(413)
      const unknown = await rpc(session.url, session.secret, {
        jsonrpc: '2.0',
        id: 3,
        method: 'resources/list',
      })
      expect(unknown.json).toMatchObject({ error: { code: -32601, message: 'method_not_found' } })
      expect(diagnostics).toHaveBeenCalledWith({ code: 'mcp_body_limit' })
      expect(JSON.stringify(diagnostics.mock.calls)).not.toContain('privatePrompt')
      expect(JSON.stringify(diagnostics.mock.calls)).not.toContain(session.secret)
    } finally {
      await server.close()
    }
  })
})
