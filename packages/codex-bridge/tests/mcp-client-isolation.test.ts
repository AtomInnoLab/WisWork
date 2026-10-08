import { expect, it, vi } from 'vitest'
import { startTrustedMcpTransport } from '../src/mcp-server.js'
import type { DocumentToolSession } from '../src/tool-router.js'

const initialize = {
  jsonrpc: '2.0',
  id: 0,
  method: 'initialize',
  params: {
    protocolVersion: '2025-06-18',
    capabilities: { elicitation: { form: {}, url: {} } },
    clientInfo: { name: 'codex-mcp-client', title: 'Codex', version: '0.147.0' },
  },
}

it('isolates trusted client initialization, replay budgets and explicit cleanup', async () => {
  const close = vi.fn()
  const session = {
    credentials: { sessionId: 'S'.repeat(43), secret: 'secret' },
    authorize: () => undefined,
    listTools: () => [],
    close,
  } as unknown as DocumentToolSession
  const server = await startTrustedMcpTransport(session, { maxActiveSessions: 2, maxRpcCalls: 3 })
  const request = (message: object, client?: string, method = 'POST') =>
    fetch(server.url, {
      method,
      headers: {
        authorization: 'Bearer secret',
        'content-type': 'application/json',
        ...(client === undefined ? {} : { 'Mcp-Session-Id': client }),
      },
      ...(method === 'POST' ? { body: JSON.stringify(message) } : {}),
    })
  const list = (id: number) => ({ jsonrpc: '2.0', id, method: 'tools/list' })
  try {
    const first = await request(initialize)
    const a = first.headers.get('mcp-session-id')!
    expect(a).toMatch(/^[A-Za-z0-9_-]{43}$/)
    const second = await request(initialize)
    const b = second.headers.get('mcp-session-id')!
    expect(b).toMatch(/^[A-Za-z0-9_-]{43}$/)
    expect(b).not.toBe(a)
    for (const client of [a, b])
      expect(
        (await request({ jsonrpc: '2.0', method: 'notifications/initialized' }, client)).status,
      ).toBe(202)
    expect((await request(initialize)).status).toBe(429)
    for (const client of [a, b])
      expect(await (await request(list(1), client)).json()).toHaveProperty('result.tools')
    expect(await (await request(list(1), a)).json()).toHaveProperty(
      'error.message',
      'request_id_consumed',
    )
    expect(await (await request(initialize, a)).json()).toHaveProperty(
      'error.message',
      'request_id_consumed',
    )
    expect((await request(list(2))).status).toBe(400)
    expect((await request(list(2), 'invalid')).status).toBe(400)
    expect((await request(list(2), `${a}, ${b}`)).status).toBe(400)
    expect((await request(list(2), 'X'.repeat(43))).status).toBe(404)
    await request(list(2), a)
    expect(await (await request(list(3), a)).json()).toHaveProperty(
      'error.message',
      'session_call_limit',
    )
    expect((await request(list(4), a)).status).toBe(404)
    expect(close).not.toHaveBeenCalled()
    expect(await (await request(list(2), b)).json()).toHaveProperty('result.tools')
    expect((await request({}, b, 'DELETE')).status).toBe(204)
    expect((await request(list(3), b)).status).toBe(404)
    expect((await request(initialize)).headers.get('mcp-session-id')).toMatch(/^[A-Za-z0-9_-]{43}$/)
  } finally {
    await server.close()
  }
  expect(close).toHaveBeenCalledOnce()
})

it('expires idle clients without evicting an in-flight call or resetting its replay state', async () => {
  let settle!: (value: { output: string; summary: string }) => void
  const callTool = vi.fn(
    () =>
      new Promise<{ output: string; summary: string }>((resolve) => {
        settle = resolve
      }),
  )
  const server = await startTrustedMcpTransport(
    {
      credentials: { sessionId: 'S'.repeat(43), secret: 'secret' },
      authorize: () => undefined,
      listTools: () => [],
      callTool,
      close: () => undefined,
    } as unknown as DocumentToolSession,
    { maxActiveSessions: 1 },
  )
  const request = (message: object, client?: string) =>
    fetch(server.url, {
      method: 'POST',
      headers: {
        authorization: 'Bearer secret',
        'content-type': 'application/json',
        ...(client === undefined ? {} : { 'Mcp-Session-Id': client }),
      },
      body: JSON.stringify(message),
    })
  const now = Date.now()
  const clock = vi.spyOn(Date, 'now').mockReturnValue(now)
  try {
    expect(await (await request({ ...initialize, params: {} })).json()).toHaveProperty(
      'error.message',
      'invalid_params',
    )
    const a = (await request(initialize)).headers.get('mcp-session-id')!
    await request({ jsonrpc: '2.0', method: 'notifications/initialized' }, a)
    const call = {
      jsonrpc: '2.0',
      id: 1,
      method: 'tools/call',
      params: { name: 'read', arguments: {} },
    }
    const pending = request(call, a)
    await vi.waitFor(() => expect(callTool).toHaveBeenCalledOnce())
    clock.mockReturnValue(now + 31 * 60_000)
    expect((await request(initialize)).status).toBe(429)
    settle({ output: 'ok', summary: 'ok' })
    expect((await pending).status).toBe(200)
    expect(await (await request(call, a)).json()).toHaveProperty(
      'error.message',
      'request_id_consumed',
    )
    clock.mockReturnValue(now + 62 * 60_000)
    expect((await request(initialize)).headers.get('mcp-session-id')).toMatch(/^[A-Za-z0-9_-]{43}$/)
    expect((await request(call, a)).status).toBe(404)
  } finally {
    clock.mockRestore()
    await server.close()
  }
})
