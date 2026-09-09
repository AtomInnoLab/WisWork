import { describe, expect, it, vi } from 'vitest'
import { ENHANCED_HOSTS, type EnhancedRolloutPolicy } from '@wiswork/agent-runtime'
import { startDynamicMcpGateway } from '@wiswork/codex-bridge'
import { createOfficeCodexProxy } from '../src/main/office-codex-proxy'
import { createShellEnhancedPolicyAuthority } from '../src/main/enhanced-policy-authority'
import { createPcBridgeAgentTransport } from '../../office-addin/src/agent/transport'

const rollout: EnhancedRolloutPolicy = {
  globalEnabled: true,
  rawOfficeEnabled: false,
  hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as any,
}

// Exercise the actual authenticated gateway, document router, Office SSE projection and parser.
async function runOfficeTools(options: {
  names: string[]
  run(call: (id: string, name: string, mutate?: boolean) => Promise<any>): Promise<void>
  executeTool?: (call: any) => Promise<{ output: string; isError?: boolean }>
}) {
  const gateway = await startDynamicMcpGateway()
  const executeTool = vi.fn(
    options.executeTool ?? (async () => ({ output: 'document', isError: false })),
  )
  const executeRetrieval = vi.fn(async () =>
    new TextEncoder().encode(
      JSON.stringify({
        results: [
          {
            title: 'Source',
            url: 'https://example.com/article?token=private-secret',
            content: 'private-body',
          },
        ],
      }),
    ),
  )
  const activities = vi.fn(),
    onError = vi.fn(),
    onToolCall = vi.fn(),
    onDone = vi.fn()
  const failures: unknown[] = []
  const headers: Record<string, string> = {
    authorization: `Bearer ${gateway.secret}`,
    'content-type': 'application/json',
  }
  const rpc = async (method: string, params: object, id?: string) => {
    const response = await fetch(gateway.url, {
      method: 'POST',
      headers,
      body: JSON.stringify({ jsonrpc: '2.0', ...(id ? { id } : {}), method, params }),
    })
    if (method === 'initialize') headers['Mcp-Session-Id'] = response.headers.get('mcp-session-id')!
    return response
  }
  try {
    await rpc(
      'initialize',
      {
        protocolVersion: '2025-06-18',
        capabilities: { elicitation: { form: {}, url: {} } },
        clientInfo: { name: 'codex-mcp-client', title: 'Codex', version: '0.147.0' },
      },
      'init',
    )
    await rpc('notifications/initialized', {})
    const proxy = createOfficeCodexProxy({
      runtime: {
        async runOfficeTurn(input: any) {
          const unregister = gateway.register({
            ownerId: input.toolSession.identity.ownerId,
            documentId: input.documentId,
            generation: input.generation,
            session: input.toolSession,
            onToolEvent: input.onEvent,
            summarizeProposal: input.summarizeProposal,
            onProposal() {},
          })
          const { capability } = gateway.beginTurn({
            documentId: input.documentId,
            generation: input.generation,
            threadId: 'thread',
          })
          try {
            await options.run(async (id, name, mutate = false) => {
              const response = await rpc(
                'tools/call',
                {
                  name: mutate ? 'wiswork_propose' : 'wiswork_read',
                  arguments: {
                    capability,
                    callId: id,
                    toolName: name,
                    input: { query: 'public query' },
                  },
                },
                id,
              )
              expect(response.status).toBe(200)
              return response.json()
            })
            input.onEvent({ type: 'terminal', status: 'completed' })
          } catch (error) {
            failures.push(error)
            throw error
          } finally {
            gateway.revokeTurn(capability)
            unregister()
          }
        },
      } as any,
      rollout,
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
    })
    const transport = createPcBridgeAgentTransport({
      snapshot: () => ({ enhanced: { session_generation: 3 } }),
      authenticatedFetch: async (_path, init) => {
        const response = await proxy({
          body: JSON.parse(init.body as string),
          signal: init.signal!,
          host: 'PowerPoint',
          sessionId: 'session_12345678',
          requestId: 'request_12345678',
          statement: {
            version: 1,
            runtime_mode: 'enhanced',
            runtime_instance: 'runtime_0123456789abcdef',
            component_version: '0.147.0',
            host: 'office-powerpoint',
            raw_office: false,
            expires_at: Date.now() + 60_000,
            policy_generation: 0,
            session_generation: 3,
          },
          executeTool,
          executeRetrieval,
        })
        const iterator = (response.body as AsyncIterable<Uint8Array>)[Symbol.asyncIterator]()
        return new Response(
          new ReadableStream<Uint8Array>({
            async pull(controller) {
              try {
                const next = await iterator.next()
                if (next.done) controller.close()
                else controller.enqueue(next.value)
              } catch (error) {
                controller.error(error)
              }
            },
            async cancel() {
              await iterator.return?.()
            },
          }),
        )
      },
    })
    transport.setToolActivityHandler?.(activities)
    transport.stream(
      {
        system: '',
        messages: [],
        tools: options.names.map((name) => ({
          name,
          description: name,
          inputSchema: { type: 'object' },
        })),
      },
      { onDelta() {}, onToolCall, onError, onDone },
    )
    await vi.waitFor(() => expect(onDone).toHaveBeenCalledOnce(), { timeout: 10_000 })
    expect(failures).toEqual([])
    return {
      events: activities.mock.calls.map(([event]) => event),
      onError,
      onToolCall,
      executeTool,
      executeRetrieval,
    }
  } finally {
    await gateway.close()
  }
}

describe('Office semantic lifecycle delivery', () => {
  it('completes 33 real gateway retrievals with one start and terminal per call', async () => {
    const result = await runOfficeTools({
      names: ['web_search'],
      async run(call) {
        for (let index = 0; index < 33; index++) {
          const receipt = await call(`search_${index}`, 'web_search')
          expect(receipt.result.isError).not.toBe(true)
        }
      },
    })
    expect(result.onError).not.toHaveBeenCalled()
    expect(result.executeRetrieval).toHaveBeenCalledTimes(33)
    expect(result.events).toHaveLength(66)
    expect(new Set(result.events.map((event) => event.callId)).size).toBe(33)
    for (let index = 0; index < 66; index += 2) {
      expect(result.events[index]).toMatchObject({ state: 'running' })
      expect(result.events[index + 1]).toMatchObject({
        callId: result.events[index].callId,
        state: 'complete',
        display: {
          kind: 'links',
          items: [{ title: 'Source', url: 'https://example.com/article' }],
        },
      })
    }
    expect(result.executeTool).not.toHaveBeenCalled()
    expect(result.onToolCall).not.toHaveBeenCalled()
    expect(JSON.stringify(result.events)).not.toMatch(/private-secret|private-body/)
  })

  it('shows a read rejected by a pending write without dispatching the read', async () => {
    let writeStarted = false
    let finish!: () => void
    const pending = new Promise<void>((resolve) => {
      finish = resolve
    })
    const result = await runOfficeTools({
      names: ['get_presentation_state', 'edit_slide_text'],
      executeTool: async (call) => {
        expect(call.toolName).toBe('edit_slide_text')
        writeStarted = true
        await pending
        return { output: 'write receipt', isError: false }
      },
      async run(call) {
        const write = call('write_1', 'edit_slide_text', true)
        try {
          await vi.waitFor(() => expect(writeStarted).toBe(true))
          const receipt = await call('read_1', 'get_presentation_state')
          expect(JSON.stringify(receipt)).toContain('tool_call_in_progress')
        } finally {
          finish()
          await write
        }
      },
    })
    expect(result.onError).not.toHaveBeenCalled()
    expect(result.executeTool).toHaveBeenCalledOnce()
    expect(result.events.filter((event) => event.toolName === 'edit_slide_text')).toEqual([
      expect.objectContaining({
        callId: result.executeTool.mock.calls[0][0].callId,
        state: 'running',
      }),
      expect.objectContaining({
        callId: result.executeTool.mock.calls[0][0].callId,
        state: 'complete',
      }),
    ])
    expect(
      result.events
        .filter((event) => event.toolName === 'get_presentation_state')
        .map((event) => event.state),
    ).toEqual(['running', 'error'])
    expect(result.onToolCall).not.toHaveBeenCalled()
  })
})
