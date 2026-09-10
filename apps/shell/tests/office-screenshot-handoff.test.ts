import { afterEach, describe, expect, it, vi } from 'vitest'
import { ENHANCED_HOSTS } from '@wiswork/agent-runtime'
import { createOfficeCodexProxy } from '../src/main/office-codex-proxy'
import { createShellEnhancedPolicyAuthority } from '../src/main/enhanced-policy-authority'
import { createOfficeImageHandoff } from '../src/main/office-image-handoff'
import { PNG } from 'pngjs'
import { startTrustedMcpTransport } from '../../../packages/codex-bridge/src/mcp-server'
import { createOfficeAgentSession } from '../../office-addin/src/agent/use-office-agent'
import { createStructuredProposalController } from '../../office-addin/src/agent/proposal-controller'
import { prepareOfficeScreenshotPreview } from '../../office-addin/src/agent/office-screenshot-preview'
import { createPowerPointSkill } from '../../office-addin/src/skills/powerpoint/powerpoint-skill'
import type { PowerPointAdapter } from '../../office-addin/src/skills/powerpoint/browser-powerpoint-adapter'
import { createOfficeRelaySession, type RelayWebSocket } from '../../office-addin/src/relay/session'
import { createOfficeDiagnostics } from '../../office-addin/src/diagnostics/office-diagnostics'

afterEach(() => vi.unstubAllGlobals())

describe('Office screenshot delivery', () => {
  it('delivers the native screenshot through the agent, unchanged Relay v2, PC proxy, and MCP image content', async () => {
    const png = {
      mime: 'image/png' as const,
      base64:
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR4AWP4DwQACfsD/c8LaHIAAAAASUVORK5CYII=',
    }
    vi.stubGlobal('createImageBitmap', async () => ({ width: 1, height: 1, close() {} }))
    const sent: Record<string, any>[] = []
    const socket: RelayWebSocket = {
      readyState: 1,
      onopen: null,
      onclose: null,
      onerror: null,
      onmessage: null,
      send: (value) => sent.push(JSON.parse(value)),
      close: vi.fn(),
    }
    const receive = (frame: Record<string, unknown>) =>
      socket.onmessage?.({ data: JSON.stringify({ version: 2, ...frame }) })
    const bridge = createOfficeRelaySession({
      createSocket: () => socket,
      persistentPairing: false,
      capabilities: ['agent.v1'],
      randomUUID: () => 'request_12345678',
    })
    const connecting = bridge.connect('powerpoint')
    socket.onopen?.()
    receive({
      type: 'office.created',
      pairing_id: 'pair_12345678',
      verification_code: '123456',
      expires_in: 120,
    })
    receive({
      type: 'office.approved',
      session_id: 'session_12345678',
      capability: 'capability_12345678',
      capabilities: ['agent.v1'],
      expires_in: 1800,
    })
    await connecting
    const statement = {
      version: 1,
      runtime_mode: 'enhanced',
      runtime_instance: 'runtime_0123456789abcdef',
      component_version: '0.147.0',
      host: 'office-powerpoint',
      raw_office: false,
      expires_at: Date.now() + 60_000,
      policy_generation: 0,
      session_generation: 1,
    } as const
    receive({
      type: 'relay.session_state',
      session_id: 'session_12345678',
      generation: 1,
      enhanced: statement,
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({
      platform: 'Mac',
      proposals,
      prepareScreenshot: prepareOfficeScreenshotPreview,
      adapter: { screenshotSlide: async () => png } as unknown as PowerPointAdapter,
    })
    const diagnostics = createOfficeDiagnostics({ host: 'powerpoint', build: 'test' })
    const agent = createOfficeAgentSession({
      transport: { stream: () => ({ cancel() {} }) },
      skill,
      proposals,
      remoteTools: bridge,
      diagnostics,
    })
    const pending = bridge.capabilityFetch('agent.v1', {}).catch(() => undefined)
    let modelResult: any
    let runtimeError: unknown
    const proxy = createOfficeCodexProxy({
      prepareImageHandoff: createOfficeImageHandoff({
        createFromBuffer(bytes) {
          const decoded = PNG.sync.read(bytes)
          return {
            isEmpty: () => false,
            getSize: () => ({ width: decoded.width, height: decoded.height }),
            resize() {
              throw new Error('fitting preview must not resize')
            },
            toPNG() {
              throw new Error('fitting preview must not reencode')
            },
            toJPEG() {
              throw new Error('fitting preview must not reencode')
            },
          }
        },
      }),
      rollout: {
        globalEnabled: true,
        rawOfficeEnabled: false,
        hosts: Object.fromEntries(ENHANCED_HOSTS.map((host) => [host, true])) as any,
      },
      policyAuthority: createShellEnhancedPolicyAuthority(() => 0),
      runtime: {
        async runOfficeTurn(input: any) {
          const server = await startTrustedMcpTransport(input.toolSession)
          let client: string | undefined
          const rpc = async (message: object) => {
            const response = await fetch(server.url, {
              method: 'POST',
              headers: {
                authorization: `Bearer ${server.secret}`,
                'content-type': 'application/json',
                ...(client ? { 'mcp-session-id': client } : {}),
              },
              body: JSON.stringify(message),
            })
            client = response.headers.get('mcp-session-id') ?? client
            return response.status === 202 ? undefined : response.json()
          }
          try {
            await rpc({
              jsonrpc: '2.0',
              id: 1,
              method: 'initialize',
              params: {
                protocolVersion: '2025-06-18',
                capabilities: { elicitation: { form: {}, url: {} } },
                clientInfo: { name: 'codex-mcp-client', title: 'Codex', version: '0.147.0' },
              },
            })
            await rpc({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} })
            modelResult = await rpc({
              jsonrpc: '2.0',
              id: 2,
              method: 'tools/call',
              params: { name: 'screenshot_slide', arguments: { slide_index: 0 } },
            })
          } catch (error) {
            runtimeError = error
          } finally {
            await server.close()
            input.onEvent({ type: 'terminal', status: 'completed' })
          }
        },
      } as any,
    })
    try {
      const response = await proxy({
        body: JSON.parse(
          JSON.stringify({
            system: skill.systemPrompt,
            messages: [],
            tools: skill.tools.map((tool) => ({
              name: tool.name,
              description: tool.description,
              input_schema: tool.inputSchema,
            })),
          }),
        ),
        signal: new AbortController().signal,
        host: 'PowerPoint',
        sessionId: 'session_12345678',
        requestId: 'request_12345678',
        statement,
        executeTool: async (call) => {
          receive({
            type: 'relay.tool_call',
            session_id: 'session_12345678',
            request_id: 'request_12345678',
            turn_id: call.turnId,
            call_id: call.callId,
            generation: call.generation,
            tool_name: call.toolName,
            input: call.input,
          })
          await vi.waitFor(() =>
            expect(
              sent.find(
                (frame) => frame.type === 'office.tool_result' && frame.call_id === call.callId,
              ),
            ).toBeDefined(),
          )
          const frame = sent.find(
            (value) => value.type === 'office.tool_result' && value.call_id === call.callId,
          )!
          expect(Object.keys(frame).sort()).toEqual(
            [
              'version',
              'type',
              'session_id',
              'capability',
              'request_id',
              'turn_id',
              'call_id',
              'generation',
              'output',
              'is_error',
            ].sort(),
          )
          expect(Buffer.byteLength(JSON.stringify(frame))).toBeLessThan(256 * 1024)
          expect(JSON.parse(frame.output).visualAvailableToModel).toBe(false)
          return { output: frame.output, isError: frame.is_error }
        },
      })
      for await (const _chunk of response.body as AsyncIterable<Uint8Array>) {
        /* drain */
      }
      expect(runtimeError).toBeUndefined()
      expect(modelResult?.result).toMatchObject({
        isError: false,
        content: [
          { type: 'text', text: expect.stringContaining('"visualAvailableToModel":true') },
          { type: 'image', data: png.base64, mimeType: png.mime },
        ],
      })
      expect(modelResult.result.content[0].text).not.toContain(png.base64)
      expect(diagnostics.exportJson()).not.toContain(png.base64)
      expect(
        agent
          .snapshot()
          .timeline.some(
            (event) =>
              event.kind === 'tool' &&
              event.display?.items?.[0]?.url === `data:image/png;base64,${png.base64}`,
          ),
      ).toBe(true)
    } finally {
      agent.dispose()
      bridge.disconnect()
      await pending
    }
  })
})
