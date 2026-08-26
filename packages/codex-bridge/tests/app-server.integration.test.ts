import { access, mkdir, mkdtemp, rm } from 'node:fs/promises'
import { isAbsolute, join } from 'node:path'
import { tmpdir } from 'node:os'
import { describe, expect, it, vi } from 'vitest'
import { startResponsesBridge } from '../src/local-server.js'
import { startDocumentMcpServer } from '../src/mcp-server.js'
import { CodexProcessManager, type OwnedCodexDirectories } from '../src/process-manager.js'
import type { MessagesRequest } from '../src/types.js'

const executable = process.env.WISWORK_CODEX_INTEGRATION_EXECUTABLE
const realIt = executable !== undefined && isAbsolute(executable) ? it : it.skip

function anthropicTextStream(): ReadableStream<Uint8Array> {
  const events = [
    {
      type: 'message_start',
      message: {
        id: 'response_integration_1',
        type: 'message',
        role: 'assistant',
        content: [],
        model: 'openai/gpt-5.6-sol',
        stop_reason: null,
        stop_sequence: null,
        usage: { input_tokens: 1, output_tokens: 0 },
      },
    },
    { type: 'content_block_start', index: 0, content_block: { type: 'text', text: '' } },
    { type: 'content_block_delta', index: 0, delta: { type: 'text_delta', text: 'bridge-ok' } },
    { type: 'content_block_stop', index: 0 },
    {
      type: 'message_delta',
      delta: { stop_reason: 'end_turn', stop_sequence: null },
      usage: { output_tokens: 1 },
    },
    { type: 'message_stop' },
  ]
  const encoder = new TextEncoder()
  return new ReadableStream({
    start(controller) {
      for (const event of events) {
        controller.enqueue(
          encoder.encode(`event: ${event.type}\ndata: ${JSON.stringify(event)}\n\n`),
        )
      }
      controller.close()
    },
  })
}

function anthropicToolStream(): ReadableStream<Uint8Array> {
  const events = [
    {
      type: 'message_start',
      message: {
        id: 'response_integration_tool_1',
        type: 'message',
        role: 'assistant',
        content: [],
        model: 'openai/gpt-5.6-sol',
        stop_reason: null,
        stop_sequence: null,
        usage: { input_tokens: 1, output_tokens: 0 },
      },
    },
    {
      type: 'content_block_start',
      index: 0,
      content_block: { type: 'tool_use', id: 'custom_integration_1', name: 'exec', input: {} },
    },
    {
      type: 'content_block_delta',
      index: 0,
      delta: {
        type: 'input_json_delta',
        partial_json: '{"code":"text(await tools.mcp__wiswork__read_document({}))"}',
      },
    },
    { type: 'content_block_stop', index: 0 },
    {
      type: 'message_delta',
      delta: { stop_reason: 'tool_use', stop_sequence: null },
      usage: { output_tokens: 1 },
    },
    { type: 'message_stop' },
  ]
  const encoder = new TextEncoder()
  return new ReadableStream({
    start(controller) {
      for (const event of events) {
        controller.enqueue(
          encoder.encode(`event: ${event.type}\ndata: ${JSON.stringify(event)}\n\n`),
        )
      }
      controller.close()
    },
  })
}

describe('real Codex 0.147.0 app-server contract', () => {
  realIt(
    'runs an authenticated text turn through the local bridge and cleans owned state',
    async () => {
      const requests: MessagesRequest[] = []
      const bridge = await startResponsesBridge({
        fetchWithAuth: async (request) => {
          requests.push(request)
          return new Response(anthropicTextStream(), {
            status: 200,
            headers: { 'content-type': 'text/event-stream' },
          })
        },
      })
      const root = await mkdtemp(join(tmpdir(), 'wiswork-codex-integration-'))
      const directories: OwnedCodexDirectories = {
        root,
        codexHome: join(root, 'home'),
        cwd: join(root, 'workspace'),
      }
      await mkdir(directories.codexHome)
      await mkdir(directories.cwd)
      const removeDirectories = vi.fn(async () => rm(root, { recursive: true, force: true }))
      const diagnostics: string[] = []
      const manager = new CodexProcessManager({
        executablePath: executable!,
        bridge,
        developerInstructions: 'You are the fixed WisWork document assistant.',
        createDirectories: async () => directories,
        removeDirectories,
        diagnostics: ({ code }) => diagnostics.push(code),
      })
      try {
        const client = await manager.start()
        await client.initialize()
        const thread = await client.startThread()
        let text = ''
        const completed = new Promise<string>((resolve, reject) => {
          const timer = setTimeout(
            () => reject(new Error('integration_notification_timeout')),
            20_000,
          )
          timer.unref()
          const unsubscribe = client.onNotification((notification) => {
            if (notification.method === 'item/agentMessage/delta') {
              text += notification.params.delta
            }
            if (notification.method === 'turn/completed') {
              clearTimeout(timer)
              unsubscribe()
              resolve(notification.params.turn.id)
            }
          })
        })
        const turn = await client.startTurn(thread.thread.id, 'Say bridge-ok.')
        await expect(completed).resolves.toBe(turn.turn.id)
        expect(text).toBe('bridge-ok')
        expect(requests).toHaveLength(1)
        expect(requests[0]).toMatchObject({ model: 'openai/gpt-5.6-sol', stream: true })
      } finally {
        await manager.stop()
        await bridge.close()
      }
      expect(removeDirectories).toHaveBeenCalledOnce()
      await expect(access(root)).rejects.toMatchObject({ code: 'ENOENT' })
      expect(diagnostics).not.toContain('codex_process_stop_timeout')
    },
    30_000,
  )

  realIt(
    'discovers and calls only the authenticated document MCP session',
    async () => {
      const executeTool = vi.fn(() => ({
        output: 'document-via-mcp',
        summary: 'Read document',
      }))
      const mcp = await startDocumentMcpServer()
      const mcpSession = mcp.register({
        skill: {
          id: 'docs',
          systemPrompt: '',
          tools: [
            {
              name: 'read_document',
              description: 'Read the current document.',
              inputSchema: { type: 'object', additionalProperties: false },
            },
          ],
          executeTool,
        },
        policy: { read_document: 'read' },
        isOpen: () => true,
        getRevision: () => 'rev-1',
        requestApproval: () => false,
        captureSnapshot: () => 'unused',
      })
      const requests: MessagesRequest[] = []
      const bridge = await startResponsesBridge({
        fetchWithAuth: async (request) => {
          requests.push(request)
          return new Response(
            requests.length === 1 ? anthropicToolStream() : anthropicTextStream(),
            {
              status: 200,
              headers: { 'content-type': 'text/event-stream' },
            },
          )
        },
      })
      const root = await mkdtemp(join(tmpdir(), 'wiswork-codex-integration-'))
      const directories: OwnedCodexDirectories = {
        root,
        codexHome: join(root, 'home'),
        cwd: join(root, 'workspace'),
      }
      await mkdir(directories.codexHome)
      await mkdir(directories.cwd)
      const manager = new CodexProcessManager({
        executablePath: executable!,
        bridge,
        mcp: { url: mcpSession.url, secret: mcpSession.secret },
        developerInstructions: 'Use the document read tool once, then answer.',
        createDirectories: async () => directories,
        removeDirectories: async () => rm(root, { recursive: true, force: true }),
      })
      try {
        const client = await manager.start()
        await client.initialize()
        const thread = await client.startThread()
        const notifications: unknown[] = []
        const completed = new Promise<void>((resolve, reject) => {
          const timer = setTimeout(
            () => reject(new Error('integration_notification_timeout')),
            20_000,
          )
          timer.unref()
          const unsubscribe = client.onNotification((notification) => {
            notifications.push(notification)
            if (notification.method === 'turn/completed') {
              clearTimeout(timer)
              unsubscribe()
              resolve()
            }
          })
        })
        await client.startTurn(thread.thread.id, 'Read the document and answer.')
        await completed
        expect(executeTool).toHaveBeenCalledOnce()
        expect(executeTool).toHaveBeenCalledWith(
          expect.objectContaining({ name: 'read_document', input: {} }),
          expect.any(AbortSignal),
        )
        expect(requests).toHaveLength(1)
        expect(requests[0]?.tools?.map((tool) => tool.name)).toEqual(['exec'])
        expect(JSON.stringify(notifications)).toContain('mcpToolCall')
        expect(JSON.stringify(notifications)).toContain('document-via-mcp')
      } finally {
        await manager.stop()
        mcpSession.close()
        await mcp.close()
        await bridge.close()
      }
      await expect(access(root)).rejects.toMatchObject({ code: 'ENOENT' })
    },
    30_000,
  )
})
