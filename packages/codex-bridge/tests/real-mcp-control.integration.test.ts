import { isAbsolute } from 'node:path'
import { createToolExecutionSuspensionAuthority } from '@wiswork/agent-core'
import { describe, expect, it, vi } from 'vitest'
import { startDocumentMcpServer } from '../src/mcp-server.js'
import { CodexProcessManager } from '../src/process-manager.js'
import { createDocumentToolManifest } from '../src/tool-router.js'

const executable = process.env.WISWORK_CODEX_INTEGRATION_EXECUTABLE
const realIt = executable && isAbsolute(executable) ? it : it.skip

describe('real 0.147 MCP transport control', () => {
  realIt(
    'initializes the strict document MCP server and lists its tool',
    async () => {
      const diagnostics: string[] = []
      const suspensionAuthority = createToolExecutionSuspensionAuthority()
      const server = await startDocumentMcpServer({ diagnostics: (code) => diagnostics.push(code) })
      const grant = Object.freeze({})
      const session = server.register({
        identity: { ownerId: 'o', host: 'docs', documentId: 'd', sessionId: 's', generation: 1 },
        manifest: createDocumentToolManifest({
          policyGrant: grant,
          consumePolicyGrant: (candidate) => {
            if (candidate !== grant) throw new Error('invalid_enhanced_policy_handle')
            return {
              generation: 1,
              host: 'docs' as const,
              policy: {
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
              },
              capabilities: ['semantic-read'] as const,
            }
          },
          tools: [
            { name: 'get_document_context', description: 'Read.', inputSchema: { type: 'object' } },
          ],
          policy: { get_document_context: 'read' },
        }),
        isOpen: () => true,
        executeRead: async () => ({ output: 'ok', summary: 'ok' }),
        suspendMutation: suspensionAuthority.suspend,
        ownsSuspension: suspensionAuthority.owns,
      })
      const manager = new CodexProcessManager({
        executablePath: executable!,
        bridge: { baseUrl: 'http://127.0.0.1:9', secret: 'control-secret' },
        mcp: { url: session.url, secret: session.secret },
        developerInstructions: 'Use the document tool.',
      })
      try {
        const client = await manager.start()
        await client.initialize()
        await client.startThread()
        await vi.waitFor(() => expect(diagnostics).toContain('mcp_tools_list'), { timeout: 10_000 })
      } finally {
        await manager.stop()
        session.close()
        await server.close()
      }
    },
    20_000,
  )
})
