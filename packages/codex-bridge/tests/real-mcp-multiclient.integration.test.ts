import { isAbsolute } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import { startDynamicMcpGateway } from '../src/dynamic-mcp-gateway.js'
import { CodexProcessManager } from '../src/process-manager.js'

const executable = process.env.WISWORK_CODEX_INTEGRATION_EXECUTABLE
const realIt = executable && isAbsolute(executable) ? it : it.skip

describe('real 0.147 shared gateway clients', () => {
  realIt(
    'initializes independent clients for successive native threads',
    async () => {
      const diagnostics: string[] = []
      const gateway = await startDynamicMcpGateway((code) => diagnostics.push(code))
      const manager = new CodexProcessManager({
        executablePath: executable!,
        bridge: { baseUrl: 'http://127.0.0.1:9', secret: 'local-control-only' },
        mcp: { url: gateway.url, secret: gateway.secret },
        developerInstructions: 'Use bounded document tools.',
      })
      try {
        const client = await manager.start()
        await client.initialize()
        for (let index = 0; index < 3; index++) {
          const before = diagnostics.filter((code) => code === 'mcp_tools_list').length
          await client.startThread()
          await vi.waitFor(
            () => {
              expect(
                diagnostics.filter((code) => code === 'mcp_tools_list').length,
              ).toBeGreaterThan(before)
            },
            { timeout: 5_000 },
          )
        }
      } finally {
        await manager.stop()
        await gateway.close()
      }
    },
    30_000,
  )
})
