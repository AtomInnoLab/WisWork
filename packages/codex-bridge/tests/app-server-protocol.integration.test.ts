import { isAbsolute } from 'node:path'
import { describe, expect, it } from 'vitest'
import { CodexProcessManager } from '../src/process-manager.js'

const executable = process.env.WISWORK_CODEX_INTEGRATION_EXECUTABLE
const realIt = executable !== undefined && isAbsolute(executable) ? it : it.skip

describe('real pinned Codex app-server stdio lifecycle', () => {
  realIt(
    'launches 0.147.0, initializes over stdio, and shuts down cleanly',
    async () => {
      const manager = new CodexProcessManager({
        executablePath: executable!,
        bridge: { baseUrl: 'http://127.0.0.1:9', secret: 'integration-local-secret' },
        developerInstructions: 'Use only the approved document-scoped carrier.',
        startupTimeoutMs: 10_000,
      })
      try {
        const client = await manager.start()
        await expect(client.initialize()).resolves.toEqual({
          userAgent: expect.stringContaining('/0.147.0'),
          platformFamily: expect.any(String),
          platformOs: expect.any(String),
        })
        await expect(client.startThread()).resolves.toMatchObject({
          thread: { id: expect.any(String) },
        })
      } finally {
        await manager.stop()
      }
    },
    20_000,
  )
})
