import { createHash } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { describe, expect, it } from 'vitest'
import manifest from '../src/generated/schema-manifest.json'
import {
  CODEX_CLI_VERSION,
  CODEX_SCHEMA_SHA256,
  KNOWN_SERVER_NOTIFICATION_METHODS,
} from '../src/generated/index.js'

describe('pinned Codex app-server fixture contract', () => {
  it('binds the checked-in protocol subset to the reviewed 0.147.0 schema digests', async () => {
    expect(CODEX_CLI_VERSION).toBe('codex-cli 0.147.0')
    expect(CODEX_SCHEMA_SHA256).toEqual({
      protocol: manifest.sha256['codex_app_server_protocol.schemas.json'],
      protocolV2: manifest.sha256['codex_app_server_protocol.v2.schemas.json'],
      initialize: manifest.sha256['v1/InitializeParams.json'],
      initializeResponse: manifest.sha256['v1/InitializeResponse.json'],
      threadStart: manifest.sha256['v2/ThreadStartParams.json'],
      threadStartResponse: manifest.sha256['v2/ThreadStartResponse.json'],
      turnStart: manifest.sha256['v2/TurnStartParams.json'],
      turnStartResponse: manifest.sha256['v2/TurnStartResponse.json'],
      turnInterrupt: manifest.sha256['v2/TurnInterruptParams.json'],
      turnInterruptResponse: manifest.sha256['v2/TurnInterruptResponse.json'],
      serverNotification: manifest.sha256['ServerNotification.json'],
    })
    const bindings = await readFile(
      new URL('../src/generated/codex-app-server-0.147.ts', import.meta.url),
    )
    expect(createHash('sha256').update(bindings).digest('hex')).toBe(manifest.bindingsSha256)
    expect(new Set(KNOWN_SERVER_NOTIFICATION_METHODS).size).toBe(
      KNOWN_SERVER_NOTIFICATION_METHODS.length,
    )
  })
})
