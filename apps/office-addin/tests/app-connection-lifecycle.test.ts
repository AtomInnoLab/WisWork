// @vitest-environment jsdom

import React, { act } from 'react'
import { createRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import { ConfiguredApp } from '../src/App.js'
import type { OfficeDocumentClient } from '../src/office-document.js'
import type { OfficeRelaySession, OfficeRelaySnapshot } from '../src/relay/session.js'

;(
  globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT: boolean }
).IS_REACT_ACT_ENVIRONMENT = true

const flush = async () => {
  await Promise.resolve()
  await Promise.resolve()
}

describe('configured Office connection lifecycle', () => {
  it('automatically connects the initialized supported host and disconnects on unmount', async () => {
    const documentClient: OfficeDocumentClient = {
      initialize: vi.fn().mockResolvedValue('word'),
      readSelection: vi.fn().mockResolvedValue(''),
      replaceSelection: vi.fn().mockResolvedValue(undefined),
      appendText: vi.fn().mockResolvedValue(undefined),
    }
    const snapshot: OfficeRelaySnapshot = { status: 'offline' }
    const bridge: OfficeRelaySession = {
      snapshot: () => snapshot,
      subscribe: () => () => undefined,
      connect: vi.fn().mockResolvedValue(undefined),
      disconnect: vi.fn(),
      forget: vi.fn().mockResolvedValue(undefined),
      authenticatedFetch: vi.fn(),
      capabilityFetch: vi.fn(),
      sendDiagnostic: vi.fn().mockResolvedValue(undefined),
    }
    const container = documentElement()
    const root = createRoot(container)

    await act(async () => {
      root.render(
        React.createElement(
          ConfiguredApp as React.ComponentType<{
            documentClient?: OfficeDocumentClient
            connectionBridge?: OfficeRelaySession
          }>,
          { documentClient, connectionBridge: bridge },
        ),
      )
      await flush()
    })

    expect(documentClient.initialize).toHaveBeenCalledOnce()
    expect(bridge.connect).toHaveBeenCalledOnce()
    expect(bridge.connect).toHaveBeenCalledWith('word')

    act(() => root.unmount())
    expect(bridge.disconnect).toHaveBeenCalledOnce()
  })
})

function documentElement(): HTMLDivElement {
  const element = globalThis.document.createElement('div')
  globalThis.document.body.append(element)
  return element
}
