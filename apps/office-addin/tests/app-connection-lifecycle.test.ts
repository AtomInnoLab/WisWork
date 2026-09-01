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

  it('keeps the taskpane disconnected and offers a retry when durable forget fails', async () => {
    const documentClient: OfficeDocumentClient = {
      initialize: vi.fn().mockResolvedValue('word'),
      readSelection: vi.fn().mockResolvedValue(''),
      replaceSelection: vi.fn().mockResolvedValue(undefined),
      appendText: vi.fn().mockResolvedValue(undefined),
    }
    let snapshot: OfficeRelaySnapshot = { status: 'connected', capabilities: ['agent.v1'] }
    let notify: () => void = () => undefined
    let attempts = 0
    const bridge: OfficeRelaySession = {
      snapshot: () => snapshot,
      subscribe: (listener) => {
        notify = listener
        return () => undefined
      },
      connect: vi.fn().mockResolvedValue(undefined),
      disconnect: vi.fn(),
      forget: vi.fn(async () => {
        snapshot = { status: 'offline' }
        notify()
        attempts += 1
        if (attempts === 1) throw new Error('binding_storage_unavailable')
      }),
      authenticatedFetch: vi.fn(),
      capabilityFetch: vi.fn(),
      sendDiagnostic: vi.fn().mockResolvedValue(undefined),
    }
    const container = documentElement()
    const root = createRoot(container)
    const agentSnapshot = {
      assistantText: '',
      activity: '',
      busy: false,
      applying: false,
      status: 'idle',
      retryable: false,
      timeline: [{ id: 'user-1', kind: 'user', text: 'Hello' }],
    }
    const agentSession = {
      snapshot: () => agentSnapshot,
      subscribe: () => () => undefined,
      send: vi.fn(),
      stop: vi.fn(),
      confirm: vi.fn(),
      reject: vi.fn(),
      newTask: vi.fn(),
      retry: vi.fn(),
      logout: vi.fn(),
      authenticationLost: vi.fn(),
      dispose: vi.fn(),
    }
    const runtime = {
      dispose: vi.fn(),
      clearSession: vi.fn(),
      disableElevatedOffice: vi.fn(),
    }
    const ui = {
      attachments: () => [],
      skills: () => [],
      skillPackagesEnabled: false,
      upload: vi.fn(),
      clear: vi.fn(),
    }
    const workspaceFactory = () => ({ runtime, session: agentSession, ui })

    await act(async () => {
      root.render(
        React.createElement(
          ConfiguredApp as React.ComponentType<{
            documentClient?: OfficeDocumentClient
            connectionBridge?: OfficeRelaySession
            workspaceFactory?: () => {
              runtime: unknown
              session: unknown
              ui: unknown
            }
          }>,
          { documentClient, connectionBridge: bridge, workspaceFactory },
        ),
      )
      await flush()
    })
    const logout = Array.from(container.querySelectorAll('button')).find(
      (button) => button.textContent === '退出登录',
    )
    expect(logout, container.innerHTML).toBeDefined()

    await act(async () => {
      logout!.click()
      await flush()
    })

    expect(container.textContent).toContain('Couldn’t forget this Office pairing')
    expect(container.textContent).toContain('Try forgetting again')
    expect(snapshot).toEqual({ status: 'offline' })
    expect(agentSession.logout).toHaveBeenCalledOnce()
    expect(runtime.disableElevatedOffice).toHaveBeenCalled()
    expect(runtime.clearSession).toHaveBeenCalled()
    expect(runtime.dispose).not.toHaveBeenCalled()

    const retry = Array.from(container.querySelectorAll('button')).find(
      (button) => button.textContent === 'Try forgetting again',
    )
    await act(async () => {
      retry!.click()
      await flush()
    })

    expect(bridge.forget).toHaveBeenCalledTimes(2)
    expect(container.textContent).not.toContain('Couldn’t forget this Office pairing')
    expect(container.textContent).toContain('Connect to WisWork PC')
    act(() => root.unmount())
    expect(runtime.dispose).toHaveBeenCalledOnce()
  })
})

function documentElement(): HTMLDivElement {
  const element = globalThis.document.createElement('div')
  globalThis.document.body.append(element)
  return element
}
