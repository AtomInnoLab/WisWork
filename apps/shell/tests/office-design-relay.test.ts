import { expect, it, vi } from 'vitest'
import { createOfficeRelayClient, type RelaySocket } from '../src/main/office-relay-client'

it('negotiates DESIGN.md and dispatches file requests without invoking the model', async () => {
  const sent: Record<string, any>[] = []
  const listeners = new Map<string, (event: any) => void>()
  const socket: RelaySocket = {
    readyState: 0,
    addEventListener: (name, listener) => {
      listeners.set(name, listener)
    },
    send: (raw) => {
      sent.push(JSON.parse(raw))
    },
    close: () => {
      socket.readyState = 3
    },
  }
  const receive = (value: unknown) => listeners.get('message')?.({ data: JSON.stringify(value) })
  const proxy = vi.fn(async () => ({ status: 200, body: new Uint8Array() }))
  const designDocument = vi.fn(async () => ({
    documentId: 'document_12345678',
    markdown: '# DESIGN.md',
    revision: 'a'.repeat(64),
  }))
  const client = createOfficeRelayClient({
    endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
    connect: () => socket,
    getValidAccountStatus: async () => ({ loggedIn: true }),
    getAccessToken: async () => 'test-token',
    proxy,
    designDocument,
    negotiateCapabilities: true,
    onPending() {},
  })
  try {
    const claim = client.claim('123456')
    await vi.waitFor(() => expect(listeners.has('open')).toBe(true))
    socket.readyState = 1
    listeners.get('open')?.({})
    await claim
    expect(sent.find((frame) => frame.type === 'pc.negotiate')?.capabilities).toContain(
      'design-document.v1',
    )
    const capabilities = ['agent.v1', 'design-document.v1']
    receive({ version: 2, type: 'pc.negotiated', pairing_version: 2, capabilities })
    receive({
      version: 2,
      type: 'pc.claimed',
      pairing_id: 'pairing_12345678',
      host: 'PowerPoint',
      origin: 'https://office.8-216-134-194.sslip.io',
      verification_code: '123456',
      expires_in: 120,
      capabilities,
    })
    expect(await client.approve('pairing_12345678')).toBe(true)
    receive({
      version: 2,
      type: 'pc.approved',
      session_id: 'session_12345678',
      capability: 'capability_12345678',
      expires_in: 1800,
      capabilities,
    })
    const body = { action: 'open', documentId: 'document_12345678', markdown: '# DESIGN.md' }
    receive({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: 'request_12345678',
      capability_name: 'design-document.v1',
      body,
    })
    await vi.waitFor(() => expect(sent.some((frame) => frame.type === 'pc.done')).toBe(true))
    expect(designDocument).toHaveBeenCalledWith(
      expect.objectContaining({
        body,
        sessionId: 'session_12345678',
        host: 'PowerPoint',
        signal: expect.any(AbortSignal),
      }),
    )
    expect(proxy).not.toHaveBeenCalled()
    const chunk = sent.find((frame) => frame.type === 'pc.chunk')!
    expect(JSON.parse(Buffer.from(chunk.data, 'base64').toString())).toMatchObject({
      markdown: '# DESIGN.md',
    })
  } finally {
    client.revoke('test-complete')
  }
})

async function requestBudgetSession(designNegotiated: boolean, maxRequestIds?: number) {
  const listeners = new Map<string, (event: any) => void>()
  let finishRequest: ((completed: boolean) => void) | undefined
  const socket: RelaySocket = {
    readyState: 0,
    addEventListener: (name, listener) => {
      listeners.set(name, listener)
    },
    send: (raw) => {
      if (JSON.parse(raw).type === 'pc.done') finishRequest?.(true)
    },
    close: () => {
      socket.readyState = 3
    },
  }
  const receive = (value: unknown) => listeners.get('message')?.({ data: JSON.stringify(value) })
  const designDocument = vi.fn(async () => ({
    documentId: 'document_12345678',
    markdown: '# DESIGN.md',
    revision: 'a'.repeat(64),
  }))
  const client = createOfficeRelayClient({
    endpoint: 'wss://office.8-216-134-194.sslip.io/office-relay',
    connect: () => socket,
    getValidAccountStatus: async () => ({ loggedIn: true }),
    getAccessToken: async () => 'test-token',
    proxy: async () => ({ status: 200, body: new Uint8Array() }),
    designDocument,
    negotiateCapabilities: true,
    ...(maxRequestIds === undefined ? {} : { maxRequestIds }),
    onPending() {},
    onStatus(status) {
      if (status.startsWith('disconnected')) finishRequest?.(false)
    },
  })
  const claiming = client.claim('123456')
  await vi.waitFor(() => expect(listeners.has('open')).toBe(true))
  socket.readyState = 1
  listeners.get('open')?.({})
  await claiming
  const capabilities = designNegotiated ? ['agent.v1', 'design-document.v1'] : ['agent.v1']
  receive({ version: 2, type: 'pc.negotiated', pairing_version: 2, capabilities })
  receive({
    version: 2,
    type: 'pc.claimed',
    pairing_id: 'pairing_12345678',
    host: 'PowerPoint',
    origin: 'https://office.8-216-134-194.sslip.io',
    verification_code: '123456',
    expires_in: 120,
    capabilities,
  })
  await client.approve('pairing_12345678')
  receive({
    version: 2,
    type: 'pc.approved',
    session_id: 'session_12345678',
    capability: 'capability_12345678',
    expires_in: 1800,
    capabilities,
  })
  const dispatch = async (requestId: string, design = designNegotiated) => {
    const completed = new Promise<boolean>((resolve) => {
      finishRequest = resolve
    })
    receive({
      version: 2,
      type: 'relay.request',
      session_id: 'session_12345678',
      request_id: requestId,
      capability_name: design ? 'design-document.v1' : 'agent.v1',
      body: design ? { action: 'read', documentId: 'document_12345678' } : {},
    })
    const result = await completed
    finishRequest = undefined
    return result
  }
  return { client, dispatch, designDocument }
}

it('keeps an eight-hour DESIGN polling session alive and still rejects its first request ID replay', async () => {
  const f = await requestBudgetSession(true)
  try {
    const eightHourPolls = Math.ceil((8 * 60 * 60 * 1000) / 5000)
    for (let index = 0; index < eightHourPolls; index++) {
      expect(await f.dispatch(`request_poll_${index}`), `poll ${index}`).toBe(true)
    }
    expect(f.designDocument).toHaveBeenCalledTimes(eightHourPolls)
    expect(await f.dispatch('request_poll_0')).toBe(false)
    expect(f.designDocument).toHaveBeenCalledTimes(eightHourPolls)
    expect(f.client.status()).toBe('disconnected:protocol_violation')
  } finally {
    f.client.revoke('test-complete')
  }
})

it.each([
  { design: true, override: undefined, limit: 5760 + 2048 },
  { design: false, override: undefined, limit: 2048 },
  { design: true, override: 2, limit: 2 },
])(
  'bounds the negotiated design=$design request budget at $limit and preserves explicit overrides',
  async ({ design, override, limit }) => {
    const f = await requestBudgetSession(design, override)
    try {
      for (let index = 0; index < limit; index++) {
        expect(
          await f.dispatch(`request_budget_${index}`, design && index < 5760),
          `request ${index}`,
        ).toBe(true)
      }
      expect(await f.dispatch('request_budget_exhausted')).toBe(false)
      expect(f.client.status()).toBe('disconnected:protocol_violation')
    } finally {
      f.client.revoke('test-complete')
    }
  },
)
