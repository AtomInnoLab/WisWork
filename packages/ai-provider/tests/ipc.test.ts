import { beforeEach, describe, expect, it, vi } from 'vitest'
import { AI_CONNECT_TIMEOUT_MS } from '../src/watchdog'
import {
  AI_IPC_LIMITS,
  registerWisworkModelIpc,
  type IpcMainLike,
  type WisworkIpcEvent,
} from '../src/ipc'
import { defaultAiSettings } from '../src/providers'
import { errorResponse, okResponse, sseStream } from './test-utils'

const ID = '11111111-1111-4111-8111-111111111111'
const OTHER_ID = '22222222-2222-4222-8222-222222222222'

function validRequest(requestId = ID) {
  return {
    requestId,
    settings: defaultAiSettings(),
    system: 'system',
    messages: [{ role: 'user' as const, text: 'hi' }],
    tools: [],
    maxTokens: 1024,
  }
}

function harness(options?: {
  trusted?: number[]
  accessToken?: string | null
  getAccessToken?: () => Promise<string | null>
  fetchWithAuth?: (request: (accessToken: string) => Promise<Response>) => Promise<Response>
}) {
  const accessToken =
    options && 'accessToken' in options ? options.accessToken : 'login-access-token'
  const handlers = new Map<string, (event: WisworkIpcEvent, ...args: unknown[]) => unknown>()
  const ipcMain: IpcMainLike = {
    handle: (channel, handler) => {
      handlers.set(channel, handler)
    },
  }
  const sent: unknown[] = []
  const event = (senderId: number): WisworkIpcEvent => ({
    sender: {
      id: senderId,
      isDestroyed: () => false,
      send: (_channel, chunk) => sent.push(chunk),
    },
  })
  const saved: unknown[] = []
  registerWisworkModelIpc({
    ipcMain,
    channels: {
      getSettings: 'get',
      setSettings: 'set',
      stream: 'stream',
      streamChunk: 'chunk',
      cancel: 'cancel',
      chat: 'chat',
    },
    isTrustedSender: (id) => (options?.trusted ?? [1]).includes(id),
    loadSettings: () => defaultAiSettings(),
    saveSettings: (settings) => saved.push(settings),
    getAccessToken: options?.getAccessToken ?? (async () => accessToken ?? null),
    fetchWithAuth:
      options?.fetchWithAuth ??
      (async (request) => {
        if (!accessToken) throw new Error('auth_required')
        return request(accessToken)
      }),
  })
  const invoke = (channel: string, senderId: number, ...args: unknown[]) =>
    handlers.get(channel)!(event(senderId), ...args)
  return { invoke, sent, saved }
}

beforeEach(() => vi.unstubAllGlobals())

describe('registerWisworkModelIpc', () => {
  it('rejects untrusted senders on every registered handler', async () => {
    const { invoke } = harness()
    for (const [channel, args] of [
      ['get', []],
      ['set', [defaultAiSettings()]],
      ['stream', [validRequest()]],
      ['cancel', [ID]],
      ['chat', [{ settings: defaultAiSettings(), system: 's', user: 'u' }]],
    ] as const) {
      await expect(Promise.resolve().then(() => invoke(channel, 9, ...args))).rejects.toMatchObject(
        {
          code: 'untrusted_sender',
        },
      )
    }
  })

  it('rejects malformed ids, excessive tokens, oversized text, and unknown fields', async () => {
    const { invoke } = harness()
    await expect(invoke('stream', 1, validRequest('short'))).rejects.toMatchObject({
      code: 'invalid_request_id',
    })
    await expect(
      invoke('stream', 1, { ...validRequest(), maxTokens: AI_IPC_LIMITS.maxTokens + 1 }),
    ).rejects.toMatchObject({ code: 'invalid_payload' })
    await expect(
      invoke('stream', 1, {
        ...validRequest(),
        system: 'x'.repeat(AI_IPC_LIMITS.maxSystemChars + 1),
      }),
    ).rejects.toMatchObject({ code: 'payload_too_large' })
    await expect(
      invoke('stream', 1, { ...validRequest(), unexpected: true }),
    ).rejects.toMatchObject({ code: 'invalid_payload' })
  })

  it('enforces collection, image, and nested JSON limits', async () => {
    const { invoke } = harness()
    await expect(
      invoke('stream', 1, {
        ...validRequest(),
        messages: Array.from({ length: AI_IPC_LIMITS.maxMessages + 1 }, () => ({
          role: 'user',
          text: 'x',
        })),
      }),
    ).rejects.toMatchObject({ code: 'payload_too_large' })
    await expect(
      invoke('stream', 1, {
        ...validRequest(),
        messages: [
          {
            role: 'user',
            text: '',
            images: [{ mime: 'image/png', base64: 'x'.repeat(AI_IPC_LIMITS.maxImageChars + 1) }],
          },
        ],
      }),
    ).rejects.toMatchObject({ code: 'payload_too_large' })
    let nested: unknown = 'leaf'
    for (let depth = 0; depth <= AI_IPC_LIMITS.maxJsonDepth; depth += 1) nested = { next: nested }
    await expect(
      invoke('stream', 1, {
        ...validRequest(),
        tools: [{ name: 'nested', description: '', inputSchema: nested }],
      }),
    ).rejects.toMatchObject({ code: 'payload_too_large' })
  })

  it('accepts bounded tool image content and rejects oversized tool images', async () => {
    const { invoke } = harness({
      fetchWithAuth: async () => okResponse(sseStream([])),
    })
    const messages: any[] = [
      {
        role: 'tool',
        results: [
          {
            id: 't1',
            name: 'shot',
            output: 'ok',
            content: [{ type: 'image', image: { mime: 'image/png', base64: 'AAAA' } }],
          },
        ],
      },
    ]
    await expect(invoke('stream', 1, { ...validRequest(), messages })).resolves.toBeUndefined()
    messages[0].results[0].content[0].image.base64 = 'x'.repeat(AI_IPC_LIMITS.maxImageChars + 1)
    await expect(invoke('stream', 1, { ...validRequest(), messages })).rejects.toMatchObject({
      code: 'payload_too_large',
    })
  })
  it('rejects duplicate request ids and cross-sender cancellation', async () => {
    let release!: () => void
    vi.stubGlobal(
      'fetch',
      vi.fn().mockImplementation(
        () =>
          new Promise<Response>((resolve) => {
            release = () => resolve(okResponse(sseStream([])))
          }),
      ),
    )
    const { invoke } = harness({ trusted: [1, 2] })
    const first = invoke('stream', 1, validRequest())
    await vi.waitFor(() => expect(release).toBeTypeOf('function'))
    await expect(invoke('stream', 1, validRequest())).rejects.toMatchObject({
      code: 'duplicate_request_id',
    })
    await expect(Promise.resolve().then(() => invoke('cancel', 2, ID))).rejects.toMatchObject({
      code: 'request_owner_mismatch',
    })
    release()
    await first
  })

  it('maps a WisWork watchdog timeout to the stable timeout code', async () => {
    vi.useFakeTimers()
    try {
      vi.stubGlobal(
        'fetch',
        vi.fn().mockImplementation(
          (_url: string, init: RequestInit) =>
            new Promise((_resolve, reject) =>
              init.signal!.addEventListener('abort', () => reject(new Error('aborted')), {
                once: true,
              }),
            ),
        ),
      )
      const { invoke, sent } = harness()
      const run = invoke('stream', 1, validRequest())
      await vi.advanceTimersByTimeAsync(AI_CONNECT_TIMEOUT_MS)
      await run
      expect(sent.at(-1)).toMatchObject({ type: 'error', error: 'timeout', errorCode: 'timeout' })
    } finally {
      vi.useRealTimers()
    }
  })

  it('maps 401/403 bodies containing credentials to stable safe errors', async () => {
    for (const status of [401, 403]) {
      vi.stubGlobal(
        'fetch',
        vi
          .fn()
          .mockResolvedValue(errorResponse(status, 'Bearer fake-upstream-secret and fake-key')),
      )
      const { invoke, sent } = harness()
      await invoke('stream', 1, validRequest(status === 401 ? ID : OTHER_ID))
      expect(sent.at(-1)).toMatchObject({
        type: 'error',
        error: status === 401 ? 'auth_required' : 'model_credentials_missing',
        errorCode: status === 401 ? 'auth_required' : 'model_credentials_missing',
        diagnostic: { stage: 'response', httpStatus: status },
      })
      expect(JSON.stringify(sent)).not.toContain('fake-upstream-secret')
      expect(JSON.stringify(sent)).not.toContain('fake-key')
    }
  })

  it('fails closed with auth_required in standalone and never calls fetch', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    const { invoke, sent } = harness({ accessToken: null })
    await invoke('stream', 1, validRequest())
    expect(sent).toContainEqual(
      expect.objectContaining({ type: 'error', errorCode: 'auth_required' }),
    )
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('fails closed without leaking details when session validation cannot refresh', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    const { invoke, sent } = harness({
      getAccessToken: async () => {
        throw new Error('temporary refresh failure containing private response')
      },
    })
    await invoke('stream', 1, validRequest())
    expect(sent).toContainEqual(
      expect.objectContaining({ type: 'error', errorCode: 'model_upstream_unavailable' }),
    )
    await expect(
      invoke('chat', 1, { settings: defaultAiSettings(), system: 's', user: 'u' }),
    ).resolves.toMatchObject({ ok: false, errorCode: 'model_upstream_unavailable' })
    expect(fetchMock).not.toHaveBeenCalled()
    expect(JSON.stringify(sent)).not.toContain('private response')
  })

  it('preserves auth_required when the session has expired during validation', async () => {
    const authFailure = Object.assign(new Error('private expired-session detail'), {
      code: 'auth_required',
    })
    const { invoke, sent } = harness({
      getAccessToken: async () => {
        throw authFailure
      },
    })

    await invoke('stream', 1, validRequest())
    expect(sent.at(-1)).toMatchObject({ type: 'error', errorCode: 'auth_required' })
    await expect(
      invoke('chat', 1, { settings: defaultAiSettings(), system: 's', user: 'u' }),
    ).resolves.toMatchObject({ ok: false, errorCode: 'auth_required' })
    expect(JSON.stringify(sent)).not.toContain('private expired-session detail')
  })

  it('forces the default model and strips all provider keys/base URLs from settings', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
    vi.stubGlobal('fetch', fetchMock)
    const settings = defaultAiSettings()
    settings.providers.wiswork.model = 'renderer-model'
    settings.providers.custom.apiKey = 'renderer-key'
    settings.providers.custom.baseUrl = 'https://renderer.invalid/v1'
    const { invoke, saved } = harness()
    const returned = (await invoke('get', 1)) as ReturnType<typeof defaultAiSettings>
    await invoke('set', 1, settings)
    for (const value of [returned, saved[0] as typeof returned]) {
      expect(value.providers.wiswork.model).toBe('openai/gpt-5.6-sol')
      for (const config of Object.values(value.providers)) {
        expect(config.apiKey).toBe('')
        expect(config.baseUrl).toBeUndefined()
      }
    }
    await invoke('stream', 1, { ...validRequest(), settings })
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string)
    expect(body.model).toBe('openai/gpt-5.6-sol')
  })

  it('delegates WisUsage requests to the auth client so a 401 refreshes and retries once', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce(errorResponse(401, 'expired'))
      .mockResolvedValueOnce(
        okResponse(
          sseStream([
            'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"ok"}}',
          ]),
        ),
      )
    vi.stubGlobal('fetch', fetchMock)
    const fetchWithAuth = async (request: (accessToken: string) => Promise<Response>) => {
      let response = await request('expired-login-token')
      if (response.status === 401) response = await request('refreshed-login-token')
      return response
    }
    const { invoke, sent } = harness({ fetchWithAuth })

    await invoke('stream', 1, validRequest())

    expect(fetchMock).toHaveBeenCalledTimes(2)
    expect(fetchMock.mock.calls[0]![0]).toBe('https://wisusage.dev.atominnolab.com/v1/messages')
    expect(fetchMock.mock.calls[0]![1].headers).toMatchObject({
      Authorization: 'Bearer expired-login-token',
    })
    expect(fetchMock.mock.calls[1]![1].headers).toMatchObject({
      Authorization: 'Bearer refreshed-login-token',
    })
    expect(sent).toContainEqual(expect.objectContaining({ type: 'delta', text: 'ok' }))
    expect(JSON.stringify(sent)).not.toContain('login-token')
  })
})
