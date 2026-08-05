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

function harness(options?: { trusted?: number[]; loggedIn?: boolean }) {
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
    getLoggedIn: async () => options?.loggedIn ?? true,
  })
  const invoke = (channel: string, senderId: number, ...args: unknown[]) =>
    handlers.get(channel)!(event(senderId), ...args)
  return { invoke, sent, saved }
}

beforeEach(() => {
  vi.unstubAllGlobals()
  vi.stubEnv('WISWORK_MODEL_API_KEY', 'fake-main-key')
})

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
  it('rejects duplicate request ids and cross-sender cancellation', async () => {
    let release!: () => void
    vi.stubGlobal(
      'fetch',
      vi.fn().mockImplementation(
        () =>
          new Promise<Response>((resolve) => {
            release = () => resolve(okResponse(sseStream(['data: [DONE]'])))
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
        vi
          .fn()
          .mockImplementation(
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
        error: 'model_credentials_missing',
        errorCode: 'model_credentials_missing',
      })
      expect(JSON.stringify(sent)).not.toContain('fake-upstream-secret')
      expect(JSON.stringify(sent)).not.toContain('fake-key')
    }
  })

  it('fails closed with auth_required in standalone and never calls fetch', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    const { invoke, sent } = harness({ loggedIn: false })
    await invoke('stream', 1, validRequest())
    expect(sent).toContainEqual(
      expect.objectContaining({ type: 'error', errorCode: 'auth_required' }),
    )
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('forces the default model and strips all provider keys/base URLs from settings', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream(['data: [DONE]'])))
    vi.stubGlobal('fetch', fetchMock)
    const settings = defaultAiSettings()
    settings.providers.wiswork.model = 'renderer-model'
    settings.providers.custom.apiKey = 'renderer-key'
    settings.providers.custom.baseUrl = 'https://renderer.invalid/v1'
    const { invoke, saved } = harness()
    const returned = (await invoke('get', 1)) as ReturnType<typeof defaultAiSettings>
    await invoke('set', 1, settings)
    for (const value of [returned, saved[0] as typeof returned]) {
      expect(value.providers.wiswork.model).toBe('deepseek/deepseek-v4-flash-0731')
      for (const config of Object.values(value.providers)) {
        expect(config.apiKey).toBe('')
        expect(config.baseUrl).toBeUndefined()
      }
    }
    await invoke('stream', 1, { ...validRequest(), settings })
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string)
    expect(body.model).toBe('deepseek/deepseek-v4-flash-0731')
  })
})
