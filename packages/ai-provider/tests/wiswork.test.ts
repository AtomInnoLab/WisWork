import { readFileSync } from 'node:fs'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage, AgentToolCall } from '@wiswork/agent-core'
import { chatForProvider } from '../src/chat'
import { resolveWisworkMainRequest, sanitizeWisworkSettings } from '../src/main-config'
import {
  AI_PROVIDERS,
  WISWORK_DEFAULT_MODEL,
  WISWORK_MESSAGES_URL,
  defaultAiSettings,
} from '../src/providers'
import { streamForProvider } from '../src/stream'
import { errorResponse, jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => vi.unstubAllGlobals())

const config = {
  apiKey: '',
  model: 'renderer-model-must-be-ignored',
}

const withToken =
  (token = 'login-access-token') =>
  async (request: (accessToken: string) => Promise<Response>) =>
    request(token)

function collector(signal = new AbortController().signal) {
  const deltas: string[] = []
  const toolCalls: AgentToolCall[] = []
  return {
    deltas,
    toolCalls,
    cb: {
      signal,
      onDelta: (text: string) => deltas.push(text),
      onToolCall: (call: AgentToolCall) => toolCalls.push(call),
    },
  }
}

describe('WisWork provider defaults and main-process config', () => {
  it('uses wiswork and the fixed WisModel model by default', () => {
    const meta = AI_PROVIDERS.find((provider) => provider.id === 'wiswork')
    expect(meta?.defaultModel).toBe('qwen/qwen3.8-max')
    expect(
      defaultAiSettings({ wiswork: 'renderer-key-must-be-ignored' }).providers.wiswork,
    ).toEqual({
      apiKey: '',
      model: WISWORK_DEFAULT_MODEL,
      baseUrl: undefined,
    })
  })

  it('fails closed when no authenticated session token is available', () => {
    expect(resolveWisworkMainRequest(false, undefined)).toEqual({
      ok: false,
      errorCode: 'auth_required',
    })
  })

  it('uses no service key and ignores renderer endpoint/key overrides', () => {
    expect(
      resolveWisworkMainRequest(true, {
        apiKey: 'renderer-key-must-be-ignored',
        model: 'deepseek/deepseek-v4-flash-0731',
        baseUrl: 'https://renderer.example.test/v1',
      }),
    ).toEqual({
      ok: true,
      provider: 'wiswork',
      config: { apiKey: '', model: 'qwen/qwen3.8-max' },
    })
  })

  it('uses the fixed default model and strips WisWork secrets before settings persistence/IPC', () => {
    const settings = defaultAiSettings()
    settings.provider = 'custom'
    settings.providers.wiswork = {
      apiKey: 'renderer-key-must-not-persist',
      model: '',
      baseUrl: 'https://renderer.example.test/v1',
    }
    const safe = sanitizeWisworkSettings(settings)
    expect(safe.provider).toBe('wiswork')
    expect(safe.providers.wiswork).toEqual({
      apiKey: '',
      model: WISWORK_DEFAULT_MODEL,
      baseUrl: undefined,
    })
  })
})

describe('WisUsage Anthropic Messages calls', () => {
  it('uses the exact fixed URL and login bearer token for one-shot chat', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ content: [{ type: 'text', text: 'ok' }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider('wiswork', config, 'sys', 'hi', undefined, withToken())
    expect(fetchMock).toHaveBeenCalledWith(
      WISWORK_MESSAGES_URL,
      expect.objectContaining({
        headers: {
          Authorization: 'Bearer login-access-token',
          'Content-Type': 'application/json',
        },
      }),
    )
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string)
    expect(body).toEqual({
      model: 'qwen/qwen3.8-max',
      max_tokens: 8192,
      system: 'sys',
      messages: [{ role: 'user', content: 'hi' }],
    })
  })

  it('emits Anthropic SSE text and fragmented tool calls', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([
              'data: {"type":"content_block_delta","index":0,"delta":{"type":"text_delta","text":"hello "}}',
              'data: {"type":"content_block_start","index":1,"content_block":{"type":"tool_use","id":"call-1","name":"lookup"}}',
              'data: {"type":"content_block_delta","index":1,"delta":{"type":"input_json_delta","partial_json":"{\\"q\\":\\"paper\\"}"}}',
              'data: {"type":"content_block_stop","index":1}',
              'data: {"type":"message_delta","delta":{"stop_reason":"tool_use"}}',
            ]),
          ),
        ),
    )
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider(
      'wiswork',
      config,
      'sys',
      [{ role: 'user', text: 'hi' }],
      [],
      100,
      cb,
      withToken(),
    )
    expect(deltas.join('')).toBe('hello ')
    expect(toolCalls).toEqual([
      { id: 'call-1', name: 'lookup', input: { q: 'paper' }, inputError: undefined },
    ])
  })

  it('uses Anthropic image blocks in the fixed request payload', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream([])))
    vi.stubGlobal('fetch', fetchMock)
    const messages: AgentMessage[] = [
      {
        role: 'user',
        text: 'look',
        images: [{ base64: 'aGVsbG8=', mime: 'image/png' }],
      },
    ]
    await streamForProvider(
      'wiswork',
      config,
      'sys',
      messages,
      [],
      100,
      collector().cb,
      withToken(),
    )
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string)
    expect(body.messages[0].content[1]).toEqual({
      type: 'image',
      source: { type: 'base64', media_type: 'image/png', data: 'aGVsbG8=' },
    })
  })

  it('requires the authenticated request boundary before network I/O', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)
    await expect(
      streamForProvider('wiswork', config, 'sys', [], [], 100, collector().cb),
    ).rejects.toMatchObject({ code: 'auth_required' })
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('preserves auth_required when the authenticated request boundary cannot refresh', async () => {
    const authFailure = Object.assign(new Error('session expired'), { code: 'auth_required' })
    const rejectedAuth = async () => {
      throw authFailure
    }

    await expect(
      streamForProvider('wiswork', config, 'sys', [], [], 100, collector().cb, rejectedAuth),
    ).rejects.toMatchObject({ code: 'auth_required' })
    await expect(
      chatForProvider('wiswork', config, 'sys', 'hi', undefined, rejectedAuth),
    ).rejects.toMatchObject({ code: 'auth_required' })
  })

  it('surfaces a repeated 401 as auth_required without including the bearer credential', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(errorResponse(401, 'invalid credential Bearer fake-upstream-secret')),
    )
    const run = streamForProvider(
      'wiswork',
      config,
      'sys',
      [],
      [],
      100,
      collector().cb,
      withToken('fake-upstream-secret'),
    )
    const result = expect(run).rejects.toMatchObject({ code: 'auth_required' })
    await result
  })

  it('honors caller cancellation', async () => {
    const controller = new AbortController()
    vi.stubGlobal(
      'fetch',
      vi.fn().mockImplementation((_url: string, init: RequestInit) => {
        return new Promise((_resolve, reject) => {
          init.signal?.addEventListener('abort', () => reject(new Error('cancelled')), {
            once: true,
          })
        })
      }),
    )
    const run = streamForProvider(
      'wiswork',
      config,
      'sys',
      [],
      [],
      100,
      collector(controller.signal).cb,
      withToken(),
    )
    controller.abort()
    await expect(run).rejects.toThrow('cancelled')
  })

  it('accepts a complete Anthropic non-stream JSON fallback', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        new Response(JSON.stringify({ content: [{ type: 'text', text: 'fallback' }] }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        }),
      ),
    )
    const { deltas, cb } = collector()
    await streamForProvider('wiswork', config, 'sys', [], [], 100, cb, withToken())
    expect(deltas).toEqual(['fallback'])
  })
})

describe('desktop WisModel security boundary', () => {
  const readRepo = (relative: string) =>
    readFileSync(new URL(`../../../${relative}`, import.meta.url), 'utf8')

  it('routes Docs, Sheets and Slides through the main-process resolver and auth status', () => {
    for (const relative of [
      'apps/docs/src/main/docs-main.ts',
      'apps/sheets/src/main/sheets-main.ts',
      'apps/slides/src/main/ai-ipc.ts',
    ]) {
      const source = readRepo(relative)
      expect(source).toContain('registerWisworkModelIpc')
      expect(source).toContain('getAccessToken()')
      expect(source).toContain('fetchWithAuth(')
      expect(source).toContain('isTrustedSender:')
    }
    const shell = readRepo('apps/shell/src/main/index.ts')
    expect(shell).toContain('HOME_CHANNELS.accountStatus')
    expect(shell).toContain('getValidAccountStatus()')
  })

  it('keeps the service env name and main config resolver out of renderer transports', () => {
    for (const relative of [
      'apps/docs/src/renderer/ai/transport.ts',
      'apps/sheets/src/renderer/ai/transport.ts',
      'apps/slides/src/renderer/ai/transport.ts',
      'apps/pdf/src/renderer/ai/transport.ts',
    ]) {
      const source = readRepo(relative)
      expect(source).not.toContain('WISWORK_MODEL_API_KEY')
      expect(source).not.toContain('resolveWisworkMainRequest')
    }
  })
})
