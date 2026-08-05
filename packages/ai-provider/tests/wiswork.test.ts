import { readFileSync } from 'node:fs'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { AgentMessage, AgentToolCall } from '@wiswork/agent-core'
import { chatForProvider } from '../src/chat'
import { resolveWisworkMainRequest, sanitizeWisworkSettings } from '../src/main-config'
import { AI_PROVIDERS, WISWORK_DEFAULT_MODEL, defaultAiSettings } from '../src/providers'
import { streamForProvider } from '../src/stream'
import { errorResponse, jsonResponse, okResponse, sseStream } from './test-utils'

afterEach(() => vi.unstubAllGlobals())

const config = {
  apiKey: 'main-process-test-key',
  model: 'renderer-model-must-be-ignored',
}

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
    expect(meta?.defaultModel).toBe('deepseek/deepseek-v4-flash-0731')
    expect(
      defaultAiSettings({ wiswork: 'renderer-key-must-be-ignored' }).providers.wiswork,
    ).toEqual({
      apiKey: '',
      model: WISWORK_DEFAULT_MODEL,
      baseUrl: undefined,
    })
  })

  it('fails authentication before checking the service credential', () => {
    expect(resolveWisworkMainRequest(false, undefined, {})).toEqual({
      ok: false,
      errorCode: 'auth_required',
    })
  })

  it('reports a missing service key separately after login', () => {
    expect(resolveWisworkMainRequest(true, undefined, {})).toEqual({
      ok: false,
      errorCode: 'model_credentials_missing',
    })
  })

  it('takes only the main-process env key and ignores renderer endpoint/key overrides', () => {
    expect(
      resolveWisworkMainRequest(
        true,
        {
          apiKey: 'renderer-key-must-be-ignored',
          model: 'deepseek/deepseek-v4-flash-0731',
          baseUrl: 'https://renderer.example.test/v1',
        },
        { WISWORK_MODEL_API_KEY: 'main-process-test-key' },
      ),
    ).toEqual({
      ok: true,
      provider: 'wiswork',
      config: { apiKey: 'main-process-test-key', model: 'deepseek/deepseek-v4-flash-0731' },
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

describe('WisWork OpenAI-compatible calls', () => {
  it('uses the fixed URL and bearer header for one-shot chat', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(jsonResponse({ choices: [{ message: { content: 'ok' } }] }))
    vi.stubGlobal('fetch', fetchMock)
    await chatForProvider('wiswork', config, 'sys', 'hi')
    expect(fetchMock).toHaveBeenCalledWith(
      'https://wismodel-proxy-dev.atominnolab.com/api/v1/chat/completions',
      expect.objectContaining({
        headers: expect.objectContaining({ Authorization: 'Bearer main-process-test-key' }),
      }),
    )
  })

  it('emits SSE text and fragmented tool calls', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(
          okResponse(
            sseStream([
              'data: {"choices":[{"delta":{"content":"hello "},"finish_reason":null}]}',
              'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call-1","function":{"name":"lookup","arguments":"{\\"q\\":"}}]},"finish_reason":null}]}',
              'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"\\"paper\\"}"}}]},"finish_reason":"tool_calls"}]}',
              'data: [DONE]',
            ]),
          ),
        ),
    )
    const { deltas, toolCalls, cb } = collector()
    await streamForProvider('wiswork', config, 'sys', [{ role: 'user', text: 'hi' }], [], 100, cb)
    expect(deltas.join('')).toBe('hello ')
    expect(toolCalls).toEqual([
      { id: 'call-1', name: 'lookup', input: { q: 'paper' }, inputError: undefined },
    ])
  })

  it('preserves image data URLs in the OpenAI-compatible payload', async () => {
    const fetchMock = vi.fn().mockResolvedValue(okResponse(sseStream(['data: [DONE]'])))
    vi.stubGlobal('fetch', fetchMock)
    const messages: AgentMessage[] = [
      {
        role: 'user',
        text: 'look',
        images: [{ base64: 'aGVsbG8=', mime: 'image/png' }],
      },
    ]
    await streamForProvider('wiswork', config, 'sys', messages, [], 100, collector().cb)
    const body = JSON.parse(fetchMock.mock.calls[0]![1].body as string)
    expect(body.messages[1].content[1]).toEqual({
      type: 'image_url',
      image_url: { url: 'data:image/png;base64,aGVsbG8=' },
    })
  })

  it('surfaces 401 without including the bearer credential', async () => {
    vi.stubGlobal(
      'fetch',
      vi
        .fn()
        .mockResolvedValue(errorResponse(401, 'invalid credential Bearer fake-upstream-secret')),
    )
    const run = streamForProvider('wiswork', config, 'sys', [], [], 100, collector().cb)
    const result = expect(run).rejects.toMatchObject({ code: 'model_credentials_missing' })
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
    )
    controller.abort()
    await expect(run).rejects.toThrow('cancelled')
  })

  it('accepts a complete non-stream JSON fallback', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        new Response(JSON.stringify({ choices: [{ message: { content: 'fallback' } }] }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        }),
      ),
    )
    const { deltas, cb } = collector()
    await streamForProvider('wiswork', config, 'sys', [], [], 100, cb)
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
      expect(source).toContain('getValidAccountStatus()')
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
