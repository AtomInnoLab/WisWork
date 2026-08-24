import { describe, expect, it, vi } from 'vitest'
import { createSheetsChatBindingCoordinator } from '../src/renderer/ai/chat-binding'

const deferred = <T>() => {
  let resolve!: (value: T) => void
  let reject!: (reason: unknown) => void
  const promise = new Promise<T>((yes, no) => {
    resolve = yes
    reject = no
  })
  return { promise, resolve, reject }
}

const target = (documentId: string, path: string | null, sessionId: string) => ({
  documentId,
  path,
  sessionId,
})

describe('Sheets chat binding coordinator', () => {
  it.each([
    [target('doc-a', null, 'session-1'), target('doc-a', '/saved.xlsx', 'session-2')],
    [target('doc-a', '/a.xlsx', 'session-1'), target('doc-a', '/b.xlsx', 'session-2')],
    [target('doc-a', '/a.xlsx', 'session-1'), target('doc-a', '/a.xlsx', 'session-2')],
  ])(
    'suspends persistence during same-document rebind and flushes to the new IDs',
    async (first, next) => {
      const rebound = deferred<{ projectId: string; chatId: string } | null>()
      const appendChat = vi.fn().mockResolvedValue(undefined)
      const onBinding = vi.fn()
      const api = {
        resolveChat: vi.fn().mockResolvedValue({ projectId: 'p', chatId: 'old' }),
        rebindChat: vi.fn().mockReturnValue(rebound.promise),
        loadChat: vi.fn().mockResolvedValue([]),
        appendChat,
      }
      const binding = createSheetsChatBindingCoordinator({
        api,
        createTempChatId: () => 'unsaved-1',
        onBinding,
        onHistory: vi.fn(),
        onReset: vi.fn(),
      })
      await binding.bind(first)

      const rebind = binding.bind(next)
      expect(binding.current()).toBeNull()
      expect(api.rebindChat).toHaveBeenCalledWith({
        projectId: 'p',
        tempChatId: 'old',
        sessionId: next.sessionId,
      })
      binding.persist({ role: 'user', text: 'during rebind' })
      expect(appendChat).not.toHaveBeenCalled()

      rebound.resolve({ projectId: 'p', chatId: 'new' })
      await rebind
      await Promise.resolve()
      expect(binding.current()).toEqual({ projectId: 'p', chatId: 'new' })
      expect(appendChat).toHaveBeenCalledWith({
        projectId: 'p',
        chatId: 'new',
        role: 'user',
        text: 'during rebind',
      })
      expect(onBinding.mock.calls.slice(-2)).toEqual([[null], [{ projectId: 'p', chatId: 'new' }]])
    },
  )

  it.each(['null', 'reject'] as const)(
    'falls back to the authoritative saved target after rebind %s',
    async (outcome) => {
      const appendChat = vi.fn().mockResolvedValue(undefined)
      const api = {
        resolveChat: vi
          .fn()
          .mockResolvedValueOnce({ projectId: 'p', chatId: 'old' })
          .mockResolvedValueOnce({ projectId: 'p', chatId: 'saved' }),
        rebindChat:
          outcome === 'null'
            ? vi.fn().mockResolvedValue(null)
            : vi.fn().mockRejectedValue(new Error('failed')),
        loadChat: vi.fn().mockResolvedValue([]),
        appendChat,
      }
      const binding = createSheetsChatBindingCoordinator({
        api,
        createTempChatId: () => 'unsaved-1',
        onBinding: vi.fn(),
        onHistory: vi.fn(),
        onReset: vi.fn(),
      })
      await binding.bind(target('doc-a', null, 'session-1'))
      const pending = binding.bind(target('doc-a', '/saved.xlsx', 'session-2'))
      binding.persist({ role: 'assistant', text: 'queued' })
      await pending
      await Promise.resolve()

      expect(api.resolveChat).toHaveBeenLastCalledWith({
        filePath: '/saved.xlsx',
        sessionId: 'session-2',
        tempChatId: 'unsaved-1',
      })
      expect(appendChat).toHaveBeenCalledWith({
        projectId: 'p',
        chatId: 'saved',
        role: 'assistant',
        text: 'queued',
      })
    },
  )

  it('stays suspended and queued when rebind and fallback both fail', async () => {
    const appendChat = vi.fn()
    const api = {
      resolveChat: vi
        .fn()
        .mockResolvedValueOnce({ projectId: 'p', chatId: 'old' })
        .mockRejectedValueOnce(new Error('no authoritative path')),
      rebindChat: vi.fn().mockRejectedValue(new Error('failed')),
      loadChat: vi.fn().mockResolvedValue([]),
      appendChat,
    }
    const binding = createSheetsChatBindingCoordinator({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding: vi.fn(),
      onHistory: vi.fn(),
      onReset: vi.fn(),
    })
    await binding.bind(target('doc-a', null, 'session-1'))
    const pending = binding.bind(target('doc-a', '/saved.xlsx', 'session-2'))
    binding.persist({ role: 'user', text: 'never old' })
    await pending
    expect(binding.current()).toBeNull()
    expect(appendChat).not.toHaveBeenCalled()
  })

  it('drops stale rebind work and queued messages after document switch or dispose', async () => {
    const rebound = deferred<{ projectId: string; chatId: string } | null>()
    const appendChat = vi.fn().mockResolvedValue(undefined)
    const api = {
      resolveChat: vi.fn(({ filePath }: { filePath: string | null }) =>
        Promise.resolve({ projectId: 'p', chatId: filePath === '/b.xlsx' ? 'B' : 'A' }),
      ),
      rebindChat: vi.fn().mockReturnValue(rebound.promise),
      loadChat: vi.fn().mockResolvedValue([]),
      appendChat,
    }
    const binding = createSheetsChatBindingCoordinator({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding: vi.fn(),
      onHistory: vi.fn(),
      onReset: vi.fn(),
    })
    await binding.bind(target('doc-a', '/a.xlsx', 'session-1'))
    void binding.bind(target('doc-a', '/a.xlsx', 'session-2'))
    binding.persist({ role: 'user', text: 'stale queued' })
    await binding.bind(target('doc-b', '/b.xlsx', 'session-3'))
    rebound.resolve({ projectId: 'p', chatId: 'late-A' })
    await rebound.promise
    await Promise.resolve()
    expect(binding.current()).toEqual({ projectId: 'p', chatId: 'B' })
    expect(JSON.stringify(appendChat.mock.calls)).not.toContain('stale queued')

    binding.dispose()
    binding.persist({ role: 'user', text: 'after dispose' })
    expect(binding.current()).toBeNull()
    expect(JSON.stringify(appendChat.mock.calls)).not.toContain('after dispose')
  })
})
