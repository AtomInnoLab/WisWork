import { describe, expect, it, vi } from 'vitest'
import { createSlidesChatBindingCoordinator } from '../src/renderer/ai/chat-binding'

const deferred = <T>() => {
  let resolve!: (value: T) => void
  let reject!: (reason: unknown) => void
  const promise = new Promise<T>((yes, no) => {
    resolve = yes
    reject = no
  })
  return { promise, resolve, reject }
}

describe('Slides chat binding coordinator', () => {
  it('suspends first-save persistence and flushes queued messages only to rebound IDs', async () => {
    const moved = deferred<{ projectId: string; chatId: string } | null>()
    const appendChat = vi.fn().mockResolvedValue(undefined)
    const onBinding = vi.fn()
    const api = {
      resolveChat: vi.fn().mockResolvedValue({ projectId: 'p', chatId: 'unsaved-1' }),
      rebindChat: vi.fn().mockReturnValue(moved.promise),
      loadChat: vi.fn().mockResolvedValue([]),
      appendChat,
    }
    const binding = createSlidesChatBindingCoordinator({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding,
      onHistory: vi.fn(),
      onReset: vi.fn(),
    })
    await binding.bind(null)
    const save = binding.bind('/saved.pptx')
    expect(binding.current()).toBeNull()
    binding.persist({ role: 'user', text: 'queued' })
    expect(appendChat).not.toHaveBeenCalled()
    moved.resolve({ projectId: 'p', chatId: 'saved' })
    await save
    expect(appendChat).toHaveBeenCalledWith({
      projectId: 'p',
      chatId: 'saved',
      role: 'user',
      text: 'queued',
    })
    expect(onBinding.mock.calls.slice(-2)).toEqual([[null], [{ projectId: 'p', chatId: 'saved' }]])
  })

  it.each(['null', 'reject'] as const)(
    'uses authoritative resolve after rebind %s',
    async (kind) => {
      const appendChat = vi.fn().mockResolvedValue(undefined)
      const api = {
        resolveChat: vi
          .fn()
          .mockResolvedValueOnce({ projectId: 'p', chatId: 'unsaved-1' })
          .mockResolvedValueOnce({ projectId: 'p', chatId: 'saved' }),
        rebindChat:
          kind === 'null'
            ? vi.fn().mockResolvedValue(null)
            : vi.fn().mockRejectedValue(new Error('move failed')),
        loadChat: vi.fn().mockResolvedValue([]),
        appendChat,
      }
      const binding = createSlidesChatBindingCoordinator({
        api,
        createTempChatId: () => 'unsaved-1',
        onBinding: vi.fn(),
        onHistory: vi.fn(),
        onReset: vi.fn(),
      })
      await binding.bind(null)
      const save = binding.bind('/saved.pptx')
      binding.persist({ role: 'assistant', text: 'queued' })
      await save
      expect(api.resolveChat).toHaveBeenLastCalledWith({
        filePath: '/saved.pptx',
        tempChatId: 'unsaved-1',
      })
      expect(appendChat.mock.calls[0]?.[0]).toMatchObject({ chatId: 'saved', text: 'queued' })
    },
  )

  it('keeps messages queued if both rebind and authoritative resolve fail', async () => {
    const appendChat = vi.fn()
    const api = {
      resolveChat: vi
        .fn()
        .mockResolvedValueOnce({ projectId: 'p', chatId: 'unsaved-1' })
        .mockRejectedValueOnce(new Error('resolve failed')),
      rebindChat: vi.fn().mockRejectedValue(new Error('move failed')),
      loadChat: vi.fn().mockResolvedValue([]),
      appendChat,
    }
    const binding = createSlidesChatBindingCoordinator({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding: vi.fn(),
      onHistory: vi.fn(),
      onReset: vi.fn(),
    })
    await binding.bind(null)
    const save = binding.bind('/saved.pptx')
    binding.persist({ role: 'user', text: 'queued' })
    await save
    expect(binding.current()).toBeNull()
    expect(appendChat).not.toHaveBeenCalled()
  })

  it('drops the first StrictMode resolution and publishes only the replay', async () => {
    const first = deferred<{ projectId: string; chatId: string }>()
    const onBinding = vi.fn()
    const api = {
      resolveChat: vi
        .fn()
        .mockReturnValueOnce(first.promise)
        .mockResolvedValueOnce({ projectId: 'p', chatId: 'replay' }),
      rebindChat: vi.fn(),
      loadChat: vi.fn().mockResolvedValue([]),
      appendChat: vi.fn(),
    }
    const binding = createSlidesChatBindingCoordinator({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding,
      onHistory: vi.fn(),
      onReset: vi.fn(),
    })
    void binding.bind('/a.pptx')
    binding.deactivate()
    await binding.bind('/a.pptx')
    first.resolve({ projectId: 'p', chatId: 'stale' })
    await first.promise
    await Promise.resolve()
    expect(binding.current()).toEqual({ projectId: 'p', chatId: 'replay' })
    expect(onBinding).not.toHaveBeenCalledWith({ projectId: 'p', chatId: 'stale' })
  })

  it('invalidates path-switch and unmount work without publishing stale IDs', async () => {
    const a = deferred<{ projectId: string; chatId: string }>()
    const api = {
      resolveChat: vi.fn(({ filePath }: { filePath: string | null }) =>
        filePath === '/a.pptx' ? a.promise : Promise.resolve({ projectId: 'p', chatId: 'B' }),
      ),
      rebindChat: vi.fn(),
      loadChat: vi.fn().mockResolvedValue([]),
      appendChat: vi.fn(),
    }
    const binding = createSlidesChatBindingCoordinator({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding: vi.fn(),
      onHistory: vi.fn(),
      onReset: vi.fn(),
    })
    void binding.bind('/a.pptx')
    binding.deactivate()
    await binding.bind('/b.pptx')
    a.resolve({ projectId: 'p', chatId: 'A' })
    await a.promise
    expect(binding.current()).toEqual({ projectId: 'p', chatId: 'B' })
    binding.deactivate()
    binding.persist({ role: 'user', text: 'after unmount' })
    expect(binding.current()).toBeNull()
    expect(api.appendChat).not.toHaveBeenCalled()
  })
})
