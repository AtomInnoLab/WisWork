import { describe, expect, it, vi } from 'vitest'
import { createChatBindingCoordinator } from '../src/renderer/ai/chat-binding'

const deferred = <T>() => {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((r) => (resolve = r))
  return { promise, resolve }
}

describe('chat binding coordinator', () => {
  it('drops stale A resolution and loads only the latest B history', async () => {
    const a = deferred<{ projectId: string; chatId: string }>()
    const histories: string[] = []
    const api = {
      resolveChat: vi.fn(({ filePath }: { filePath: string | null }) =>
        filePath === '/A.docx' ? a.promise : Promise.resolve({ projectId: 'p', chatId: 'B' }),
      ),
      rebindChat: vi.fn(),
      loadChat: vi.fn(({ chatId }: { chatId: string }) => Promise.resolve([chatId])),
    }
    const bindings: Array<string | null> = []
    const coordinator = createChatBindingCoordinator<string>({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding: (ids) => bindings.push(ids?.chatId ?? null),
      onHistory: (messages) => histories.push(...messages),
      onReset: vi.fn(),
    })

    void coordinator.bind('/A.docx')
    await coordinator.bind('/B.docx')
    a.resolve({ projectId: 'p', chatId: 'A' })
    await a.promise
    await Promise.resolve()

    expect(api.loadChat).toHaveBeenCalledTimes(1)
    expect(api.loadChat).toHaveBeenCalledWith({ projectId: 'p', chatId: 'B', limit: 200 })
    expect(bindings.at(-1)).toBe('B')
    expect(histories).toEqual(['B'])
  })

  it('rebinds an untitled chat without resetting or loading another transcript', async () => {
    const reset = vi.fn()
    const api = {
      resolveChat: vi.fn().mockResolvedValue({ projectId: 'p', chatId: 'unsaved-1' }),
      rebindChat: vi.fn().mockResolvedValue({ projectId: 'p', chatId: 'saved' }),
      loadChat: vi.fn().mockResolvedValue([]),
    }
    const coordinator = createChatBindingCoordinator({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding: vi.fn(),
      onHistory: vi.fn(),
      onReset: reset,
    })
    await coordinator.bind(null)
    await coordinator.bind('/first.docx')

    expect(api.rebindChat).toHaveBeenCalledWith({
      projectId: 'p',
      tempChatId: 'unsaved-1',
      newFilePath: '/first.docx',
    })
    expect(api.resolveChat).toHaveBeenCalledTimes(1)
    expect(api.loadChat).toHaveBeenCalledTimes(1)
    expect(reset).not.toHaveBeenCalled()
    expect(coordinator.current()).toEqual({ projectId: 'p', chatId: 'saved' })
  })

  it('invalidates binding immediately on A to B and ignores work after dispose', async () => {
    const b = deferred<{ projectId: string; chatId: string }>()
    const onBinding = vi.fn()
    const reset = vi.fn()
    const api = {
      resolveChat: vi.fn(({ filePath }: { filePath: string | null }) =>
        filePath === '/B.docx' ? b.promise : Promise.resolve({ projectId: 'p', chatId: 'A' }),
      ),
      rebindChat: vi.fn(),
      loadChat: vi.fn().mockResolvedValue([]),
    }
    const coordinator = createChatBindingCoordinator({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding,
      onHistory: vi.fn(),
      onReset: reset,
    })
    await coordinator.bind('/A.docx')
    void coordinator.bind('/B.docx')
    expect(coordinator.current()).toBeNull()
    expect(reset).toHaveBeenCalledTimes(1)
    coordinator.dispose()
    b.resolve({ projectId: 'p', chatId: 'B' })
    await b.promise
    await Promise.resolve()
    expect(coordinator.current()).toBeNull()
    expect(onBinding).not.toHaveBeenCalledWith({ projectId: 'p', chatId: 'B' })
  })
})
