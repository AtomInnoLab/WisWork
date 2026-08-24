import { describe, expect, it, vi } from 'vitest'
import { createChatBindingCoordinator } from '../src/renderer/ai/chat-binding'

const deferred = <T>() => {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((r) => (resolve = r))
  return { promise, resolve }
}

describe('markdown chat binding coordinator', () => {
  it('cancels stale resolve/load across concurrent path changes', async () => {
    const stale = deferred<{ projectId: string; chatId: string }>()
    const api = {
      resolveChat: vi.fn(({ filePath }: { filePath: string | null }) =>
        filePath === '/A.md' ? stale.promise : Promise.resolve({ projectId: 'p', chatId: 'B' }),
      ),
      rebindChat: vi.fn(),
      loadChat: vi.fn(({ chatId }: { chatId: string }) => Promise.resolve([chatId])),
    }
    const history = vi.fn()
    const binding = createChatBindingCoordinator<string>({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding: vi.fn(),
      onHistory: history,
      onReset: vi.fn(),
    })
    void binding.bind('/A.md')
    await binding.bind('/B.md')
    stale.resolve({ projectId: 'p', chatId: 'A' })
    await stale.promise
    await Promise.resolve()
    expect(api.loadChat).toHaveBeenCalledTimes(1)
    expect(binding.current()?.chatId).toBe('B')
    expect(history).toHaveBeenCalledWith(['B'])
  })

  it('keeps untitled context on first save and invalidates A immediately on Save As', async () => {
    const reset = vi.fn()
    const api = {
      resolveChat: vi.fn(({ filePath }: { filePath: string | null }) =>
        Promise.resolve({ projectId: 'p', chatId: filePath ? filePath : 'unsaved-1' }),
      ),
      rebindChat: vi.fn().mockResolvedValue({ projectId: 'p', chatId: 'first' }),
      loadChat: vi.fn().mockResolvedValue([]),
    }
    const binding = createChatBindingCoordinator({
      api,
      createTempChatId: () => 'unsaved-1',
      onBinding: vi.fn(),
      onHistory: vi.fn(),
      onReset: reset,
    })
    await binding.bind(null)
    await binding.bind('/first.md')
    expect(reset).not.toHaveBeenCalled()
    expect(binding.current()?.chatId).toBe('first')
    const saveAs = binding.bind('/B.md')
    expect(binding.current()).toBeNull()
    await saveAs
    expect(reset).toHaveBeenCalledTimes(1)
    expect(binding.current()?.chatId).toBe('/B.md')
  })

  it('ignores a load that finishes after unmount', async () => {
    const load = deferred<string[]>()
    const history = vi.fn()
    const binding = createChatBindingCoordinator<string>({
      api: {
        resolveChat: vi.fn().mockResolvedValue({ projectId: 'p', chatId: 'A' }),
        rebindChat: vi.fn(),
        loadChat: vi.fn(() => load.promise),
      },
      createTempChatId: () => 'unsaved-1',
      onBinding: vi.fn(),
      onHistory: history,
      onReset: vi.fn(),
    })
    const pending = binding.bind('/A.md')
    await Promise.resolve()
    binding.dispose()
    load.resolve(['stale'])
    await pending
    expect(history).not.toHaveBeenCalled()
    expect(binding.current()).toBeNull()
  })
})
