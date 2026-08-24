export interface ChatIds {
  projectId: string
  chatId: string
}

interface ChatBindingApi<TMessage> {
  resolveChat(args: { filePath: string | null; tempChatId: string }): Promise<ChatIds>
  rebindChat(args: {
    projectId: string
    tempChatId: string
    newFilePath: string
  }): Promise<ChatIds | null | undefined>
  loadChat(args: { projectId: string; chatId: string; limit: number }): Promise<TMessage[]>
}

interface ChatBindingOptions<TMessage> {
  api: ChatBindingApi<TMessage>
  createTempChatId: () => string
  onBinding: (ids: ChatIds | null) => void
  onHistory: (messages: TMessage[]) => void
  onReset: () => void
}

/** Owns the async lifetime of a file-backed chat independently from React's mount lifetime. */
export function createChatBindingCoordinator<TMessage = unknown>({
  api,
  createTempChatId,
  onBinding,
  onHistory,
  onReset,
}: ChatBindingOptions<TMessage>) {
  let generation = 0
  let path: string | null | undefined
  let binding: ChatIds | null = null
  let disposed = false

  const isCurrent = (token: number) => !disposed && token === generation
  const publish = (ids: ChatIds | null) => {
    binding = ids
    onBinding(ids)
  }

  const bind = async (nextPath: string | null): Promise<void> => {
    if (disposed || path === nextPath) return
    const previousPath = path
    const previousBinding = binding
    path = nextPath
    const token = ++generation

    if (
      previousPath === null &&
      nextPath !== null &&
      previousBinding?.chatId.startsWith('unsaved-')
    ) {
      publish(null)
      try {
        const rebound = await api.rebindChat({
          projectId: previousBinding.projectId,
          tempChatId: previousBinding.chatId,
          newFilePath: nextPath,
        })
        if (isCurrent(token)) publish(rebound ?? previousBinding)
      } catch {
        if (!isCurrent(token)) return
        try {
          const ids = await api.resolveChat({
            filePath: nextPath,
            tempChatId: createTempChatId(),
          })
          if (isCurrent(token)) publish(ids)
        } catch {
          // Never restore the old untitled target after the document has a path.
        }
      }
      return
    }

    if (previousPath !== undefined) {
      publish(null)
      onReset()
    }

    try {
      const ids = await api.resolveChat({ filePath: nextPath, tempChatId: createTempChatId() })
      if (!isCurrent(token)) return
      publish(ids)
      const messages = await api.loadChat({ ...ids, limit: 200 })
      if (isCurrent(token)) onHistory(messages)
    } catch {
      // Persistence/history is best effort. A later path change retries it.
    }
  }

  return {
    bind,
    current: () => binding,
    dispose: () => {
      disposed = true
      generation += 1
      publish(null)
    },
  }
}
