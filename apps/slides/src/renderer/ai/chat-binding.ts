import type { AppendChatArgs, ChatMessage } from '@wiswork/project-store'

export interface SlidesChatIds {
  projectId: string
  chatId: string
}

interface SlidesChatApi {
  resolveChat(args: { filePath: string | null; tempChatId: string }): Promise<SlidesChatIds>
  rebindChat(args: {
    projectId: string
    tempChatId: string
    newFilePath: string
  }): Promise<SlidesChatIds | null | undefined>
  loadChat(args: { projectId: string; chatId: string; limit: number }): Promise<ChatMessage[]>
  appendChat(args: AppendChatArgs): Promise<void>
}

interface SlidesChatBindingOptions {
  api: SlidesChatApi
  createTempChatId: () => string
  onBinding: (ids: SlidesChatIds | null) => void
  onHistory: (messages: ChatMessage[]) => void
  onReset: () => void
}

type PendingMessage = Omit<AppendChatArgs, 'projectId' | 'chatId'>

/** Owns chat persistence across StrictMode replay and draft-to-file rebinding. */
export function createSlidesChatBindingCoordinator({
  api,
  createTempChatId,
  onBinding,
  onHistory,
  onReset,
}: SlidesChatBindingOptions) {
  let generation = 0
  let path: string | null | undefined
  let binding: SlidesChatIds | null = null
  let pending: PendingMessage[] = []
  let active = true

  const isCurrent = (token: number) => active && token === generation
  const publish = (ids: SlidesChatIds | null) => {
    binding = ids
    onBinding(ids)
    if (!ids || pending.length === 0) return
    const messages = pending
    pending = []
    for (const message of messages) {
      void api.appendChat({ ...ids, ...message }).catch(() => undefined)
    }
  }

  const resolve = (nextPath: string | null) =>
    api.resolveChat({ filePath: nextPath, tempChatId: createTempChatId() })

  const bind = async (nextPath: string | null): Promise<void> => {
    active = true
    if (path === nextPath) return
    const previousPath = path
    const previousBinding = binding
    path = nextPath
    const token = ++generation

    if (previousPath !== undefined) {
      publish(null)
      if (previousBinding && nextPath !== null) {
        let rebound: SlidesChatIds | null | undefined
        try {
          rebound = await api.rebindChat({
            projectId: previousBinding.projectId,
            tempChatId: previousBinding.chatId,
            newFilePath: nextPath,
          })
        } catch {
          // Resolve the saved path authoritatively below.
        }
        if (!isCurrent(token)) return
        if (rebound) {
          publish(rebound)
          return
        }
      }
      try {
        const ids = await resolve(nextPath)
        if (isCurrent(token)) publish(ids)
      } catch {
        // Fail closed: keep persistence suspended and pending.
      }
      return
    }

    try {
      const ids = await resolve(nextPath)
      if (!isCurrent(token)) return
      publish(ids)
      const messages = await api.loadChat({ ...ids, limit: 200 })
      if (isCurrent(token)) onHistory(messages)
    } catch {
      // History is best effort; a later path transition retries persistence.
    }
  }

  return {
    bind,
    current: () => binding,
    persist: (message: PendingMessage) => {
      if (!active) return
      if (!binding) {
        pending.push(message)
        return
      }
      void api.appendChat({ ...binding, ...message }).catch(() => undefined)
    },
    deactivate: () => {
      active = false
      generation += 1
      path = undefined
      pending = []
      publish(null)
    },
    reset: () => {
      active = false
      generation += 1
      path = undefined
      pending = []
      publish(null)
      onReset()
    },
  }
}
