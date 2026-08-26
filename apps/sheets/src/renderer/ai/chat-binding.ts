import type { AppendChatArgs, ChatMessage, ProjectApi } from '@wiswork/project-store'

export interface SheetsChatIds {
  projectId: string
  chatId: string
}

export interface SheetsChatTarget {
  documentId: string | number
  path: string | null
  sessionId: string | number
}

interface SheetsChatBindingOptions {
  api: Pick<ProjectApi, 'resolveChat' | 'rebindChat' | 'loadChat' | 'appendChat'>
  createTempChatId: () => string
  onBinding: (ids: SheetsChatIds | null) => void
  onHistory: (messages: ChatMessage[]) => void
  onReset: () => void
}

type PendingMessage = Omit<AppendChatArgs, 'projectId' | 'chatId'>

/** Coordinates workbook identity, persistence suspension, and async chat lifetimes. */
export function createSheetsChatBindingCoordinator({
  api,
  createTempChatId,
  onBinding,
  onHistory,
  onReset,
}: SheetsChatBindingOptions) {
  let generation = 0
  let target: SheetsChatTarget | undefined
  let binding: SheetsChatIds | null = null
  let pending: PendingMessage[] = []
  let disposed = false

  const isCurrent = (token: number) => !disposed && token === generation
  const publish = (ids: SheetsChatIds | null) => {
    binding = ids
    onBinding(ids)
    if (!ids || pending.length === 0) return
    const messages = pending
    pending = []
    for (const message of messages) {
      void api.appendChat({ ...ids, ...message }).catch(() => undefined)
    }
  }
  const resolve = (next: SheetsChatTarget) =>
    api.resolveChat({
      filePath: next.path,
      sessionId: String(next.sessionId),
      tempChatId: createTempChatId(),
    })

  const bind = async (next: SheetsChatTarget): Promise<void> => {
    if (
      disposed ||
      (target?.documentId === next.documentId &&
        target.path === next.path &&
        target.sessionId === next.sessionId)
    )
      return
    const previous = target
    const previousBinding = binding
    target = next
    const token = ++generation

    if (previous?.documentId === next.documentId) {
      // Fail closed before asking the store to move the JSONL target.
      publish(null)
      let rebound: SheetsChatIds | null | undefined
      if (previousBinding) {
        try {
          rebound = await api.rebindChat({
            projectId: previousBinding.projectId,
            tempChatId: previousBinding.chatId,
            sessionId: String(next.sessionId),
          })
        } catch {
          // Resolve the authoritative target below.
        }
      }
      if (!isCurrent(token)) return
      if (rebound) {
        publish(rebound)
        return
      }
      try {
        const ids = await resolve(next)
        if (isCurrent(token)) publish(ids)
      } catch {
        // Leave both the binding and pending persistence suspended.
      }
      return
    }

    if (previous) {
      publish(null)
      pending = []
      onReset()
    }
    try {
      const ids = await resolve(next)
      if (!isCurrent(token)) return
      publish(ids)
      const messages = await api.loadChat({ ...ids, limit: 200 })
      if (isCurrent(token)) onHistory(messages)
    } catch {
      // Chat persistence/history is best effort; a later target change retries it.
    }
  }

  return {
    bind,
    current: () => binding,
    persist: (message: PendingMessage) => {
      if (disposed) return
      if (!binding) {
        pending.push(message)
        return
      }
      void api.appendChat({ ...binding, ...message }).catch(() => undefined)
    },
    clear: () => {
      if (disposed) return
      generation += 1
      target = undefined
      pending = []
      publish(null)
      onReset()
    },
    dispose: () => {
      disposed = true
      generation += 1
      pending = []
      publish(null)
    },
  }
}
