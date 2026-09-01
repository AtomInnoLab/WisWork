import type { ShellCodexRuntime, CodexOwner } from './codex-runtime'
import { CODEX_RUNTIME_CHANNELS, type CodexRuntimeStartRequest } from '../shared/codex-api'

interface IpcMainLike {
  handle(
    channel: string,
    handler: (event: { sender: CodexOwner }, ...args: unknown[]) => unknown,
  ): void
}

export function registerCodexRuntimeIpc(options: {
  readonly ipcMain: IpcMainLike
  readonly runtime: ShellCodexRuntime
  readonly documentIdForOwner: (owner: CodexOwner) => string | null
}): void {
  const exactStartRequest = (value: unknown): CodexRuntimeStartRequest => {
    if (typeof value !== 'object' || value === null || Array.isArray(value)) {
      throw new Error('enhanced_invalid_request')
    }
    if (Object.getPrototypeOf(value) !== Object.prototype) {
      throw new Error('enhanced_invalid_request')
    }
    const descriptors = Object.getOwnPropertyDescriptors(value)
    if (
      Object.getOwnPropertySymbols(value).length !== 0 ||
      Object.keys(descriptors).length !== 2 ||
      !('value' in (descriptors.documentId ?? {})) ||
      !('value' in (descriptors.text ?? {}))
    ) {
      throw new Error('enhanced_invalid_request')
    }
    return {
      documentId: descriptors.documentId!.value as string,
      text: descriptors.text!.value as string,
    }
  }
  const authoritativeDocument = (owner: CodexOwner): string => {
    if (owner.isDestroyed()) throw new Error('enhanced_untrusted_request')
    const documentId = options.documentIdForOwner(owner)
    if (!documentId || !options.runtime.ownsDocument(owner, documentId)) {
      throw new Error('enhanced_untrusted_request')
    }
    return documentId
  }
  options.ipcMain.handle(CODEX_RUNTIME_CHANNELS.status, (event, ...args) => {
    if (args.length !== 0) throw new Error('enhanced_invalid_request')
    const documentId = options.documentIdForOwner(event.sender)
    return {
      activeAgentRuntime: options.runtime.activeAgentRuntime,
      state: options.runtime.state,
      documentId:
        documentId && options.runtime.ownsDocument(event.sender, documentId) ? documentId : null,
    }
  })
  options.ipcMain.handle(CODEX_RUNTIME_CHANNELS.startTurn, async (event, ...args) => {
    if (args.length !== 1) throw new Error('enhanced_invalid_request')
    const request = exactStartRequest(args[0])
    const documentId = authoritativeDocument(event.sender)
    if (request.documentId !== documentId) {
      throw new Error('enhanced_untrusted_request')
    }
    await options.runtime.startTurn(event.sender, documentId, request.text)
  })
  options.ipcMain.handle(CODEX_RUNTIME_CHANNELS.cancelTurn, async (event, ...args) => {
    const documentId = authoritativeDocument(event.sender)
    if (args.length !== 1 || args[0] !== documentId) throw new Error('enhanced_untrusted_request')
    await options.runtime.cancelTurn(event.sender, documentId)
  })
}
