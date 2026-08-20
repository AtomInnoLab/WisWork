import { randomUUID } from 'node:crypto'
import { LATEX_CHANNELS } from '../shared/ipc.js'

interface IpcMainLike {
  on(channel: string, handler: (event: { sender: object }, payload: unknown) => void): void
  removeListener(
    channel: string,
    handler: (event: { sender: object }, payload: unknown) => void,
  ): void
}

export interface FlushWebContents {
  id: number
  isDestroyed(): boolean
  send(channel: string, payload: unknown): void
  once(event: 'destroyed', handler: () => void): void
  removeListener(event: 'destroyed', handler: () => void): void
}

interface PendingFlush {
  contents: FlushWebContents
  holds: number
  waiters: Array<(ok: boolean) => void>
  finish(ok: boolean): void
}

function flushAck(value: unknown): { requestId: string; ok: boolean } | null {
  if (!value || typeof value !== 'object') return null
  const record = value as Record<string, unknown>
  if (
    Object.keys(record).sort().join(',') !== 'ok,requestId' ||
    typeof record.requestId !== 'string' ||
    !record.requestId ||
    typeof record.ok !== 'boolean'
  )
    return null
  return { requestId: record.requestId, ok: record.ok }
}

export class LatexEditFlushCoordinator {
  private readonly pending = new Map<string, PendingFlush>()
  private readonly frozen = new Map<
    number,
    {
      contents: FlushWebContents
      requestId: string
      destroyed: () => void
      holds: number
    }
  >()
  private readonly timeoutMs: number
  private readonly randomId: () => string
  private readonly onAck = (event: { sender: object }, payload: unknown) => {
    const ack = flushAck(payload)
    if (!ack) return
    const pending = this.pending.get(ack.requestId)
    if (!pending || event.sender !== pending.contents) return
    pending.finish(ack.ok)
  }

  constructor(
    private readonly ipcMain: IpcMainLike,
    options: { timeoutMs?: number; randomId?: () => string } = {},
  ) {
    this.timeoutMs = options.timeoutMs ?? 2_000
    this.randomId = options.randomId ?? randomUUID
    ipcMain.on(LATEX_CHANNELS.editFlushAck, this.onAck)
  }

  request(contents: FlushWebContents): Promise<boolean> {
    if (contents.isDestroyed()) return Promise.resolve(false)
    const existing = this.frozen.get(contents.id)
    if (existing?.contents === contents) {
      existing.holds += 1
      return Promise.resolve(true)
    }
    const pendingExisting = [...this.pending.values()].find((entry) => entry.contents === contents)
    if (pendingExisting) {
      pendingExisting.holds += 1
      return new Promise<boolean>((resolve) => pendingExisting.waiters.push(resolve))
    }
    let requestId = this.randomId()
    while (this.pending.has(requestId)) requestId = this.randomId()
    return new Promise<boolean>((resolve) => {
      let settled = false
      const destroyed = () => {
        this.frozen.delete(contents.id)
        finish(false)
      }
      const timer = setTimeout(() => finish(false), this.timeoutMs)
      const sendRelease = () => {
        if (contents.isDestroyed()) return
        try {
          contents.send(LATEX_CHANNELS.editFlushRelease, { requestId })
        } catch {}
      }
      const finish = (ok: boolean) => {
        if (settled) return
        settled = true
        clearTimeout(timer)
        const holds = this.pending.get(requestId)?.holds ?? 1
        this.pending.delete(requestId)
        if (ok && !contents.isDestroyed()) {
          this.frozen.set(contents.id, { contents, requestId, destroyed, holds })
        } else {
          contents.removeListener('destroyed', destroyed)
          sendRelease()
        }
        for (const waiter of pending.waiters) waiter(ok)
      }
      const pending: PendingFlush = { contents, holds: 1, waiters: [resolve], finish }
      this.pending.set(requestId, pending)
      contents.once('destroyed', destroyed)
      try {
        contents.send(LATEX_CHANNELS.editFlushRequest, { requestId })
      } catch {
        finish(false)
      }
    })
  }

  release(contents: FlushWebContents): void {
    const entry = this.frozen.get(contents.id)
    if (!entry || entry.contents !== contents) return
    entry.holds -= 1
    if (entry.holds > 0) return
    this.frozen.delete(contents.id)
    contents.removeListener('destroyed', entry.destroyed)
    if (contents.isDestroyed()) return
    try {
      contents.send(LATEX_CHANNELS.editFlushRelease, { requestId: entry.requestId })
    } catch {}
  }

  dispose(): void {
    this.ipcMain.removeListener(LATEX_CHANNELS.editFlushAck, this.onAck)
    for (const pending of [...this.pending.values()]) pending.finish(false)
    for (const entry of [...this.frozen.values()]) {
      entry.holds = 1
      this.release(entry.contents)
    }
  }
}
