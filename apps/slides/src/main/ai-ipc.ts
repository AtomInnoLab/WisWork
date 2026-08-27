/**
 * AI IPC for the slides main process, extracted from slides-main.ts:
 * settings persistence, the streaming proxy (main process does the networking
 * to avoid renderer CORS), search tools, and the slides-only ai:* channels
 * (image generation, media analysis, style templates).
 */
import { app, ipcMain } from 'electron'
import type { IpcMainInvokeEvent } from 'electron'
import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { AiIpcError, registerWisworkModelIpc, validateAiSearchArgs } from '@wiswork/ai-provider'
import { AuthError, getElectronAuthRuntimeOrNull } from '@wiswork/auth'
import { webSearch, imageSearch } from '@wiswork/ai-search'
import { fetchWithSsrfGuard } from '@wiswork/electron-utils'
import { addPicture, replacePictureBytes } from '@wiswork/pptx-engine'
import { parsePresentationTransaction } from '@wiswork/presentation-ops'
import { EMU_PER_PX_96 } from '@wiswork/pptx-render'
import { tm } from './i18n-main'
import {
  acquirePresentationTransactionLease,
  pushHistory,
  rebuildSlide,
  scheduleHistoryNotify,
  sessions,
} from './session-state'
import { registerUnsupportedMediaIpc } from './unsupported-ipc'
import { DesktopPresentationHost } from './operations/desktop-host'
import { PresentationTransactionExecutor } from './operations/executor'

// ---- AI settings + streaming proxy (the main process does the networking to avoid renderer CORS; implementation shared via @wiswork/ai-provider) ----

const AI_SETTINGS_PATH = () => join(app.getPath('userData'), 'ai-settings.json')

function readJson<T>(path: string, fallback: T): T {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8')) as T
  } catch {
    /* Corrupted state file: fall back to defaults */
  }
  return fallback
}

function writeJson(path: string, value: unknown): void {
  mkdirSync(join(path, '..'), { recursive: true })
  writeFileSync(path, JSON.stringify(value, null, 2))
}

function assertAuthIpc(event: IpcMainInvokeEvent, args: readonly unknown[]): void {
  if (!sessions.has(event.sender.id)) throw new Error('Untrusted IPC sender.')
  if (args.length !== 0) throw new Error('Invalid auth IPC payload.')
}

function assertAiIpcSender(event: IpcMainInvokeEvent): void {
  if (!sessions.has(event.sender.id)) throw new AiIpcError('untrusted_sender')
}

function validateSlidesAiObject(
  value: unknown,
  allowed: readonly string[],
  maxChars = 2_000_000,
): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new AiIpcError('invalid_payload')
  }
  const object = value as Record<string, unknown>
  if (Object.keys(object).some((key) => !allowed.includes(key))) {
    throw new AiIpcError('invalid_payload')
  }
  let encoded: string
  try {
    encoded = JSON.stringify(object)
  } catch {
    throw new AiIpcError('invalid_payload')
  }
  if (encoded.length > maxChars) throw new AiIpcError('payload_too_large')
  return object
}

function validateSlidesAiString(value: unknown, maxChars: number): string {
  if (typeof value !== 'string') throw new AiIpcError('invalid_payload')
  if (value.length > maxChars) throw new AiIpcError('payload_too_large')
  return value
}

export function registerAiIpc(): void {
  registerWisworkModelIpc({
    ipcMain,
    channels: {
      getSettings: 'ai:get-settings',
      setSettings: 'ai:set-settings',
      stream: 'ai:stream',
      streamChunk: 'ai:stream-chunk',
      cancel: 'ai:stream-cancel',
    },
    isTrustedSender: (senderId) => sessions.has(senderId),
    loadSettings: () => readJson(AI_SETTINGS_PATH(), {}),
    saveSettings: (settings) => writeJson(AI_SETTINGS_PATH(), settings),
    getAccessToken: async () => {
      const runtime = getElectronAuthRuntimeOrNull()
      return runtime?.client.getAccessToken() ?? null
    },
    fetchWithAuth: (request) => {
      const runtime = getElectronAuthRuntimeOrNull()
      if (!runtime) throw new AuthError('auth_required')
      return runtime.client.fetchWithAuth(request)
    },
  })

  ipcMain.handle('auth:status', (event, ...args: unknown[]) => {
    assertAuthIpc(event, args)
    return getElectronAuthRuntimeOrNull()?.client.getValidAccountStatus() ?? { loggedIn: false }
  })
  ipcMain.handle('auth:login', (event, ...args: unknown[]) => {
    assertAuthIpc(event, args)
    const runtime = getElectronAuthRuntimeOrNull()
    if (!runtime) throw new AuthError('auth_unavailable_in_standalone')
    return runtime.beginLogin()
  })
  ipcMain.handle('auth:logout', (event, ...args: unknown[]) => {
    assertAuthIpc(event, args)
    return getElectronAuthRuntimeOrNull()?.client.logout()
  })

  // Search tools (content + images), Serper with DuckDuckGo fallback
  ipcMain.handle('ai:web-search', async (event, query: string, maxResults?: number) => {
    assertAiIpcSender(event)
    try {
      const input = validateAiSearchArgs(query, maxResults, 6)
      return await webSearch(input.query, input.maxResults)
    } catch (err) {
      return { results: [], method: 'error', error: String(err) }
    }
  })

  ipcMain.handle('ai:image-search', async (event, query: string, maxResults?: number) => {
    assertAiIpcSender(event)
    try {
      const input = validateAiSearchArgs(query, maxResults, 8)
      return await imageSearch(input.query, input.maxResults)
    } catch (err) {
      return { images: [], method: 'error', error: String(err) }
    }
  })
}

// ── ai:* handlers unique to slides ──────────────────────────────────────
// Must be registered inside registerSlidesIpc (not registerAiIpc): in shell aggregate mode the
// generic ai:* channels are registered by docs-main.registerAiIpc, and slides' registerAiIpc is
// never called; docs does not have these channels, so putting them in the wrong place raises
// "No handler registered".
export function registerSlidesOnlyAiIpc(): void {
  registerUnsupportedMediaIpc(ipcMain, (senderId) => sessions.has(senderId))

  const transactionExecutors = new WeakMap<
    NonNullable<ReturnType<typeof sessions.get>>,
    PresentationTransactionExecutor<Awaited<ReturnType<DesktopPresentationHost['captureSnapshot']>>>
  >()
  const transactionControllers = new Map<string, Set<AbortController>>()
  const controllerKey = (senderId: number, transactionId: string) => `${senderId}:${transactionId}`

  ipcMain.handle('slides:presentation-transaction', async (event, value: unknown) => {
    assertAiIpcSender(event)
    const transaction = parsePresentationTransaction(value)
    const session = sessions.get(event.sender.id)!
    if (session.masterEdit || session.historyBatch || session.transformPreview) {
      throw new AiIpcError('invalid_payload')
    }
    let executor = transactionExecutors.get(session)
    if (!executor) {
      executor = new PresentationTransactionExecutor(new DesktopPresentationHost(session))
      transactionExecutors.set(session, executor)
    }
    const key = controllerKey(event.sender.id, transaction.transactionId)
    const controller = new AbortController()
    const releaseLease = acquirePresentationTransactionLease(session)
    if (!releaseLease) {
      return {
        status: 'unchanged',
        transactionId: transaction.transactionId,
        code: 'write_not_applied',
        operationCount: transaction.operations.length,
      }
    }
    const controllers = transactionControllers.get(key) ?? new Set<AbortController>()
    controllers.add(controller)
    transactionControllers.set(key, controllers)
    try {
      return await executor.execute(transaction, controller.signal)
    } finally {
      releaseLease()
      controllers.delete(controller)
      if (controllers.size === 0) transactionControllers.delete(key)
    }
  })

  ipcMain.handle(
    'slides:presentation-transaction-cancel',
    (event, transactionId: unknown): boolean => {
      assertAiIpcSender(event)
      if (
        typeof transactionId !== 'string' ||
        transactionId.length === 0 ||
        transactionId.length > 128 ||
        !/^[A-Za-z0-9][A-Za-z0-9._:-]*$/.test(transactionId)
      ) {
        throw new AiIpcError('invalid_payload')
      }
      const controllers = transactionControllers.get(controllerKey(event.sender.id, transactionId))
      if (!controllers?.size) return false
      for (const controller of controllers) controller.abort()
      return true
    },
  )

  // Download an image from a URL and insert it into the given page (image search -> insert in one step; download in the main process avoids CORS)
  ipcMain.handle(
    'ai:insert-image-url',
    async (
      e,
      op: {
        slideIndex: number
        url: string
        xPx: number
        yPx: number
        wPx: number
        hPx: number
        fitWidthPx: number
      },
    ) => {
      assertAiIpcSender(e)
      validateSlidesAiObject(op, ['slideIndex', 'url', 'xPx', 'yPx', 'wPx', 'hPx', 'fitWidthPx'])
      validateSlidesAiString(op.url, 4_096)
      if (!Number.isInteger(op.slideIndex) || op.slideIndex < 0 || op.slideIndex > 10_000)
        throw new AiIpcError('invalid_payload')
      for (const value of [op.xPx, op.yPx, op.wPx, op.hPx, op.fitWidthPx]) {
        if (typeof value !== 'number' || !Number.isFinite(value) || Math.abs(value) > 1_000_000)
          throw new AiIpcError('invalid_payload')
      }
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const slide = session.opened.deck.slides[op.slideIndex]
      if (!slide) return null
      try {
        // the URL originates from AI tool calls (prompt-injectable via image
        // search results), so refuse non-http schemes and private/link-local
        // targets; redirects are followed manually so every hop is validated
        const resp = await fetchWithSsrfGuard(String(op.url), {
          headers: { 'User-Agent': 'Mozilla/5.0' },
        })
        if (!resp || !resp.ok) return null
        const buf = Buffer.from(await resp.arrayBuffer())
        const ct = resp.headers.get('content-type') ?? ''
        const ext = ct.includes('png') ? 'png' : ct.includes('gif') ? 'gif' : 'jpg'
        const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
        const scale = op.fitWidthPx / baseWidthPx
        const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
        pushHistory(session)
        const el = addPicture(session.opened, slide, {
          bytes: new Uint8Array(buf),
          ext,
          offset: {
            x: toEmu(op.xPx),
            y: toEmu(op.yPx),
            cx: Math.max(1, toEmu(op.wPx)),
            cy: Math.max(1, toEmu(op.hPx)),
          },
        })
        if (!el) {
          session.undoStack.pop()
          scheduleHistoryNotify(session)
          return null
        }
        session.fitWidthPx = op.fitWidthPx
        const rebuilt = rebuildSlide(session, op.slideIndex)
        return rebuilt ? { slide: rebuilt, sourceId: el.id } : null
      } catch {
        return null
      }
    },
  )

  // Download an image from a URL and swap it into an existing picture in place
  // (frame/z-order/effects survive). Same URL hardening as ai:insert-image-url.
  ipcMain.handle(
    'ai:replace-picture-url',
    async (e, op: { slideIndex: number; sourceId: string; url: string; keepSrcRect?: boolean }) => {
      assertAiIpcSender(e)
      validateSlidesAiObject(op, ['slideIndex', 'sourceId', 'url', 'keepSrcRect'])
      validateSlidesAiString(op.sourceId, 4_096)
      validateSlidesAiString(op.url, 4_096)
      if (!Number.isInteger(op.slideIndex) || op.slideIndex < 0 || op.slideIndex > 10_000)
        throw new AiIpcError('invalid_payload')
      if (op.keepSrcRect !== undefined && typeof op.keepSrcRect !== 'boolean')
        throw new AiIpcError('invalid_payload')
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const slide = session.opened.deck.slides[op.slideIndex]
      if (!slide) return null
      try {
        const resp = await fetchWithSsrfGuard(op.url, {
          headers: { 'User-Agent': 'Mozilla/5.0' },
        })
        if (!resp || !resp.ok) return null
        const buf = Buffer.from(await resp.arrayBuffer())
        const ct = resp.headers.get('content-type') ?? ''
        const ext = ct.includes('png') ? 'png' : ct.includes('gif') ? 'gif' : 'jpg'
        pushHistory(session)
        const ok = replacePictureBytes(
          session.opened,
          slide,
          String(op.sourceId),
          new Uint8Array(buf),
          ext,
          op.keepSrcRect ? { keepSrcRect: true } : undefined,
        )
        if (!ok) {
          session.undoStack.pop()
          scheduleHistoryNotify(session)
          return null
        }
        return rebuildSlide(session, op.slideIndex)
      } catch {
        return null
      }
    },
  )

  // ── Style Skill sidecar persistence: write a same-named .styleskill.json next to the draft (fail-open)
  ipcMain.handle(
    'ai:save-sidecar',
    async (
      event,
      data: { topic: string; styleSkill: string; createdAt: string },
    ): Promise<{ ok: boolean }> => {
      assertAiIpcSender(event)
      validateSlidesAiObject(data, ['topic', 'styleSkill', 'createdAt'])
      validateSlidesAiString(data.topic, 100_000)
      validateSlidesAiString(data.styleSkill, 1_000_000)
      validateSlidesAiString(data.createdAt, 128)
      try {
        const session = sessions.get(event.sender.id)
        const draftPath = session?.path
        if (!draftPath || !draftPath.endsWith('.pptx')) return { ok: false }
        const sidecarPath = draftPath.replace(/\.pptx$/i, '.styleskill.json')
        writeFileSync(sidecarPath, JSON.stringify(data, null, 2))
        return { ok: true }
      } catch {
        return { ok: false }
      }
    },
  )

  // ── Style template save: stored in userData/style-templates/<name>.json
  const STYLE_TEMPLATES_DIR = () => join(app.getPath('userData'), 'style-templates')

  ipcMain.handle(
    'ai:save-style-template',
    (
      event,
      name: string,
      data: { topic: string; styleSkill: string; createdAt: string },
    ): { ok: boolean; error?: string } => {
      assertAiIpcSender(event)
      validateSlidesAiString(name, 256)
      validateSlidesAiObject(data, ['topic', 'styleSkill', 'createdAt'])
      validateSlidesAiString(data.topic, 100_000)
      validateSlidesAiString(data.styleSkill, 1_000_000)
      validateSlidesAiString(data.createdAt, 128)
      try {
        const dir = STYLE_TEMPLATES_DIR()
        mkdirSync(dir, { recursive: true })
        // Filename: replace illegal characters in the name with _ then truncate to 64 chars
        const safeName = name.replace(/[/\\:*?"<>|]/g, '_').slice(0, 64)
        if (!safeName) return { ok: false, error: tm('errTplNameInvalid') }
        writeJson(join(dir, `${safeName}.json`), { ...data, name: safeName })
        return { ok: true }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )

  // ── Style template list
  ipcMain.handle(
    'ai:list-style-templates',
    (event, ...args: unknown[]): Array<{ name: string; topic: string; createdAt: string }> => {
      assertAiIpcSender(event)
      if (args.length !== 0) throw new AiIpcError('invalid_payload')
      try {
        const dir = STYLE_TEMPLATES_DIR()
        if (!existsSync(dir)) return []
        const files = readdirSync(dir).filter((f) => f.endsWith('.json'))
        return files
          .map((f) => {
            try {
              const raw = readJson<{
                name?: string
                topic?: string
                createdAt?: string
                styleSkill?: string
              }>(join(dir, f), {})
              return {
                name: raw.name ?? f.replace(/\.json$/, ''),
                topic: raw.topic ?? '',
                createdAt: raw.createdAt ?? '',
              }
            } catch {
              return null
            }
          })
          .filter(Boolean) as Array<{ name: string; topic: string; createdAt: string }>
      } catch {
        return []
      }
    },
  )

  // ── Style template load
  ipcMain.handle(
    'ai:load-style-template',
    (event, name: string): { ok: boolean; styleSkill?: string; topic?: string; error?: string } => {
      assertAiIpcSender(event)
      validateSlidesAiString(name, 256)
      try {
        const dir = STYLE_TEMPLATES_DIR()
        const safeName = name.replace(/[/\\:*?"<>|]/g, '_').slice(0, 64)
        const filePath = join(dir, `${safeName}.json`)
        if (!existsSync(filePath)) return { ok: false, error: tm('errTplMissing', { name }) }
        const raw = readJson<{ styleSkill?: string; topic?: string }>(filePath, {})
        if (!raw.styleSkill) return { ok: false, error: tm('errTplNoSkill', { name }) }
        return { ok: true, styleSkill: raw.styleSkill, topic: raw.topic ?? '' }
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) }
      }
    },
  )
}
