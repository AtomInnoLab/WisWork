import { AiIpcError } from '@wiswork/ai-provider'

interface InvokeEventLike {
  sender: { id: number }
}

interface IpcMainLike {
  handle(channel: string, handler: (event: InvokeEventLike, ...args: unknown[]) => unknown): void
}

type IsTrustedSender = (senderId: number) => boolean

const UNSUPPORTED = Object.freeze({
  ok: false as const,
  errorCode: 'unsupported_feature' as const,
  error: 'This feature is not available in the current WisWork development version.',
})

function plainObject(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new AiIpcError('invalid_payload')
  }
  const prototype = Object.getPrototypeOf(value)
  if (prototype !== Object.prototype && prototype !== null) {
    throw new AiIpcError('invalid_payload')
  }
  return value as Record<string, unknown>
}

function exactKeys(object: Record<string, unknown>, allowed: readonly string[]): void {
  if (Object.keys(object).some((key) => !allowed.includes(key))) {
    throw new AiIpcError('invalid_payload')
  }
}

function boundedString(value: unknown, maxChars: number, required = false): string | undefined {
  if (value === undefined && !required) return undefined
  if (typeof value !== 'string') throw new AiIpcError('invalid_payload')
  if (required && value.trim().length === 0) throw new AiIpcError('invalid_payload')
  if (value.length > maxChars) throw new AiIpcError('payload_too_large')
  return value
}

function boundedJson(value: unknown, maxChars: number): void {
  let encoded: string
  try {
    encoded = JSON.stringify(value)
  } catch {
    throw new AiIpcError('invalid_payload')
  }
  if (encoded.length > maxChars) throw new AiIpcError('payload_too_large')
}

function assertTrusted(event: InvokeEventLike, isTrustedSender: IsTrustedSender): void {
  if (!isTrustedSender(event.sender.id)) throw new AiIpcError('untrusted_sender')
}

function assertOneArg(args: unknown[]): unknown {
  if (args.length !== 1) throw new AiIpcError('invalid_payload')
  return args[0]
}

function validateUrlArray(value: unknown, maxItems: number, required = false): void {
  if (!Array.isArray(value) || (required && value.length === 0)) {
    throw new AiIpcError('invalid_payload')
  }
  if (value.length > maxItems) throw new AiIpcError('payload_too_large')
  for (const url of value) boundedString(url, 4_096, true)
}

function validateImagePayload(value: unknown): void {
  const op = plainObject(value)
  exactKeys(op, ['prompt', 'model', 'referenceImageUrls', 'aspectRatio', 'imageSize'])
  boundedString(op.prompt, 32_000, true)
  boundedString(op.model, 256)
  boundedString(op.aspectRatio, 256)
  boundedString(op.imageSize, 256)
  if (op.referenceImageUrls !== undefined) validateUrlArray(op.referenceImageUrls, 8)
  boundedJson(op, 100_000)
}

function validateMediaPayload(value: unknown): void {
  const op = plainObject(value)
  exactKeys(op, ['mediaUrls', 'requirements'])
  validateUrlArray(op.mediaUrls, 20, true)
  boundedString(op.requirements, 32_000, true)
  boundedJson(op, 150_000)
}

function validateCloudPayload(value: unknown): void {
  const op = plainObject(value)
  exactKeys(op, ['brief', 'title', 'styleSkill', 'deckContext', 'images', 'width', 'height'])
  boundedString(op.brief, 32_000, true)
  boundedString(op.title, 1_000)
  boundedString(op.styleSkill, 128_000)
  if (op.deckContext !== undefined) plainObject(op.deckContext)
  if (op.images !== undefined) {
    if (!Array.isArray(op.images)) throw new AiIpcError('invalid_payload')
    if (op.images.length > 20) throw new AiIpcError('payload_too_large')
    for (const value of op.images) {
      const image = plainObject(value)
      exactKeys(image, ['url', 'caption'])
      boundedString(image.url, 4_096, true)
      boundedString(image.caption, 4_096)
    }
  }
  for (const dimension of [op.width, op.height]) {
    if (dimension === undefined) continue
    if (
      typeof dimension !== 'number' ||
      !Number.isFinite(dimension) ||
      dimension <= 0 ||
      dimension > 10_000
    ) {
      throw new AiIpcError('invalid_payload')
    }
  }
  boundedJson(op, 2_000_000)
}

export function registerUnsupportedMediaIpc(
  ipcMain: IpcMainLike,
  isTrustedSender: IsTrustedSender,
): void {
  ipcMain.handle('ai:generate-image', (event, ...args) => {
    assertTrusted(event, isTrustedSender)
    validateImagePayload(assertOneArg(args))
    return UNSUPPORTED
  })
  ipcMain.handle('ai:analyze-media', (event, ...args) => {
    assertTrusted(event, isTrustedSender)
    validateMediaPayload(assertOneArg(args))
    return UNSUPPORTED
  })
}

export function registerUnsupportedCloudIpc(
  ipcMain: IpcMainLike,
  isTrustedSender: IsTrustedSender,
): void {
  ipcMain.handle('slides:cloud-gen-status', (event, ...args) => {
    assertTrusted(event, isTrustedSender)
    if (args.length !== 0) throw new AiIpcError('invalid_payload')
    return { enabled: false as const, errorCode: 'unsupported_feature' as const }
  })
  ipcMain.handle('slides:cloud-page-generate', (event, ...args) => {
    assertTrusted(event, isTrustedSender)
    validateCloudPayload(assertOneArg(args))
    return UNSUPPORTED
  })
}
