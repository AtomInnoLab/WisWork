import type { AgentImage, AgentToolContent } from './types'

export const OFFICE_SCREENSHOT_SOURCE_BYTES = 4 * 1024 * 1024
export const OFFICE_SCREENSHOT_PREVIEW_BYTES = 128 * 1024
// The output is itself JSON-escaped in the Relay frame. Reserve its control-field budget.
export const OFFICE_SCREENSHOT_WIRE_BYTES = 240 * 1024
const SCHEMA = 'wiswork.office-screenshot/1'
const encoder = new TextEncoder()
function unavailable(): never {
  throw new Error('office_screenshot_unavailable')
}
const exact = (value: unknown, keys: string[]): value is Record<string, unknown> =>
  Boolean(
    value &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    Object.keys(value).length === keys.length &&
    Object.keys(value).every((key) => keys.includes(key)),
  )

/** Canonical, bounded base64 only; never accepts a URL or a data URL. */
export function officeScreenshotBytes(image: AgentImage, maximum: number): Uint8Array {
  if (
    !exact(image, ['mime', 'base64']) ||
    !['image/png', 'image/jpeg'].includes(image.mime) ||
    typeof image.base64 !== 'string' ||
    !image.base64.length ||
    image.base64.length > Math.ceil(maximum / 3) * 4 ||
    image.base64.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]*={0,2}$/.test(image.base64)
  )
    unavailable()
  let binary: string
  try {
    binary = atob(image.base64)
  } catch {
    return unavailable()
  }
  if (binary.length > maximum || btoa(binary) !== image.base64) unavailable()
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0))
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let width = 0
  let height = 0
  if (
    image.mime === 'image/png'
      ? bytes.length < 45 ||
        ![137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value) ||
        view.getUint32(8) !== 13 ||
        view.getUint32(12) !== 0x49484452
      : bytes.length < 10 ||
        bytes[0] !== 255 ||
        bytes[1] !== 216 ||
        bytes.at(-2) !== 255 ||
        bytes.at(-1) !== 217
  )
    unavailable()
  if (image.mime === 'image/png') {
    width = view.getUint32(16)
    height = view.getUint32(20)
  } else {
    // Read JPEG dimensions before the native decoder allocates pixel storage.
    let offset = 2
    while (offset + 4 <= bytes.length) {
      if (bytes[offset++] !== 255) unavailable()
      while (bytes[offset] === 255) offset++
      const marker = bytes[offset++]!
      if (marker === 0xda || marker === 0xd9 || offset + 2 > bytes.length) break
      const size = view.getUint16(offset)
      if (size < 2 || offset + size > bytes.length) unavailable()
      if (
        [0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(
          marker,
        )
      ) {
        if (size < 8) unavailable()
        height = view.getUint16(offset + 3)
        width = view.getUint16(offset + 5)
        break
      }
      offset += size
    }
  }
  if (!width || !height || width > 8192 || height > 8192 || width * height > 16_000_000)
    unavailable()
  return bytes
}

function boundedWire(output: string): void {
  if (
    output.length > OFFICE_SCREENSHOT_WIRE_BYTES ||
    encoder.encode(JSON.stringify(output)).byteLength > OFFICE_SCREENSHOT_WIRE_BYTES
  )
    unavailable()
}

export function encodeOfficeScreenshotResult(
  output: string,
  content: AgentToolContent[] | undefined,
): string {
  if (!Array.isArray(content) || content.length !== 1) unavailable()
  const block = content[0]!
  if (!exact(block, ['type', 'image']) || block.type !== 'image') unavailable()
  officeScreenshotBytes(block.image, OFFICE_SCREENSHOT_PREVIEW_BYTES)
  if (output.length > OFFICE_SCREENSHOT_WIRE_BYTES) unavailable()
  let metadata: Record<string, unknown>
  try {
    metadata = JSON.parse(output)
  } catch {
    return unavailable()
  }
  if (!metadata || typeof metadata !== 'object' || Array.isArray(metadata)) unavailable()
  const wire = JSON.stringify({
    schema: SCHEMA,
    visualAvailableToModel: false,
    metadata: { ...metadata, visualAvailableToModel: false },
    image: block.image,
  })
  boundedWire(wire)
  return wire
}

export function decodeOfficeScreenshotResult(output: string): {
  output: string
  modelContent: AgentToolContent[]
} {
  boundedWire(output)
  let envelope: Record<string, unknown>
  try {
    envelope = JSON.parse(output)
  } catch {
    return unavailable()
  }
  if (
    !exact(envelope, ['schema', 'visualAvailableToModel', 'metadata', 'image']) ||
    envelope.schema !== SCHEMA ||
    envelope.visualAvailableToModel !== false ||
    !envelope.metadata ||
    typeof envelope.metadata !== 'object' ||
    Array.isArray(envelope.metadata) ||
    (envelope.metadata as Record<string, unknown>).visualAvailableToModel !== false
  )
    unavailable()
  const image = envelope.image as AgentImage
  const bytes = officeScreenshotBytes(image, OFFICE_SCREENSHOT_PREVIEW_BYTES)
  return {
    output: JSON.stringify({
      ...(envelope.metadata as Record<string, unknown>),
      mime: image.mime,
      bytes: bytes.byteLength,
      // This only unpacks the envelope. The PC must decode pixels before promoting it.
      visualAvailableToModel: false,
    }),
    modelContent: [{ type: 'image', image }],
  }
}
