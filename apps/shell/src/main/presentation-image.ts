export const PRESENTATION_IMAGE_INPUT_LIMIT = 10 * 1024 * 1024
export const PRESENTATION_IMAGE_CACHE_LIMIT = 4 * 1024 * 1024
export interface NormalizedPresentationImage {
  bytes: Uint8Array
  width: number
  height: number
}
function invalid(): never {
  throw new Error('parse_failed')
}
/** Header admission only. The native decoder must subsequently validate the full image. */
export function inspectPresentationImage(input: Uint8Array): {
  mime: 'image/png' | 'image/jpeg'
  width: number
  height: number
} {
  if (!input.length || input.length > PRESENTATION_IMAGE_INPUT_LIMIT) invalid()
  const bytes = Buffer.from(input.buffer, input.byteOffset, input.byteLength)
  let width = 0,
    height = 0,
    mime: 'image/png' | 'image/jpeg'
  if (
    bytes.length >= 33 &&
    bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
  ) {
    mime = 'image/png'
    if (
      bytes.readUInt32BE(8) !== 13 ||
      bytes.toString('ascii', 12, 16) !== 'IHDR' ||
      bytes.toString('ascii', bytes.length - 8, bytes.length - 4) !== 'IEND'
    )
      invalid()
    width = bytes.readUInt32BE(16)
    height = bytes.readUInt32BE(20)
  } else if (
    bytes.length >= 4 &&
    bytes[0] === 255 &&
    bytes[1] === 216 &&
    bytes[bytes.length - 2] === 255 &&
    bytes[bytes.length - 1] === 217
  ) {
    mime = 'image/jpeg'
    let offset = 2
    while (offset + 4 <= bytes.length && bytes[offset] === 255) {
      const marker = bytes[offset + 1]!
      if (marker === 255) {
        offset++
        continue
      }
      if (marker === 0xda || marker === 0xd9) break
      if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        offset += 2
        continue
      }
      const length = bytes.readUInt16BE(offset + 2)
      if (length < 2 || offset + 2 + length > bytes.length) invalid()
      if ([0xc0, 0xc1, 0xc2].includes(marker)) {
        if (length < 8) invalid()
        height = bytes.readUInt16BE(offset + 5)
        width = bytes.readUInt16BE(offset + 7)
        break
      }
      offset += 2 + length
    }
  } else invalid()
  if (!width || !height || width > 8192 || height > 8192 || width * height > 16_000_000) invalid()
  return { mime, width, height }
}
/** Electron performs the full decode; no URLs, filenames, or external codec process are used. */
export async function normalizePresentationImage(
  input: Uint8Array,
): Promise<NormalizedPresentationImage> {
  const expected = inspectPresentationImage(input)
  const { nativeImage } = await import('electron')
  const image = nativeImage.createFromBuffer(Buffer.from(input))
  if (image.isEmpty()) invalid()
  const { width, height } = image.getSize()
  if (width !== expected.width || height !== expected.height) invalid()
  const bytes = image.toPNG()
  if (!bytes.length || bytes.length > PRESENTATION_IMAGE_CACHE_LIMIT) invalid()
  const actual = inspectPresentationImage(bytes)
  if (actual.mime !== 'image/png' || actual.width !== width || actual.height !== height) invalid()
  return { bytes, width, height }
}
