type HandoffImage = { mime: 'image/png' | 'image/jpeg'; bytes: Uint8Array }
const MAX_HANDOFF_BYTES = 180 * 1024
// Source download budget is independent of the much smaller Relay payload budget.
export const MAX_OFFICE_IMAGE_SOURCE_BYTES = 10 * 1024 * 1024

interface DecodedImage {
  isEmpty(): boolean
  getSize(): { width: number; height: number }
  resize(options: { width: number; height: number; quality: 'best' }): DecodedImage
  toPNG(): Uint8Array
  toJPEG(quality: number): Uint8Array
}

export function createOfficeImageHandoff(
  nativeImage: { createFromBuffer(bytes: Buffer): DecodedImage },
  options: { reencode?: boolean } = {},
): (image: HandoffImage) => Promise<HandoffImage> {
  return async (image) => {
    if (image.bytes.byteLength > MAX_OFFICE_IMAGE_SOURCE_BYTES) throw new Error('image_limit')
    if (!image.bytes.byteLength) throw new Error('invalid_image')
    if (image.mime !== 'image/png' && image.mime !== 'image/jpeg')
      throw new Error('image_mime_unsupported')
    let decoded: DecodedImage
    try {
      decoded = nativeImage.createFromBuffer(Buffer.from(image.bytes))
    } catch {
      throw new Error('invalid_image')
    }
    if (decoded.isEmpty()) throw new Error('invalid_image')
    const { width, height } = decoded.getSize()
    if (
      !Number.isSafeInteger(width) ||
      !Number.isSafeInteger(height) ||
      width < 1 ||
      height < 1 ||
      width > 8192 ||
      height > 8192 ||
      width * height > 16_000_000
    )
      throw new Error('image_limit')
    if (!options.reencode && image.bytes.byteLength <= MAX_HANDOFF_BYTES) return image

    // This private 256 KiB tool envelope is separate from the 2 MiB retrieval output.
    // Keep room for base64 expansion and provenance; never truncate encoded bytes.
    // A large slide photo must not silently become a 256px thumbnail to fit Relay.
    for (const edge of [1536, 1024, 960]) {
      const scale = Math.min(1, edge / Math.max(width, height))
      try {
        const resized =
          scale === 1
            ? decoded
            : decoded.resize({
                width: Math.max(1, Math.round(width * scale)),
                height: Math.max(1, Math.round(height * scale)),
                quality: 'best',
              })
        const png = resized.toPNG()
        if (png.byteLength && png.byteLength <= MAX_HANDOFF_BYTES)
          return { mime: 'image/png', bytes: png }
        const jpeg = resized.toJPEG(80)
        if (jpeg.byteLength && jpeg.byteLength <= MAX_HANDOFF_BYTES)
          return { mime: 'image/jpeg', bytes: jpeg }
      } catch {
        throw new Error('invalid_image')
      }
    }
    throw new Error('image_limit')
  }
}
