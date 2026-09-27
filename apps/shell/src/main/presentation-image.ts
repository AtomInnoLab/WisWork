export const PRESENTATION_IMAGE_INPUT_LIMIT = 10 * 1024 * 1024
export const PRESENTATION_IMAGE_CACHE_LIMIT = 4 * 1024 * 1024
export interface NormalizedPresentationImage {
  bytes: Uint8Array
  width: number
  height: number
}
type ImageMime = 'image/png' | 'image/jpeg' | 'image/gif' | 'image/webp'
let decoderView: import('electron').WebContentsView | undefined
let decodeQueue: Promise<void> = Promise.resolve()
function invalid(): never {
  throw new Error('parse_failed')
}
/** Header admission only. The native decoder must subsequently validate the full image. */
export function inspectPresentationImage(input: Uint8Array): {
  mime: ImageMime
  width: number
  height: number
} {
  if (!input.length || input.length > PRESENTATION_IMAGE_INPUT_LIMIT) invalid()
  const bytes = Buffer.from(input.buffer, input.byteOffset, input.byteLength)
  let width = 0,
    height = 0,
    mime: ImageMime
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
  } else if (bytes.length >= 14 && ['GIF87a', 'GIF89a'].includes(bytes.toString('ascii', 0, 6))) {
    mime = 'image/gif'
    width = bytes.readUInt16LE(6)
    height = bytes.readUInt16LE(8)
    let offset = 13
    if (bytes[10]! & 0x80) offset += 3 * (1 << ((bytes[10]! & 7) + 1))
    let frames = 0
    let trailer = false
    while (offset < bytes.length) {
      const tag = bytes[offset++]!
      if (tag === 0x3b) {
        if (offset !== bytes.length || frames !== 1) invalid()
        trailer = true
        break
      }
      if (tag === 0x2c) {
        if (++frames > 1 || offset + 9 > bytes.length) invalid()
        const flags = bytes[offset + 8]!
        offset += 9
        if (flags & 0x80) offset += 3 * (1 << ((flags & 7) + 1))
        if (offset >= bytes.length) invalid()
        offset++
      } else if (tag === 0x21) {
        if (offset >= bytes.length) invalid()
        const label = bytes[offset++]!
        if (label === 0xff && bytes.toString('ascii', offset + 1, offset + 12) === 'NETSCAPE2.0')
          invalid()
      } else invalid()
      while (true) {
        if (offset >= bytes.length) invalid()
        const size = bytes[offset++]!
        if (!size) break
        offset += size
        if (offset > bytes.length) invalid()
      }
    }
    if (!trailer) invalid()
  } else if (
    bytes.length >= 30 &&
    bytes.toString('ascii', 0, 4) === 'RIFF' &&
    bytes.toString('ascii', 8, 12) === 'WEBP'
  ) {
    mime = 'image/webp'
    if (bytes.readUInt32LE(4) + 8 !== bytes.length) invalid()
    let offset = 12
    let imageChunks = 0
    let extended = false
    while (offset + 8 <= bytes.length) {
      const kind = bytes.toString('ascii', offset, offset + 4)
      const size = bytes.readUInt32LE(offset + 4)
      const end = offset + 8 + size + (size & 1)
      if (end > bytes.length) invalid()
      if (kind === 'VP8X') {
        if (size !== 10 || offset !== 12 || extended || bytes[offset + 8]! & 0xc3) invalid()
        extended = true
        width = 1 + bytes.readUIntLE(offset + 12, 3)
        height = 1 + bytes.readUIntLE(offset + 15, 3)
      } else if (kind === 'VP8 ' || kind === 'VP8L') {
        if (++imageChunks > 1 || (!extended && offset !== 12)) invalid()
        if (kind === 'VP8 ') {
          if (
            size < 10 ||
            bytes[offset + 11] !== 0x9d ||
            bytes[offset + 12] !== 0x01 ||
            bytes[offset + 13] !== 0x2a
          )
            invalid()
          const w = bytes.readUInt16LE(offset + 14) & 0x3fff
          const h = bytes.readUInt16LE(offset + 16) & 0x3fff
          if (width && (width !== w || height !== h)) invalid()
          width = w
          height = h
        } else {
          if (size < 5 || bytes[offset + 8] !== 0x2f) invalid()
          const packed = bytes.readUInt32LE(offset + 9)
          const w = 1 + (packed & 0x3fff)
          const h = 1 + ((packed >>> 14) & 0x3fff)
          if (width && (width !== w || height !== h)) invalid()
          width = w
          height = h
        }
      } else if (kind === 'ANIM' || kind === 'ANMF' || !extended) invalid()
      offset = end
    }
    if (offset !== bytes.length || imageChunks !== 1) invalid()
  } else invalid()
  if (!width || !height || width > 8192 || height > 8192 || width * height > 16_000_000) invalid()
  return { mime, width, height }
}
async function decodeBrowserImage(
  input: Uint8Array,
  expected: { mime: ImageMime; width: number; height: number },
): Promise<Buffer> {
  const run = async () => {
    const { app, WebContentsView, session } = await import('electron')
    if (!app.isReady()) invalid()
    if (!decoderView || decoderView.webContents.isDestroyed()) {
      const partition = 'ppt-image-decoder'
      session
        .fromPartition(partition, { cache: false })
        .webRequest.onBeforeRequest((details, callback) =>
          callback({ cancel: !details.url.startsWith('data:') }),
        )
      const view = new WebContentsView({
        webPreferences: {
          partition,
          sandbox: true,
          nodeIntegration: false,
          contextIsolation: true,
          webSecurity: true,
        },
      })
      const contents = view.webContents
      contents.setWindowOpenHandler(() => ({ action: 'deny' }))
      try {
        await contents.loadURL(
          'data:text/html,<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; img-src data:; script-src \'unsafe-inline\'">',
        )
        if (contents.isDestroyed()) invalid()
        contents.on('will-navigate', (event) => event.preventDefault())
        decoderView = view
      } catch {
        if (!contents.isDestroyed()) contents.close()
        invalid()
      }
    }
    const contents = decoderView.webContents
    const url = `data:${expected.mime};base64,${Buffer.from(input).toString('base64')}`
    const result = await contents.executeJavaScript(`new Promise((resolve, reject) => {
      const image = new Image();
      image.onerror = () => reject(Error('parse_failed'));
      image.onload = () => {
        if (image.naturalWidth !== ${expected.width} || image.naturalHeight !== ${expected.height}) return reject(Error('parse_failed'));
        const canvas = document.createElement('canvas');
        canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
        canvas.getContext('2d').drawImage(image, 0, 0);
        resolve(canvas.toDataURL('image/png'));
      };
      image.src = ${JSON.stringify(url)};
    })`)
    if (typeof result !== 'string' || !result.startsWith('data:image/png;base64,')) invalid()
    return Buffer.from(result.slice('data:image/png;base64,'.length), 'base64')
  }
  const job = decodeQueue.then(async () => {
    let timer: ReturnType<typeof setTimeout> | undefined
    try {
      return await Promise.race([
        run(),
        new Promise<never>((_, reject) => {
          timer = setTimeout(() => reject(new Error('parse_failed')), 5000)
        }),
      ])
    } catch {
      if (decoderView && !decoderView.webContents.isDestroyed()) decoderView.webContents.close()
      decoderView = undefined
      invalid()
    } finally {
      if (timer) clearTimeout(timer)
    }
  })
  decodeQueue = job.then(
    () => {},
    () => {},
  )
  return job
}
/** Electron performs the full decode; no external URL, filename, or codec process is used. */
export async function normalizePresentationImage(
  input: Uint8Array,
): Promise<NormalizedPresentationImage> {
  const expected = inspectPresentationImage(input)
  const { nativeImage } = await import('electron')
  let bytes: Buffer
  if (expected.mime === 'image/gif' || expected.mime === 'image/webp') {
    bytes = await decodeBrowserImage(input, expected)
  } else {
    const image = nativeImage.createFromBuffer(Buffer.from(input))
    if (image.isEmpty()) invalid()
    const { width, height } = image.getSize()
    if (width !== expected.width || height !== expected.height) invalid()
    bytes = image.toPNG()
  }
  if (!bytes.length || bytes.length > PRESENTATION_IMAGE_CACHE_LIMIT) invalid()
  const actual = inspectPresentationImage(bytes)
  if (
    actual.mime !== 'image/png' ||
    actual.width !== expected.width ||
    actual.height !== expected.height
  )
    invalid()
  return { bytes, width: expected.width, height: expected.height }
}
