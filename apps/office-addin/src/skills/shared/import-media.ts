import type { InMemoryVfs } from './vfs.js'

export const MAX_IMPORT_BYTES = 2 * 1024 * 1024
export const MAX_CSV_ROWS = 500
export const MAX_CSV_COLUMNS = 100
export const MAX_CSV_CELLS = 10_000
export const MAX_IMAGE_PIXELS = 16_777_216
export const MAX_IMAGE_DIMENSION = 8_192

const decoder = new TextDecoder('utf-8', { fatal: true })

export function readBoundedCsv(vfs: InMemoryVfs, path: string): string[][] {
  const bytes = vfs.readBytes(path, { maxBytes: MAX_IMPORT_BYTES + 1 })
  if (bytes.byteLength > MAX_IMPORT_BYTES) throw new Error('import_limit')
  let source: string
  try {
    source = decoder.decode(bytes)
  } catch {
    throw new Error('invalid_csv')
  }
  const rows: string[][] = []
  let row: string[] = []
  let field = ''
  let quoted = false
  for (let index = 0; index <= source.length; index += 1) {
    const character = source[index]
    if (quoted) {
      if (character === '"') {
        if (source[index + 1] === '"') {
          field += '"'
          index += 1
        } else quoted = false
      } else if (character === undefined) throw new Error('invalid_csv')
      else field += character
      continue
    }
    if (character === '"' && field.length === 0) quoted = true
    else if (character === '"') throw new Error('invalid_csv')
    else if (character === ',') {
      row.push(field)
      if (row.length > MAX_CSV_COLUMNS) throw new Error('import_limit')
      field = ''
    } else if (character === '\n' || character === '\r' || character === undefined) {
      if (character === '\r' && source[index + 1] === '\n') index += 1
      row.push(field)
      if (row.length > MAX_CSV_COLUMNS || (rows.length + 1) * row.length > MAX_CSV_CELLS)
        throw new Error('import_limit')
      field = ''
      if (!(character === undefined && row.length === 1 && row[0] === '' && rows.length > 0))
        rows.push(row)
      row = []
      if (rows.length > MAX_CSV_ROWS) throw new Error('import_limit')
    } else field += character
    if (field.length > 32_768) throw new Error('import_limit')
  }
  const columns = rows.reduce((maximum, item) => Math.max(maximum, item.length), 0)
  if (
    !rows.length ||
    !columns ||
    columns > MAX_CSV_COLUMNS ||
    rows.length * columns > MAX_CSV_CELLS
  )
    throw new Error('import_limit')
  for (const item of rows) while (item.length < columns) item.push('')
  return rows
}

function safeCsvCell(value: unknown): string {
  let text = value == null ? '' : String(value)
  if (/^(?:[\t\r]|\s*[=+\-@])/.test(text)) text = `'${text}`
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

export function exportSafeCsv(values: readonly (readonly unknown[])[]): string {
  if (!values.length || values.length > MAX_CSV_ROWS) throw new Error('export_limit')
  const columns = values[0]?.length ?? 0
  if (!columns || columns > MAX_CSV_COLUMNS || values.length * columns > MAX_CSV_CELLS)
    throw new Error('export_limit')
  if (values.some((row) => row.length !== columns)) throw new Error('office_read_failed')
  const output = values.map((row) => row.map(safeCsvCell).join(',')).join('\r\n')
  if (new TextEncoder().encode(output).byteLength > MAX_IMPORT_BYTES)
    throw new Error('export_limit')
  return output
}

export interface BoundedImage {
  mime: 'image/png' | 'image/jpeg'
  base64: string
  bytes: number
  width: number
  height: number
  fingerprint: string
}

export function readBoundedImage(vfs: InMemoryVfs, path: string): BoundedImage {
  const bytes = vfs.readBytes(path, { maxBytes: MAX_IMPORT_BYTES + 1 })
  if (bytes.byteLength > MAX_IMPORT_BYTES) throw new Error('image_limit')
  const png =
    bytes.length >= 24 &&
    [137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value)
  let mime: BoundedImage['mime']
  let width: number
  let height: number
  if (png) {
    mime = 'image/png'
    ;({ width, height } = pngDimensions(bytes))
  } else if (bytes[0] === 0xff && bytes[1] === 0xd8) {
    mime = 'image/jpeg'
    ;({ width, height } = jpegDimensions(bytes))
  } else throw new Error('image_mime_unsupported')
  if (
    !width ||
    !height ||
    width > MAX_IMAGE_DIMENSION ||
    height > MAX_IMAGE_DIMENSION ||
    width * height > MAX_IMAGE_PIXELS
  )
    throw new Error('image_limit')
  return {
    mime,
    base64: toBase64(bytes),
    bytes: bytes.byteLength,
    width,
    height,
    fingerprint: hash(bytes),
  }
}

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff
  for (const byte of bytes) {
    crc ^= byte
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0)
  }
  return (crc ^ 0xffffffff) >>> 0
}

function pngDimensions(bytes: Uint8Array): { width: number; height: number } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let offset = 8
  let chunks = 0
  let sawData = false
  let sawPalette = false
  let colorType = -1
  let width = 0
  let height = 0
  while (offset + 12 <= bytes.length) {
    const length = view.getUint32(offset)
    const end = offset + 12 + length
    if (end > bytes.length) throw new Error('invalid_image')
    const type = String.fromCharCode(...bytes.subarray(offset + 4, offset + 8))
    if (!/^[A-Za-z]{4}$/.test(type)) throw new Error('invalid_image')
    if (
      view.getUint32(offset + 8 + length) !== crc32(bytes.subarray(offset + 4, offset + 8 + length))
    )
      throw new Error('invalid_image')
    if (chunks++ === 0) {
      if (type !== 'IHDR' || length !== 13) throw new Error('invalid_image')
      width = view.getUint32(offset + 8)
      height = view.getUint32(offset + 12)
      const bitDepth = bytes[offset + 16]
      colorType = bytes[offset + 17]
      const allowedDepths: Record<number, number[]> = {
        0: [1, 2, 4, 8, 16],
        2: [8, 16],
        3: [1, 2, 4, 8],
        4: [8, 16],
        6: [8, 16],
      }
      if (
        !allowedDepths[colorType]?.includes(bitDepth) ||
        bytes[offset + 18] !== 0 ||
        bytes[offset + 19] !== 0 ||
        bytes[offset + 20] > 1
      )
        throw new Error('invalid_image')
    } else if (type === 'IHDR') throw new Error('invalid_image')
    if (type === 'PLTE') {
      if (sawPalette || sawData || length < 3 || length > 768 || length % 3 !== 0)
        throw new Error('invalid_image')
      sawPalette = true
    } else if (type === 'IDAT') {
      if (!length || (colorType === 3 && !sawPalette)) throw new Error('invalid_image')
      sawData = true
    } else if (!['IHDR', 'IEND'].includes(type) && type[0] === type[0].toUpperCase()) {
      throw new Error('invalid_image')
    }
    if (type === 'IEND') {
      if (length !== 0 || !sawData || end !== bytes.length) throw new Error('invalid_image')
      return { width, height }
    }
    offset = end
  }
  throw new Error('invalid_image')
}

function jpegDimensions(bytes: Uint8Array): { width: number; height: number } {
  let offset = 2
  let width = 0
  let height = 0
  let sawSos = false
  while (offset < bytes.length) {
    if (bytes[offset++] !== 0xff) throw new Error('invalid_image')
    while (bytes[offset] === 0xff) offset += 1
    const marker = bytes[offset++]
    if (marker === 0xd9) {
      if (!width || !height || !sawSos || offset !== bytes.length) throw new Error('invalid_image')
      return { width, height }
    }
    if (
      marker === 0x00 ||
      marker === undefined ||
      marker === 0xd8 ||
      (marker >= 0xd0 && marker <= 0xd7)
    )
      throw new Error('invalid_image')
    if (offset + 2 > bytes.length) throw new Error('invalid_image')
    const length = (bytes[offset] << 8) | bytes[offset + 1]
    if (length < 2 || offset + length > bytes.length) throw new Error('invalid_image')
    if (
      [0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(
        marker,
      )
    ) {
      const components = bytes[offset + 7]
      if (width || height || !components || length !== 8 + 3 * components)
        throw new Error('invalid_image')
      height = (bytes[offset + 3] << 8) | bytes[offset + 4]
      width = (bytes[offset + 5] << 8) | bytes[offset + 6]
    }
    if (marker === 0xda) {
      const components = bytes[offset + 2]
      if (!width || !height || !components || length !== 6 + 2 * components)
        throw new Error('invalid_image')
    }
    offset += length
    if (marker === 0xda) {
      sawSos = true
      while (offset < bytes.length) {
        if (bytes[offset++] !== 0xff) continue
        while (bytes[offset] === 0xff) offset += 1
        const next = bytes[offset]
        if (next === 0x00 || (next >= 0xd0 && next <= 0xd7)) {
          offset += 1
          continue
        }
        offset -= 1
        break
      }
    }
  }
  throw new Error('invalid_image')
}

function toBase64(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 32 * 1024)
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 32 * 1024))
  return btoa(binary)
}

function hash(bytes: Uint8Array): string {
  let value = 0x811c9dc5
  for (const byte of bytes) value = Math.imul((value ^ byte) >>> 0, 0x01000193)
  return `${bytes.byteLength}:${(value >>> 0).toString(16).padStart(8, '0')}`
}
