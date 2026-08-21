import type { InMemoryVfs } from './vfs.js'
import { validateSkillPackageImage } from './skill-package.js'

export const MAX_IMPORT_BYTES = 2 * 1024 * 1024
export const MAX_CSV_ROWS = 500
export const MAX_CSV_COLUMNS = 100
export const MAX_CSV_CELLS = 10_000
export const MAX_IMAGE_PIXELS = 16_000_000
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

export function supportsBrowserMediaValidation(): boolean {
  return typeof createImageBitmap === 'function' && typeof DecompressionStream === 'function'
}

export async function readBoundedImage(vfs: InMemoryVfs, path: string): Promise<BoundedImage> {
  const bytes = vfs.readBytes(path, { maxBytes: MAX_IMPORT_BYTES + 1 })
  if (bytes.byteLength > MAX_IMPORT_BYTES) throw new Error('image_limit')
  const png =
    bytes.length >= 24 &&
    [137, 80, 78, 71, 13, 10, 26, 10].every((value, index) => bytes[index] === value)
  let mime: BoundedImage['mime']
  let extension: 'png' | 'jpeg'
  if (png) {
    mime = 'image/png'
    extension = 'png'
  } else if (bytes[0] === 0xff && bytes[1] === 0xd8) {
    mime = 'image/jpeg'
    extension = 'jpeg'
  } else throw new Error('image_mime_unsupported')
  let width: number
  let height: number
  try {
    ;({ width, height } = await validateSkillPackageImage(extension, bytes))
  } catch (error) {
    if (error instanceof Error && error.message === 'skill_package_limit')
      throw new Error('image_limit', { cause: new Error('media_limit') })
    throw new Error('invalid_image', {
      cause: new Error(error instanceof Error ? error.message : 'media_validation_failed'),
    })
  }
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
