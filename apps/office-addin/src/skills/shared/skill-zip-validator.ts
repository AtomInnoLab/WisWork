export interface SkillZipEntry {
  name: string
  compressed: number
  uncompressed: number
  crc: number
  directory: boolean
}

const LOCAL = 0x04034b50
const CENTRAL = 0x02014b50
const EOCD = 0x06054b50
const DESCRIPTOR = 0x08074b50
const decoder = new TextDecoder('utf-8', { fatal: true })

const invalid = (): never => {
  throw new Error('invalid_skill_package')
}
const limited = (): never => {
  throw new Error('skill_package_limit')
}

function u16(view: DataView, offset: number): number {
  if (offset < 0 || offset + 2 > view.byteLength) invalid()
  return view.getUint16(offset, true)
}
function u32(view: DataView, offset: number): number {
  if (offset < 0 || offset + 4 > view.byteLength) invalid()
  return view.getUint32(offset, true)
}
function extra(bytes: Uint8Array, start: number, length: number): void {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const end = start + length
  if (end > bytes.byteLength) invalid()
  const ids = new Set<number>()
  for (let offset = start; offset < end;) {
    if (offset + 4 > end) invalid()
    const id = u16(view, offset)
    const size = u16(view, offset + 2)
    offset += 4
    if (offset + size > end || ids.has(id) || id === 0x0001 || id === 0x9901) invalid()
    ids.add(id)
    offset += size
  }
}
function name(bytes: Uint8Array, flags: number): string {
  if (!(flags & 0x0800) && bytes.some((byte) => byte > 0x7f)) invalid()
  try {
    return decoder.decode(bytes)
  } catch {
    invalid()
  }
  return invalid()
}

export function validateSkillZip(
  bytes: Uint8Array,
  limits: {
    maxEntries: number
    maxCompressedBytes: number
    maxFileBytes: number
    maxUncompressedBytes: number
  },
): Map<string, SkillZipEntry> {
  if (bytes.byteLength < 22 || bytes.byteLength > limits.maxCompressedBytes) limited()
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let eocd = -1
  for (
    let offset = bytes.byteLength - 22;
    offset >= Math.max(0, bytes.byteLength - 65_557);
    offset -= 1
  ) {
    if (u32(view, offset) === EOCD) {
      eocd = offset
      break
    }
  }
  if (eocd < 0 || eocd + 22 + u16(view, eocd + 20) !== bytes.byteLength) invalid()
  const entries = u16(view, eocd + 10)
  const centralSize = u32(view, eocd + 12)
  const centralOffset = u32(view, eocd + 16)
  if (
    u16(view, eocd + 4) !== 0 ||
    u16(view, eocd + 6) !== 0 ||
    u16(view, eocd + 8) !== entries ||
    entries === 0xffff ||
    centralSize === 0xffffffff ||
    centralOffset === 0xffffffff ||
    centralOffset + centralSize !== eocd
  )
    invalid()
  if (entries > limits.maxEntries) limited()
  const result = new Map<string, SkillZipEntry>()
  const ranges: Array<[number, number]> = []
  let offset = centralOffset
  let total = 0
  for (let index = 0; index < entries; index += 1) {
    if (u32(view, offset) !== CENTRAL) invalid()
    const madeBy = bytes[offset + 5]
    const needed = u16(view, offset + 6)
    const flags = u16(view, offset + 8)
    const method = u16(view, offset + 10)
    const crc = u32(view, offset + 16)
    const compressed = u32(view, offset + 20)
    const uncompressed = u32(view, offset + 24)
    const nameLength = u16(view, offset + 28)
    const extraLength = u16(view, offset + 30)
    const commentLength = u16(view, offset + 32)
    const localOffset = u32(view, offset + 42)
    const end = offset + 46 + nameLength + extraLength + commentLength
    if (
      end > eocd ||
      u16(view, offset + 34) !== 0 ||
      needed > 20 ||
      compressed === 0xffffffff ||
      uncompressed === 0xffffffff ||
      localOffset === 0xffffffff ||
      (flags & 1) !== 0 ||
      (flags & ~0x080e) !== 0 ||
      ![0, 8].includes(method) ||
      (method === 0 && (flags & 0x0006) !== 0)
    )
      invalid()
    const entryName = name(bytes.subarray(offset + 46, offset + 46 + nameLength), flags)
    if (result.has(entryName)) invalid()
    const directory = entryName.endsWith('/')
    const external = u32(view, offset + 38)
    if (madeBy === 3) {
      const type = (external >>> 16) & 0xf000
      if (type !== 0 && type !== (directory ? 0x4000 : 0x8000)) invalid()
      if (!directory && ((external >>> 16) & 0o111) !== 0) invalid()
    } else if ((external & 0x08) !== 0 || Boolean(external & 0x10) !== directory) invalid()
    extra(bytes, offset + 46 + nameLength, extraLength)
    if (directory && (compressed || uncompressed || method)) invalid()
    if (uncompressed > limits.maxFileBytes) limited()
    total += uncompressed
    if (total > limits.maxUncompressedBytes) limited()

    if (u32(view, localOffset) !== LOCAL) invalid()
    const localFlags = u16(view, localOffset + 6)
    const localMethod = u16(view, localOffset + 8)
    const localNameLength = u16(view, localOffset + 26)
    const localExtraLength = u16(view, localOffset + 28)
    const dataOffset = localOffset + 30 + localNameLength + localExtraLength
    const localName = name(
      bytes.subarray(localOffset + 30, localOffset + 30 + localNameLength),
      localFlags,
    )
    extra(bytes, localOffset + 30 + localNameLength, localExtraLength)
    if (
      localFlags !== flags ||
      localMethod !== method ||
      localName !== entryName ||
      dataOffset + compressed > centralOffset
    )
      invalid()
    let localEnd = dataOffset + compressed
    if ((flags & 8) === 0) {
      if (
        u32(view, localOffset + 14) !== crc ||
        u32(view, localOffset + 18) !== compressed ||
        u32(view, localOffset + 22) !== uncompressed
      )
        invalid()
    } else {
      if (
        ![0, crc].includes(u32(view, localOffset + 14)) ||
        ![0, compressed].includes(u32(view, localOffset + 18)) ||
        ![0, uncompressed].includes(u32(view, localOffset + 22))
      )
        invalid()
      if (u32(view, localEnd) === DESCRIPTOR) localEnd += 4
      if (
        u32(view, localEnd) !== crc ||
        u32(view, localEnd + 4) !== compressed ||
        u32(view, localEnd + 8) !== uncompressed
      )
        invalid()
      localEnd += 12
    }
    if (localEnd > centralOffset) invalid()
    ranges.push([localOffset, localEnd])
    result.set(entryName, { name: entryName, compressed, uncompressed, crc, directory })
    offset = end
  }
  if (offset !== eocd) invalid()
  ranges.sort((a, b) => a[0] - b[0])
  if (ranges.length && ranges[0][0] !== 0) invalid()
  for (let index = 1; index < ranges.length; index++)
    if (ranges[index][0] !== ranges[index - 1][1]) invalid()
  if (ranges.length && ranges.at(-1)?.[1] !== centralOffset) invalid()
  return result
}
