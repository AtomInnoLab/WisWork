import { isAbsolute, relative, resolve, sep } from 'node:path'
import { gunzipSync } from 'node:zlib'

export interface SyncTeXSourceLocation {
  readonly path: string
  readonly line: number
}

export interface SyncTeXPageLocation {
  readonly page: number
  readonly x: number
  readonly y: number
}

interface Point extends SyncTeXSourceLocation, SyncTeXPageLocation {}
interface RawPoint extends SyncTeXSourceLocation {
  readonly page: number
  readonly rawX: number
  readonly rawY: number
}

export interface SyncTeXIndex {
  forward(path: string, line: number): SyncTeXPageLocation | null
  inverse(page: number, x: number, y: number): SyncTeXSourceLocation | null
}

export interface SyncTeXParseOptions {
  readonly maxCompressedBytes?: number
  readonly maxDecompressedBytes?: number
  readonly maxInputs?: number
  readonly maxRecords?: number
}

const SP_PER_BP = 65781.76

function* lines(text: string): Generator<string> {
  let start = 0
  for (let index = 0; index <= text.length; index += 1) {
    if (index !== text.length && text.charCodeAt(index) !== 10) continue
    const end = index > start && text.charCodeAt(index - 1) === 13 ? index - 1 : index
    yield text.slice(start, end)
    start = index + 1
  }
}

function parsePostOffsetBp(value: string): number {
  const match = /^([+-]?\d+(?:\.\d+)?)(in|cm|mm|pt|bp|pc|sp|dd|cc|nd|nc)?$/.exec(value)
  if (!match) return Number.NaN
  const amount = Number(match[1])
  const factors: Record<string, number> = {
    in: 72,
    cm: 72 / 2.54,
    mm: 72 / 25.4,
    pt: 72 / 72.27,
    bp: 1,
    pc: (12 * 72) / 72.27,
    sp: 1 / SP_PER_BP,
    dd: (1238 / 1157) * (72 / 72.27),
    cc: 12 * (1238 / 1157) * (72 / 72.27),
    nd: (685 / 642) * (72 / 72.27),
    nc: 12 * (685 / 642) * (72 / 72.27),
    '': 1 / SP_PER_BP,
  }
  return amount * factors[match[2] ?? '']!
}
export function parseSyncTeX(
  compressed: Uint8Array,
  isolatedInputRoot: string,
  options: SyncTeXParseOptions = {},
): SyncTeXIndex {
  const maxCompressedBytes = options.maxCompressedBytes ?? 32 * 1024 * 1024
  const maxDecompressedBytes = options.maxDecompressedBytes ?? 64 * 1024 * 1024
  const maxInputs = options.maxInputs ?? 10_000
  const maxRecords = options.maxRecords ?? 250_000
  if (compressed.byteLength > maxCompressedBytes) {
    throw new Error('SyncTeX compressed size limit exceeded')
  }
  let text: string
  try {
    text = gunzipSync(compressed, { maxOutputLength: maxDecompressedBytes }).toString('utf8')
  } catch (error) {
    throw new Error('SyncTeX decompressed size limit exceeded', { cause: error })
  }
  if (Buffer.byteLength(text) > maxDecompressedBytes) {
    throw new Error('SyncTeX decompressed size limit exceeded')
  }

  const root = resolve(isolatedInputRoot)
  const inputs = new Map<number, string>()
  const rawPoints: RawPoint[] = []
  let inputCount = 0
  let recordCount = 0
  let page: number | null = null
  let unit = 1
  let preMagnification = 1000
  let preXOffsetSp = 0
  let preYOffsetSp = 0
  let postScriptum = false
  let postMagnification: number | null = null
  let postXOffsetBp: number | null = null
  let postYOffsetBp: number | null = null
  let lastRawY: number | null = null

  for (const line of lines(text)) {
    const input = /^Input:(\d+):(.*)$/.exec(line)
    if (input) {
      inputCount += 1
      if (inputCount > maxInputs) throw new Error('SyncTeX input limit exceeded')
      const inputValue = input[2]!
      const absolute = isAbsolute(inputValue) ? resolve(inputValue) : resolve(root, inputValue)
      const path = relative(root, absolute)
      if (!path.startsWith('..') && !isAbsolute(path) && absolute.startsWith(`${root}${sep}`)) {
        inputs.set(Number(input[1]), path.split(sep).join('/'))
      }
      continue
    }
    const unitLine = /^Unit:([+-]?\d+(?:\.\d+)?)$/.exec(line)
    if (unitLine) {
      unit = Number(unitLine[1])
      continue
    }
    const magnificationLine = /^Magnification:([+-]?\d+(?:\.\d+)?)$/.exec(line)
    if (magnificationLine) {
      if (postScriptum) postMagnification = Number(magnificationLine[1])
      else preMagnification = Number(magnificationLine[1])
      continue
    }
    const xOffsetLine = /^X Offset:(\S+)$/.exec(line)
    if (xOffsetLine) {
      if (postScriptum) {
        const value = parsePostOffsetBp(xOffsetLine[1]!)
        if (Number.isFinite(value)) postXOffsetBp = value
      } else {
        const value = Number(xOffsetLine[1])
        if (Number.isFinite(value)) preXOffsetSp = value
      }
      continue
    }
    const yOffsetLine = /^Y Offset:(\S+)$/.exec(line)
    if (yOffsetLine) {
      if (postScriptum) {
        const value = parsePostOffsetBp(yOffsetLine[1]!)
        if (Number.isFinite(value)) postYOffsetBp = value
      } else {
        const value = Number(yOffsetLine[1])
        if (Number.isFinite(value)) preYOffsetSp = value
      }
      continue
    }
    if (/^Post scriptum:/i.test(line)) {
      postScriptum = true
      page = null
      continue
    }
    const pageStart = /^\{(\d+)$/.exec(line)
    if (pageStart) {
      page = Number(pageStart[1])
      lastRawY = null
      continue
    }
    if (/^\}/.test(line)) {
      page = null
      lastRawY = null
      continue
    }
    const form = /^f-?\d+:-?\d+,(=|-?\d+)/.exec(line)
    if (form) {
      recordCount += 1
      if (recordCount > maxRecords) throw new Error('SyncTeX record limit exceeded')
      const formRawY: number | null = form[1] === '=' ? lastRawY : Number(form[1])
      if (formRawY !== null) lastRawY = formRawY
      continue
    }
    const record = /^(?:x|\$|v|k|h|g|\(|\[)(-?\d+),(-?\d+)(?:,-?\d+)?:(-?\d+),(=|-?\d+)/.exec(line)
    if (record && page !== null) {
      recordCount += 1
      if (recordCount > maxRecords) throw new Error('SyncTeX record limit exceeded')
      const rawY: number | null = record[4] === '=' ? lastRawY : Number(record[4])
      if (rawY === null) continue
      lastRawY = rawY
      const path = inputs.get(Number(record[1]))
      if (path) {
        rawPoints.push({
          path,
          line: Number(record[2]),
          page,
          rawX: Number(record[3]),
          rawY,
        })
      }
    }
  }

  const points: Point[] = rawPoints.map(({ rawX, rawY, ...point }) => ({
    ...point,
    x:
      (rawX * unit * (postMagnification ?? preMagnification / 1000)) / SP_PER_BP +
      (postXOffsetBp ?? preXOffsetSp / SP_PER_BP),
    y:
      (rawY * unit * (postMagnification ?? preMagnification / 1000)) / SP_PER_BP +
      (postYOffsetBp ?? preYOffsetSp / SP_PER_BP),
  }))

  return {
    forward(path, line) {
      const matches = points.filter((point) => point.path === path)
      if (matches.length === 0) return null
      const nearest = matches.reduce((best, point) =>
        Math.abs(point.line - line) < Math.abs(best.line - line) ? point : best,
      )
      return { page: nearest.page, x: nearest.x, y: nearest.y }
    },
    inverse(page, x, y) {
      const matches = points.filter((point) => point.page === page)
      if (matches.length === 0) return null
      const distance = (point: Point) => (point.x - x) ** 2 + (point.y - y) ** 2
      const nearest = matches.reduce((best, point) =>
        distance(point) < distance(best) ? point : best,
      )
      return { path: nearest.path, line: nearest.line }
    },
  }
}
