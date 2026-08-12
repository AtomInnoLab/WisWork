export type DiffLineKind = 'context' | 'add' | 'remove'

export interface DiffLine {
  kind: DiffLineKind
  text: string
  beforeLine: number | null
  afterLine: number | null
}

export interface DiffHunk {
  beforeStart: number
  afterStart: number
  lines: DiffLine[]
}

export interface LineDiff {
  hunks: DiffHunk[]
  summary: { added: number; removed: number; atLeast: boolean }
  truncated: boolean
}

export interface LineDiffLimits {
  contextLines?: number
  maxInputLines?: number
  maxInputChars?: number
  maxOutputLines?: number
  maxOutputChars?: number
}

const DEFAULT_CONTEXT_LINES = 1
const DEFAULT_MAX_INPUT_LINES = 500
const DEFAULT_MAX_INPUT_CHARS = 64_000
const DEFAULT_MAX_OUTPUT_LINES = 240
const DEFAULT_MAX_OUTPUT_CHARS = 30_000

function collectLines(
  text: string | null,
  maxLines: number,
  maxChars: number,
): { lines: string[]; truncated: boolean } {
  if (text === null || text === '') return { lines: [], truncated: false }
  const result: string[] = []
  let current = ''
  let index = 0
  let consumedChars = 0
  while (index < text.length && result.length < maxLines && consumedChars < maxChars) {
    const codePoint = text.codePointAt(index)!
    const character = String.fromCodePoint(codePoint)
    index += character.length
    consumedChars += 1
    if (character === '\n') {
      result.push(current)
      current = ''
    } else current += character
  }
  let truncated = index < text.length
  if (current || (!truncated && text.endsWith('\n'))) {
    if (result.length < maxLines) result.push(current)
    else truncated = true
  }
  return { lines: result, truncated }
}

function takeCodePoints(text: string, count: number): { text: string; count: number } {
  let value = ''
  let consumed = 0
  for (const character of text) {
    if (consumed >= count) break
    value += character
    consumed += 1
  }
  return { text: value, count: consumed }
}

function operations(
  before: readonly string[],
  after: readonly string[],
  beforeOffset: number,
  afterOffset: number,
): DiffLine[] {
  const widths = after.length + 1
  const table = new Uint16Array((before.length + 1) * widths)
  for (let left = before.length - 1; left >= 0; left -= 1) {
    for (let right = after.length - 1; right >= 0; right -= 1) {
      const index = left * widths + right
      table[index] =
        before[left] === after[right]
          ? table[(left + 1) * widths + right + 1] + 1
          : Math.max(table[(left + 1) * widths + right], table[left * widths + right + 1])
    }
  }

  const result: DiffLine[] = []
  let left = 0
  let right = 0
  while (left < before.length || right < after.length) {
    if (left < before.length && right < after.length && before[left] === after[right]) {
      result.push({
        kind: 'context',
        text: before[left]!,
        beforeLine: beforeOffset + left + 1,
        afterLine: afterOffset + right + 1,
      })
      left += 1
      right += 1
    } else if (
      right >= after.length ||
      (left < before.length &&
        table[(left + 1) * widths + right] >= table[left * widths + right + 1])
    ) {
      result.push({
        kind: 'remove',
        text: before[left]!,
        beforeLine: beforeOffset + left + 1,
        afterLine: null,
      })
      left += 1
    } else {
      result.push({
        kind: 'add',
        text: after[right]!,
        beforeLine: null,
        afterLine: afterOffset + right + 1,
      })
      right += 1
    }
  }
  return result
}

export function buildLineDiff(
  beforeText: string | null,
  afterText: string,
  limits: LineDiffLimits = {},
): LineDiff {
  const contextLines = Math.max(0, limits.contextLines ?? DEFAULT_CONTEXT_LINES)
  const maxInputLines = Math.max(1, limits.maxInputLines ?? DEFAULT_MAX_INPUT_LINES)
  const maxInputChars = Math.max(1, limits.maxInputChars ?? DEFAULT_MAX_INPUT_CHARS)
  const maxOutputLines = Math.max(1, limits.maxOutputLines ?? DEFAULT_MAX_OUTPUT_LINES)
  const maxOutputChars = Math.max(1, limits.maxOutputChars ?? DEFAULT_MAX_OUTPUT_CHARS)
  const before = collectLines(beforeText, maxInputLines, maxInputChars)
  const after = collectLines(afterText, maxInputLines, maxInputChars)
  let truncated = before.truncated || after.truncated
  const ops = operations(before.lines, after.lines, 0, 0)
  const summary = {
    added: ops.filter((line) => line.kind === 'add').length,
    removed: ops.filter((line) => line.kind === 'remove').length,
    atLeast: truncated,
  }
  const changes = ops.flatMap((line, index) => (line.kind === 'context' ? [] : [index]))
  if (!changes.length) return { hunks: [], summary, truncated }

  const ranges: Array<{ start: number; end: number }> = []
  for (const index of changes) {
    const start = Math.max(0, index - contextLines)
    const end = Math.min(ops.length - 1, index + contextLines)
    const previous = ranges.at(-1)
    if (previous && start <= previous.end) previous.end = Math.max(previous.end, end)
    else ranges.push({ start, end })
  }

  let remainingLines = maxOutputLines
  let remainingChars = maxOutputChars
  const hunks: DiffHunk[] = []
  for (const range of ranges) {
    if (remainingLines === 0 || remainingChars === 0) {
      truncated = true
      break
    }
    const hunkLines: DiffLine[] = []
    for (const line of ops.slice(range.start, range.end + 1)) {
      if (remainingLines === 0 || remainingChars === 0) {
        truncated = true
        break
      }
      const bounded = takeCodePoints(line.text, remainingChars)
      if (bounded.text.length < line.text.length) truncated = true
      hunkLines.push({ ...line, text: bounded.text })
      remainingChars -= bounded.count
      remainingLines -= 1
    }
    if (hunkLines.length) {
      hunks.push({
        beforeStart: hunkLines.find((line) => line.beforeLine !== null)?.beforeLine ?? 0,
        afterStart: hunkLines.find((line) => line.afterLine !== null)?.afterLine ?? 0,
        lines: hunkLines,
      })
    }
  }
  return { hunks, summary, truncated }
}
