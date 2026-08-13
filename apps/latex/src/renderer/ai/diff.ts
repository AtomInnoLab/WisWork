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
  notice?: 'change-location-beyond-preview-budget'
}

export interface LineDiffLimits {
  contextLines?: number
  maxScanChars?: number
  maxInputLines?: number
  maxInputChars?: number
  maxOutputLines?: number
  maxOutputChars?: number
}

const DEFAULT_CONTEXT_LINES = 1
const DEFAULT_MAX_SCAN_CHARS = 1024 * 1024
const DEFAULT_MAX_INPUT_LINES = 500
const DEFAULT_MAX_INPUT_CHARS = 64_000
const DEFAULT_MAX_OUTPUT_LINES = 240
const DEFAULT_MAX_OUTPUT_CHARS = 30_000

interface LineSpan {
  start: number
  end: number
  line: number
}

interface LineCursor {
  text: string
  index: number
  line: number
  trailingEmpty: boolean
}

function cursor(text: string | null): LineCursor {
  return { text: text ?? '', index: 0, line: 1, trailingEmpty: false }
}

function nextLine(cursor: LineCursor): LineSpan | null {
  if (cursor.trailingEmpty) {
    cursor.trailingEmpty = false
    return { start: cursor.text.length, end: cursor.text.length, line: cursor.line++ }
  }
  if (cursor.index >= cursor.text.length) return null
  const newline = cursor.text.indexOf('\n', cursor.index)
  const start = cursor.index
  const end = newline === -1 ? cursor.text.length : newline
  cursor.index = newline === -1 ? cursor.text.length : newline + 1
  if (newline === cursor.text.length - 1) cursor.trailingEmpty = true
  return { start, end, line: cursor.line++ }
}

function hasMore(cursor: LineCursor): boolean {
  return cursor.trailingEmpty || cursor.index < cursor.text.length
}

function equalLine(
  beforeText: string,
  before: LineSpan,
  afterText: string,
  after: LineSpan,
): boolean {
  const length = before.end - before.start
  if (length !== after.end - after.start) return false
  for (let index = 0; index < length; index += 1) {
    if (beforeText.charCodeAt(before.start + index) !== afterText.charCodeAt(after.start + index))
      return false
  }
  return true
}

function boundedSpan(text: string, span: LineSpan, maxChars: number) {
  let value = ''
  let consumed = 0
  let index = span.start
  while (index < span.end && consumed < maxChars) {
    const codePoint = text.codePointAt(index)!
    const character = String.fromCodePoint(codePoint)
    value += character
    index += character.length
    consumed += 1
  }
  return { text: value, count: consumed }
}

function collectWindow(
  cursor: LineCursor,
  prefix: readonly LineSpan[],
  first: LineSpan | null,
  maxLines: number,
  maxChars: number,
): { lines: string[]; startLine: number; truncated: boolean } {
  const retainedPrefix = prefix.slice(-(first ? Math.max(0, maxLines - 1) : maxLines))
  const spans = [...retainedPrefix]
  if (first) spans.push(first)
  let next = first ? nextLine(cursor) : null
  while (next && spans.length < maxLines) {
    spans.push(next)
    next = nextLine(cursor)
  }

  const lines: string[] = []
  let remainingChars = maxChars
  let truncated = Boolean(next || hasMore(cursor))
  for (let index = 0; index < spans.length; index += 1) {
    const span = spans[index]!
    const context = index < retainedPrefix.length
    const allowance = context ? Math.min(2_000, Math.max(0, remainingChars - 1)) : remainingChars
    if (allowance <= 0) {
      truncated = true
      break
    }
    const bounded = boundedSpan(cursor.text, span, allowance)
    lines.push(bounded.text)
    remainingChars -= bounded.count
    if (bounded.text.length < span.end - span.start) truncated = true
  }
  if (lines.length < spans.length) truncated = true
  return { lines, startLine: spans[0]?.line ?? cursor.line, truncated }
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
  const maxScanChars = Math.max(1, limits.maxScanChars ?? DEFAULT_MAX_SCAN_CHARS)
  const maxInputLines = Math.max(1, limits.maxInputLines ?? DEFAULT_MAX_INPUT_LINES)
  const maxInputChars = Math.max(1, limits.maxInputChars ?? DEFAULT_MAX_INPUT_CHARS)
  const maxOutputLines = Math.max(1, limits.maxOutputLines ?? DEFAULT_MAX_OUTPUT_LINES)
  const maxOutputChars = Math.max(1, limits.maxOutputChars ?? DEFAULT_MAX_OUTPUT_CHARS)
  const normalizedBefore = beforeText ?? ''
  if (normalizedBefore === afterText) {
    return {
      hunks: [],
      summary: { added: 0, removed: 0, atLeast: false },
      truncated: false,
    }
  }
  const beforeCursor = cursor(beforeText)
  const afterCursor = cursor(afterText)
  const beforePrefix: LineSpan[] = []
  const afterPrefix: LineSpan[] = []
  let beforeFirst = nextLine(beforeCursor)
  let afterFirst = nextLine(afterCursor)
  let scannedChars = 0
  while (beforeFirst && afterFirst) {
    const nextCost = beforeFirst.end - beforeFirst.start + afterFirst.end - afterFirst.start + 2
    if (scannedChars + nextCost > maxScanChars) {
      return {
        hunks: [],
        summary: { added: 0, removed: 0, atLeast: true },
        truncated: true,
        notice: 'change-location-beyond-preview-budget',
      }
    }
    if (!equalLine(normalizedBefore, beforeFirst, afterText, afterFirst)) break
    scannedChars += nextCost
    beforePrefix.push(beforeFirst)
    afterPrefix.push(afterFirst)
    if (beforePrefix.length > contextLines) beforePrefix.shift()
    if (afterPrefix.length > contextLines) afterPrefix.shift()
    beforeFirst = nextLine(beforeCursor)
    afterFirst = nextLine(afterCursor)
  }

  const before = collectWindow(
    beforeCursor,
    beforePrefix,
    beforeFirst,
    maxInputLines,
    maxInputChars,
  )
  const after = collectWindow(afterCursor, afterPrefix, afterFirst, maxInputLines, maxInputChars)
  let truncated = before.truncated || after.truncated
  const ops = operations(before.lines, after.lines, before.startLine - 1, after.startLine - 1)
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
