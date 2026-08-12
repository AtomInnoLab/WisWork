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
  summary: { added: number; removed: number }
  truncated: boolean
}

export interface LineDiffLimits {
  contextLines?: number
  maxInputLines?: number
  maxOutputLines?: number
  maxOutputChars?: number
}

const DEFAULT_CONTEXT_LINES = 1
const DEFAULT_MAX_INPUT_LINES = 500
const DEFAULT_MAX_OUTPUT_LINES = 240
const DEFAULT_MAX_OUTPUT_CHARS = 30_000

function lines(text: string | null): string[] {
  if (text === null || text === '') return []
  return text.split('\n')
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
  const maxOutputLines = Math.max(1, limits.maxOutputLines ?? DEFAULT_MAX_OUTPUT_LINES)
  const maxOutputChars = Math.max(1, limits.maxOutputChars ?? DEFAULT_MAX_OUTPUT_CHARS)
  const allBefore = lines(beforeText)
  const allAfter = lines(afterText)
  let commonPrefix = 0
  while (
    commonPrefix < allBefore.length &&
    commonPrefix < allAfter.length &&
    allBefore[commonPrefix] === allAfter[commonPrefix]
  ) {
    commonPrefix += 1
  }
  if (commonPrefix === allBefore.length && commonPrefix === allAfter.length) {
    return { hunks: [], summary: { added: 0, removed: 0 }, truncated: false }
  }
  let commonSuffix = 0
  while (
    commonSuffix < allBefore.length - commonPrefix &&
    commonSuffix < allAfter.length - commonPrefix &&
    allBefore[allBefore.length - commonSuffix - 1] === allAfter[allAfter.length - commonSuffix - 1]
  ) {
    commonSuffix += 1
  }
  const beforeStart = Math.max(0, commonPrefix - contextLines)
  const afterStart = Math.max(0, commonPrefix - contextLines)
  const beforeEnd = Math.min(allBefore.length, allBefore.length - commonSuffix + contextLines)
  const afterEnd = Math.min(allAfter.length, allAfter.length - commonSuffix + contextLines)
  let truncated = beforeEnd - beforeStart > maxInputLines || afterEnd - afterStart > maxInputLines
  const ops = operations(
    allBefore.slice(beforeStart, beforeStart + maxInputLines),
    allAfter.slice(afterStart, afterStart + maxInputLines),
    beforeStart,
    afterStart,
  )
  const summary = {
    added: ops.filter((line) => line.kind === 'add').length,
    removed: ops.filter((line) => line.kind === 'remove').length,
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
      const text = line.text.slice(0, remainingChars)
      if (text.length < line.text.length) truncated = true
      hunkLines.push({ ...line, text })
      remainingChars -= text.length
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
