import { formatAddress } from '../domain/cell-address'
import type { IRange } from '@univerjs/core'
import { t } from './i18n/locale'
import { netAxisDelta } from './view-transform'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import {
  ensureLazyRangeLoaded,
  isLazyWorkbookTargetCurrent,
  MappedRangeByteBudgetError,
  MappedRangeRequestBudgetError,
  readSheetRangeMapped,
} from './univer-sync'

export const ERROR_SCAN_DEFAULT_BUDGETS = {
  maxCells: 400_000,
  maxFindings: 10_000,
  maxBatches: 64,
  maxBytes: 16 * 1024 * 1024,
  maxMs: 10_000,
  batchCells: 18_000,
} as const

export type SpreadsheetErrorCode =
  | '#NULL!'
  | '#DIV/0!'
  | '#VALUE!'
  | '#REF!'
  | '#NAME?'
  | '#NUM!'
  | '#N/A'
  | '#GETTING_DATA'
  | '#SPILL!'
  | '#CALC!'
  | '#FIELD!'
  | '#BLOCKED!'
  | '#UNKNOWN!'
  | '#CONNECT!'
  | '#BUSY!'
  | '#PYTHON!'

const ERROR_CODES = new Set<SpreadsheetErrorCode>([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
  '#FIELD!',
  '#BLOCKED!',
  '#UNKNOWN!',
  '#CONNECT!',
  '#BUSY!',
  '#PYTHON!',
])

export interface WorkbookErrorFinding {
  readonly sheetId: string
  readonly address: string
  readonly errorCode: SpreadsheetErrorCode
}

export interface ErrorScanResult {
  readonly status: 'complete' | 'truncated' | 'unavailable' | 'cancelled'
  readonly findings: readonly WorkbookErrorFinding[]
  readonly scanned: { readonly cells: number; readonly batches: number; readonly bytes: number }
}

export interface ErrorScanIdentity {
  /** Checks the captured document/workbook/session and, for a batch, exact worksheet object. */
  readonly isCurrent: (sheetId?: string, worksheet?: unknown) => boolean
  readonly worksheet: (sheetId: string) => unknown
  readonly signal?: AbortSignal
  readonly onProgress?: (progress: ErrorScanResult['scanned']) => void
}

export interface ErrorScanBudgets {
  readonly maxCells: number
  readonly maxFindings: number
  readonly maxBatches: number
  readonly maxBytes: number
  readonly maxMs: number
  readonly batchCells: number
  readonly now?: () => number
}

function errorCode(value: unknown): SpreadsheetErrorCode | null {
  return typeof value === 'string' && ERROR_CODES.has(value as SpreadsheetErrorCode)
    ? (value as SpreadsheetErrorCode)
    : null
}

function journalStamp(state: LazyWorkbookState): string {
  const cells = [...state.editJournal.cells.entries()]
    .map(([sheetId, entries]) => `${sheetId}:${entries.size}:${[...entries.keys()].join(',')}`)
    .join('|')
  const ops = [...state.editJournal.structuralOps.entries()]
    .map(([sheetId, entries]) => `${sheetId}:${entries.length}`)
    .join('|')
  return `${cells}/${ops}/${[...state.editJournal.filterDirty].sort().join(',')}`
}

function intersectRange(left: IRange, right: IRange): IRange | null {
  const result = {
    startRow: Math.max(left.startRow, right.startRow),
    endRow: Math.min(left.endRow, right.endRow),
    startColumn: Math.max(left.startColumn, right.startColumn),
    endColumn: Math.min(left.endColumn, right.endColumn),
  }
  return result.startRow <= result.endRow && result.startColumn <= result.endColumn ? result : null
}

function subtractRange(range: IRange, blocker: IRange): IRange[] {
  const overlap = intersectRange(range, blocker)
  if (!overlap) return [range]
  const result: IRange[] = []
  if (range.startRow < overlap.startRow) result.push({ ...range, endRow: overlap.startRow - 1 })
  if (overlap.endRow < range.endRow) result.push({ ...range, startRow: overlap.endRow + 1 })
  if (range.startColumn < overlap.startColumn)
    result.push({
      startRow: overlap.startRow,
      endRow: overlap.endRow,
      startColumn: range.startColumn,
      endColumn: overlap.startColumn - 1,
    })
  if (overlap.endColumn < range.endColumn)
    result.push({
      startRow: overlap.startRow,
      endRow: overlap.endRow,
      startColumn: overlap.endColumn + 1,
      endColumn: range.endColumn,
    })
  return result
}

function terminal(
  status: ErrorScanResult['status'],
  findings: WorkbookErrorFinding[],
  cells: number,
  batches: number,
  bytes: number,
): ErrorScanResult {
  findings.sort(
    (left, right) =>
      left.sheetId.localeCompare(right.sheetId) ||
      left.address.localeCompare(right.address) ||
      left.errorCode.localeCompare(right.errorCode),
  )
  return { status, findings, scanned: { cells, batches, bytes } }
}

/**
 * Scans workbook data rather than the rendered viewport. Hidden worksheets,
 * filtered-out rows, and manually hidden rows intentionally remain in scope:
 * Error Checking diagnoses stored workbook values, not only visible cells.
 */
export async function scanStreamedWorkbookErrors(
  state: LazyWorkbookState,
  identity: ErrorScanIdentity,
  supplied: ErrorScanBudgets = ERROR_SCAN_DEFAULT_BUDGETS,
): Promise<ErrorScanResult> {
  const now = supplied.now ?? Date.now
  const started = now()
  const findings: WorkbookErrorFinding[] = []
  let cells = 0
  let batches = 0
  let bytes = 0
  const stamp = journalStamp(state)
  const scanRevision = state.scanRevision ?? 0
  const stopped = (): ErrorScanResult['status'] | null => {
    if (identity.signal?.aborted) return 'cancelled'
    if (
      !identity.isCurrent() ||
      journalStamp(state) !== stamp ||
      (state.scanRevision ?? 0) !== scanRevision
    )
      return 'unavailable'
    if (now() - started >= supplied.maxMs) return 'truncated'
    return null
  }
  const original = new Map(state.file.sheets.map((sheet) => [sheet.id, sheet]))
  const targetIds = [
    ...state.file.sheets.map((sheet) => sheet.id),
    ...[...state.editJournal.sheets.added.keys()].filter((id) => !original.has(id)),
  ].filter((id) => !state.editJournal.sheets.removed.has(id))

  for (const sheetId of targetIds) {
    const worksheet = identity.worksheet(sheetId)
    if (!worksheet || !identity.isCurrent(sheetId, worksheet))
      return terminal('unavailable', [], cells, batches, bytes)
    const cellMatrix = (
      worksheet as {
        getSheet?: () => {
          getCellMatrix?: () => { getValue?: (row: number, column: number) => unknown }
        }
      }
    )
      .getSheet?.()
      ?.getCellMatrix?.()
    const liveCellAt = (row: number, column: number) =>
      cellMatrix?.getValue?.(row, column) as { f?: unknown; v?: unknown } | undefined
    const journal = state.editJournal.cells.get(sheetId)
    const shadowed = new Set<string>()
    for (const entry of journal?.values() ?? []) {
      const stop = stopped()
      if (stop) return terminal(stop, stop === 'unavailable' ? [] : findings, cells, batches, bytes)
      if (cells >= supplied.maxCells) return terminal('truncated', findings, cells, batches, bytes)
      cells += 1
      if (!entry.hasValue) continue
      shadowed.add(`${entry.row}:${entry.column}`)
      // Formula journal entries carry the authored formula but their saved
      // value can be null; Univer owns the current computed result.
      const current = entry.formula ? liveCellAt(entry.row, entry.column) : undefined
      const code = entry.formula ? errorCode(current?.v) : null
      if (code)
        findings.push({ sheetId, address: formatAddress(entry.row, entry.column), errorCode: code })
      if (findings.length >= supplied.maxFindings)
        return terminal('truncated', findings, cells, batches, bytes)
    }
    const meta = original.get(sheetId)
    if (!meta && !state.flags.preloadComplete) {
      // A newly duplicated sheet may contain inherited cells that are not in
      // the journal. Without file metadata it cannot be declared complete.
      if (state.editJournal.sheets.added.get(sheetId)?.sourceSheetId)
        return terminal('unavailable', [], cells, batches, bytes)
      continue
    }
    const ops = state.editJournal.structuralOps.get(sheetId) ?? []
    const liveWorksheet = worksheet as { getMaxRows?: () => number; getMaxColumns?: () => number }
    const rows = meta
      ? Math.max(0, meta.rowCount + netAxisDelta(ops, 'row'))
      : Math.max(0, liveWorksheet.getMaxRows?.() ?? 0)
    const columns = meta
      ? Math.max(0, meta.columnCount + netAxisDelta(ops, 'column'))
      : Math.max(0, liveWorksheet.getMaxColumns?.() ?? 0)
    if (!rows || !columns) continue
    const fullRange: IRange = {
      startRow: 0,
      endRow: rows - 1,
      startColumn: 0,
      endColumn: columns - 1,
    }
    const loaded = state.flags.preloadComplete
      ? fullRange
      : state.loadedRanges.get(sheetId)
        ? intersectRange(fullRange, state.loadedRanges.get(sheetId)!)
        : null
    if (loaded) {
      for (
        let startColumn = loaded.startColumn;
        startColumn <= loaded.endColumn;
        startColumn += supplied.batchCells
      ) {
        const endColumn = Math.min(loaded.endColumn, startColumn + supplied.batchCells - 1)
        const loadedWidth = endColumn - startColumn + 1
        const loadedRowsPerBatch = Math.max(1, Math.floor(supplied.batchCells / loadedWidth))
        for (
          let startRow = loaded.startRow;
          startRow <= loaded.endRow;
          startRow += loadedRowsPerBatch
        ) {
          const stop = stopped()
          if (stop)
            return terminal(stop, stop === 'unavailable' ? [] : findings, cells, batches, bytes)
          const endRow = Math.min(loaded.endRow, startRow + loadedRowsPerBatch - 1)
          if (batches >= supplied.maxBatches)
            return terminal('truncated', findings, cells, batches, bytes)
          batches += 1
          for (let row = startRow; row <= endRow; row += 1) {
            for (let column = startColumn; column <= endColumn; column += 1) {
              if (cells >= supplied.maxCells)
                return terminal('truncated', findings, cells, batches, bytes)
              cells += 1
              if (shadowed.has(`${row}:${column}`)) continue
              const cell = liveCellAt(row, column)
              const code = typeof cell?.f === 'string' ? errorCode(cell.v) : null
              if (code)
                findings.push({ sheetId, address: formatAddress(row, column), errorCode: code })
              if (findings.length >= supplied.maxFindings)
                return terminal('truncated', findings, cells, batches, bytes)
            }
          }
          const afterLoaded = stopped()
          if (afterLoaded)
            return terminal(
              afterLoaded,
              afterLoaded === 'unavailable' ? [] : findings,
              cells,
              batches,
              bytes,
            )
          identity.onProgress?.({ cells, batches, bytes })
        }
      }
    }
    const sidecarRanges = !meta ? [] : loaded ? subtractRange(fullRange, loaded) : [fullRange]
    for (const scanRange of sidecarRanges) {
      for (
        let startColumn = scanRange.startColumn;
        startColumn <= scanRange.endColumn;
        startColumn += supplied.batchCells
      ) {
        const endColumn = Math.min(scanRange.endColumn, startColumn + supplied.batchCells - 1)
        const width = endColumn - startColumn + 1
        const rowsPerBatch = Math.max(1, Math.floor(supplied.batchCells / width))
        for (
          let startRow = scanRange.startRow;
          startRow <= scanRange.endRow;
          startRow += rowsPerBatch
        ) {
          const stop = stopped()
          if (stop)
            return terminal(stop, stop === 'unavailable' ? [] : findings, cells, batches, bytes)
          const endRow = Math.min(scanRange.endRow, startRow + rowsPerBatch - 1)
          if (cells >= supplied.maxCells || batches >= supplied.maxBatches)
            return terminal('truncated', findings, cells, batches, bytes)
          let mapped
          try {
            mapped = await readSheetRangeMapped(
              state,
              sheetId,
              { startRow, endRow, startColumn, endColumn },
              meta!,
              {
                ...(identity.signal ? { signal: identity.signal } : {}),
                maxBytes: Math.max(0, supplied.maxBytes - bytes),
                maxMs: Math.max(0, supplied.maxMs - (now() - started)),
                maxBatches: Math.max(0, supplied.maxBatches - batches),
                maxCells: Math.max(0, supplied.maxCells - cells),
                isCurrent: () =>
                  identity.isCurrent(sheetId, worksheet) &&
                  journalStamp(state) === stamp &&
                  (state.scanRevision ?? 0) === scanRevision,
              },
            )
          } catch (error: unknown) {
            const afterFailure = stopped()
            const status =
              afterFailure ??
              (error instanceof MappedRangeByteBudgetError ||
              error instanceof MappedRangeRequestBudgetError
                ? 'truncated'
                : 'unavailable')
            return terminal(
              status,
              afterFailure === 'unavailable' ? [] : findings,
              cells,
              batches,
              bytes,
            )
          }
          const after = stopped()
          if (after)
            return terminal(after, after === 'unavailable' ? [] : findings, cells, batches, bytes)
          if (!identity.isCurrent(sheetId, worksheet))
            return terminal('unavailable', [], cells, batches, bytes)
          if (!mapped) continue
          const acceptedBytes =
            mapped.byteCount ?? new TextEncoder().encode(JSON.stringify(mapped.screen)).byteLength
          if (bytes + acceptedBytes > supplied.maxBytes)
            return terminal('truncated', findings, cells, batches, bytes)
          bytes += acceptedBytes
          cells += mapped.requestedCellCount
          batches += mapped.requestBatchCount
          identity.onProgress?.({ cells, batches, bytes })
          if (!mapped.raw.indexingComplete && (mapped.indexedThroughScreen ?? -1) < endRow)
            return terminal('truncated', findings, cells, batches, bytes)
          for (const cell of mapped.screen.cells) {
            if (shadowed.has(`${cell.row}:${cell.column}`)) continue
            // Frozen strips and formula-closure cells can be live outside the
            // main loaded rectangle. Any live formula shadows the file cache.
            const live = liveCellAt(cell.row, cell.column)
            const code =
              typeof live?.f === 'string'
                ? errorCode(live.v)
                : cell.formula
                  ? errorCode(cell.value)
                  : null
            if (!code) continue
            findings.push({
              sheetId,
              address: formatAddress(cell.row, cell.column),
              errorCode: code,
            })
            if (findings.length >= supplied.maxFindings)
              return terminal('truncated', findings, cells, batches, bytes)
          }
        }
      }
    }
  }
  return terminal('complete', findings, cells, batches, bytes)
}

function parseAddress(address: string): [number, number] {
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(address)
  if (!match) return [Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]
  let column = 0
  for (const char of match[1]!) column = column * 26 + char.charCodeAt(0) - 64
  return [Number(match[2]) - 1, column - 1]
}

export function pickNextWorkbookError(
  findings: readonly WorkbookErrorFinding[],
  sheetOrder: ReadonlyMap<string, number>,
  activeSheetId: string,
  activeRow: number,
  activeColumn: number,
): WorkbookErrorFinding | null {
  const rank = (finding: WorkbookErrorFinding): [number, number, number] => {
    const [row, column] = parseAddress(finding.address)
    return [sheetOrder.get(finding.sheetId) ?? Number.MAX_SAFE_INTEGER, row, column]
  }
  const compare = (a: readonly number[], b: readonly number[]) =>
    (a[0] ?? 0) - (b[0] ?? 0) || (a[1] ?? 0) - (b[1] ?? 0) || (a[2] ?? 0) - (b[2] ?? 0)
  const ordered = [...findings].sort((a, b) => compare(rank(a), rank(b)))
  const active = [sheetOrder.get(activeSheetId) ?? Number.MAX_SAFE_INTEGER, activeRow, activeColumn]
  return ordered.find((finding) => compare(rank(finding), active) > 0) ?? ordered[0] ?? null
}

export interface StreamedErrorCheckDeps {
  readonly runtime: UniverRuntime
  readonly lazyWorkbookRef: { current: LazyWorkbookState | null }
  readonly setMessage: (message: string) => void
  readonly refreshSelectionEcho: () => void
}

const activeRuns = new Map<string, AbortController>()

/** A second invocation cancels the current workbook's scan; other workbooks are independent. */
export async function runStreamedErrorCheck(deps: StreamedErrorCheckDeps): Promise<void> {
  const state = deps.lazyWorkbookRef.current
  const workbook = deps.runtime.univerAPI.getActiveWorkbook()
  if (!state || !workbook || workbook.getId() !== state.expectedWorkbookId) return
  const workbookId = state.expectedWorkbookId
  const running = activeRuns.get(workbookId)
  if (running) {
    running.abort()
    deps.setMessage(t('appCancel'))
    return
  }
  const controller = new AbortController()
  activeRuns.set(workbookId, controller)
  const currentWorkbook = () => deps.runtime.univerAPI.getActiveWorkbook()
  const isCurrent = (sheetId?: string, worksheet?: unknown): boolean => {
    const current = currentWorkbook()
    if (
      controller.signal.aborted ||
      deps.lazyWorkbookRef.current !== state ||
      current?.getId() !== workbookId ||
      state.file.sessionId !== deps.lazyWorkbookRef.current?.file.sessionId ||
      state.generation !== deps.lazyWorkbookRef.current?.generation
    )
      return false
    if (!sheetId) return true
    const target = current.getSheetBySheetId(sheetId)
    return (
      Boolean(target) &&
      target?.getSheetId() ===
        (worksheet as { getSheetId?: () => string } | undefined)?.getSheetId?.() &&
      isLazyWorkbookTargetCurrent(deps.runtime, deps.lazyWorkbookRef, state, target!)
    )
  }
  try {
    deps.setMessage(t('appErrorChecking'))
    const result = await scanStreamedWorkbookErrors(state, {
      signal: controller.signal,
      isCurrent,
      worksheet: (sheetId) => workbook.getSheetBySheetId(sheetId),
      onProgress: ({ cells }) =>
        deps.setMessage(`${t('appErrorChecking')} · ${cells.toLocaleString()}`),
    })
    if (activeRuns.get(workbookId) !== controller || !isCurrent()) return
    if (result.status === 'cancelled') {
      deps.setMessage(t('appCancel'))
      return
    }
    if (result.status === 'unavailable') {
      deps.setMessage(t('appNotAvailableYet'))
      return
    }
    if (result.status === 'truncated') {
      deps.setMessage(t('appFindScanTruncated', { cells: result.scanned.cells }))
      return
    }
    if (result.findings.length === 0) {
      deps.setMessage(`${t('appErrorChecking')}: 0`)
      return
    }
    const sheets = workbook.getSheets()
    const order = new Map(sheets.map((sheet, index) => [sheet.getSheetId(), index] as const))
    const activeSheet = workbook.getActiveSheet()
    const activeRange = workbook.getActiveRange()
    const next = pickNextWorkbookError(
      result.findings,
      order,
      activeSheet?.getSheetId() ?? '',
      activeRange?.getRow() ?? -1,
      activeRange?.getColumn() ?? -1,
    )
    if (!next) return
    const target = workbook.getSheetBySheetId(next.sheetId)
    if (!target || !isCurrent(next.sheetId, target)) {
      deps.setMessage(t('appNotAvailableYet'))
      return
    }
    const [row, column] = parseAddress(next.address)
    if (target.isSheetHidden() || !target.getSheet().getRowVisible(row)) {
      deps.setMessage(
        `${t('appErrorChecking')}: ${result.findings.length} · ${next.address} · ${next.errorCode} · ${t('appNotAvailableYet')}`,
      )
      return
    }
    const range: IRange = { startRow: row, endRow: row, startColumn: column, endColumn: column }
    const loaded = await ensureLazyRangeLoaded(
      deps.runtime,
      deps.lazyWorkbookRef,
      target,
      range,
      () => undefined,
    )
    if (!loaded || !isCurrent(next.sheetId, target)) {
      deps.setMessage(t('appNotAvailableYet'))
      return
    }
    try {
      if (target !== activeSheet) workbook.setActiveSheet(target)
      if (!isCurrent(next.sheetId, target)) return
      target.getRange(row, column, 1, 1).activate()
      deps.refreshSelectionEcho()
      await deps.runtime.univerAPI.executeCommand('sheet.command.scroll-to-cell', { range })
      if (isCurrent(next.sheetId, target))
        deps.setMessage(
          `${t('appErrorChecking')}: ${result.findings.length} · ${next.address} · ${next.errorCode}`,
        )
    } catch {
      if (isCurrent()) deps.setMessage(t('appNotAvailableYet'))
    }
  } finally {
    if (activeRuns.get(workbookId) === controller) activeRuns.delete(workbookId)
  }
}
