import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import {
  SIDECAR_READ_BATCH_CELLS,
  SIDECAR_READ_MAX_TOTAL_CELLS,
  ensureLazyRangeLoaded,
  isLazyWorkbookTargetCurrent,
  readSheetRangeMapped,
} from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

type RangeCall = { sessionId: string; sheetId: string; range: Record<string, number> }

const rangeResult = async (call: RangeCall) => ({
  cells: [
    {
      row: call.range.startRow,
      column: call.range.startColumn,
      value: call.range.startRow,
      formula: `=${call.range.startRow}`,
      styleIndex: call.range.startRow,
    },
  ],
  rows: [{ row: call.range.startRow, hidden: false }],
  merges: [{ startRow: 0, endRow: 299, startColumn: 0, endColumn: 0 }],
  hyperlinks: [
    {
      row: call.range.startRow,
      column: call.range.startColumn,
      target: `row-${call.range.startRow}`,
    },
  ],
  conditionalRules: [],
  autoFilter: null,
  dataValidations: [],
  sheetProtection: null,
  indexedThroughRow: call.range.endRow,
  indexingComplete: true,
})
const readWorkbookRange = vi.fn(rangeResult)

function state(sessionId = 'session-1'): LazyWorkbookState {
  return {
    file: { sessionId, sheets: [] },
    expectedWorkbookId: 'file-sha-one',
    generation: 1,
    editJournal: { cells: new Map(), structuralOps: new Map() },
  } as unknown as LazyWorkbookState
}

const sheetMeta = {
  id: 's1',
  name: 'Sheet1',
  rowCount: 100_000,
  columnCount: 20_000,
} as unknown as LazyWorkbookState['file']['sheets'][number]

describe('readSheetRangeMapped bounded batching', () => {
  beforeEach(() => {
    readWorkbookRange.mockReset().mockImplementation(rangeResult)
    vi.stubGlobal('window', { desktopApi: { readWorkbookRange } })
  })

  afterEach(() => vi.unstubAllGlobals())

  it('reads a large viewport in ordered row batches without duplicates', async () => {
    const result = await readSheetRangeMapped(
      state(),
      's1',
      { startRow: 0, endRow: 299, startColumn: 0, endColumn: 199 },
      sheetMeta,
    )

    expect(readWorkbookRange).toHaveBeenCalledTimes(4)
    expect(readWorkbookRange.mock.calls.map(([call]) => call.range)).toEqual([
      { startRow: 0, endRow: 89, startColumn: 0, endColumn: 199 },
      { startRow: 90, endRow: 179, startColumn: 0, endColumn: 199 },
      { startRow: 180, endRow: 269, startColumn: 0, endColumn: 199 },
      { startRow: 270, endRow: 299, startColumn: 0, endColumn: 199 },
    ])
    expect(result?.screen.cells.map((cell) => cell.row)).toEqual([0, 90, 180, 270])
    expect(result?.screen.cells.map((cell) => cell.styleIndex)).toEqual([0, 90, 180, 270])
    expect(result?.screen.cells.map((cell) => cell.formula)).toEqual(['=0', '=90', '=180', '=270'])
    expect(result?.screen.merges).toHaveLength(1)
    expect(result?.byteCount).toBeGreaterThan(0)
    expect(result?.truncated).toBeNull()
  })

  it('keeps an exact-cap request in one batch', async () => {
    await readSheetRangeMapped(
      state(),
      's1',
      { startRow: 0, endRow: 89, startColumn: 0, endColumn: 199 },
      sheetMeta,
    )
    expect(readWorkbookRange).toHaveBeenCalledTimes(1)
  })

  it('fails closed when one row is wider than the batch cap', async () => {
    await expect(
      readSheetRangeMapped(
        state(),
        's1',
        { startRow: 0, endRow: 0, startColumn: 0, endColumn: SIDECAR_READ_BATCH_CELLS },
        sheetMeta,
      ),
    ).rejects.toThrow('single row')
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

  it('fails before I/O when the total read budget is exceeded', async () => {
    const rows = Math.floor(SIDECAR_READ_MAX_TOTAL_CELLS / 200) + 1
    await expect(
      readSheetRangeMapped(
        state(),
        's1',
        { startRow: 0, endRow: rows - 1, startColumn: 0, endColumn: 199 },
        sheetMeta,
      ),
    ).rejects.toThrow('total limit')
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

  it('returns an explicit indexing receipt instead of silently truncating', async () => {
    readWorkbookRange.mockImplementationOnce(async (call) => ({
      ...(await rangeResult(call)),
      indexedThroughRow: 20,
      indexingComplete: false,
    }))
    const result = await readSheetRangeMapped(
      state(),
      's1',
      { startRow: 0, endRow: 299, startColumn: 0, endColumn: 199 },
      sheetMeta,
    )
    expect(readWorkbookRange).toHaveBeenCalledTimes(1)
    expect(result?.truncated).toEqual({ reason: 'indexing', nextRow: 21 })
    expect(result?.indexedThroughScreen).toBe(20)
  })

  it('rejects the whole read when a later batch throws', async () => {
    readWorkbookRange.mockImplementationOnce(rangeResult)
    readWorkbookRange.mockRejectedValueOnce(new Error('batch two failed'))
    await expect(
      readSheetRangeMapped(
        state(),
        's1',
        { startRow: 0, endRow: 299, startColumn: 0, endColumn: 199 },
        sheetMeta,
      ),
    ).rejects.toThrow('batch two failed')
  })

  it('aborts between batches without returning partial data', async () => {
    const controller = new AbortController()
    readWorkbookRange.mockImplementationOnce(async (call) => {
      controller.abort()
      return {
        cells: [],
        rows: [],
        merges: [],
        hyperlinks: [],
        conditionalRules: [],
        autoFilter: null,
        dataValidations: [],
        sheetProtection: null,
        indexedThroughRow: call.range.endRow,
        indexingComplete: true,
      }
    })
    await expect(
      readSheetRangeMapped(
        state(),
        's1',
        { startRow: 0, endRow: 299, startColumn: 0, endColumn: 199 },
        sheetMeta,
        { signal: controller.signal },
      ),
    ).rejects.toMatchObject({ name: 'AbortError' })
    expect(readWorkbookRange).toHaveBeenCalledTimes(1)
  })

  it('rejects if workbook identity drifts after a batch', async () => {
    let current = true
    readWorkbookRange.mockImplementationOnce(async (call) => {
      current = false
      return {
        cells: [],
        rows: [],
        merges: [],
        hyperlinks: [],
        conditionalRules: [],
        autoFilter: null,
        dataValidations: [],
        sheetProtection: null,
        indexedThroughRow: call.range.endRow,
        indexingComplete: true,
      }
    })
    await expect(
      readSheetRangeMapped(
        state(),
        's1',
        { startRow: 0, endRow: 299, startColumn: 0, endColumn: 199 },
        sheetMeta,
        { isCurrent: () => current },
      ),
    ).rejects.toThrow('workbook changed')
    expect(readWorkbookRange).toHaveBeenCalledTimes(1)
  })

  it('rejects an old worksheet callback after the active workbook switches before I/O', async () => {
    const oldState = state()
    const newState = { ...state('session-2'), expectedWorkbookId: 'file-sha-two' }
    const oldCore = { getUnitId: () => oldState.expectedWorkbookId }
    const newCore = { getUnitId: () => newState.expectedWorkbookId }
    const active = {
      getId: () => newState.expectedWorkbookId,
      getWorkbook: () => newCore,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => active } }
    const oldWorksheet = { getWorkbook: () => oldCore }
    const lazyRef = { current: newState }

    await expect(
      readSheetRangeMapped(
        oldState,
        's1',
        { startRow: 0, endRow: 299, startColumn: 0, endColumn: 199 },
        sheetMeta,
        {
          isCurrent: () =>
            isLazyWorkbookTargetCurrent(
              runtime as never,
              lazyRef as never,
              oldState,
              oldWorksheet as never,
            ),
        },
      ),
    ).rejects.toThrow('workbook changed')
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

  it('does not patch when an old worksheet callback enters after a switch', async () => {
    const newState = {
      ...state('session-2'),
      expectedWorkbookId: 'file-sha-two',
      file: { ...state('session-2').file, sheets: [sheetMeta] },
      loadedRanges: new Map(),
    } as unknown as LazyWorkbookState
    const oldCore = { getUnitId: () => 'file-sha-one' }
    const newCore = { getUnitId: () => newState.expectedWorkbookId }
    const getRange = vi.fn()
    const oldWorksheet = {
      getSheetId: () => 's1',
      getWorkbook: () => oldCore,
      getRange,
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({
          getId: () => newState.expectedWorkbookId,
          getWorkbook: () => newCore,
        }),
      },
    }

    await expect(
      ensureLazyRangeLoaded(
        runtime as never,
        { current: newState },
        oldWorksheet as never,
        { startRow: 0, endRow: 10, startColumn: 0, endColumn: 10 },
        vi.fn(),
      ),
    ).resolves.toBe(false)
    expect(readWorkbookRange).not.toHaveBeenCalled()
    expect(getRange).not.toHaveBeenCalled()
  })

  it('rejects a wrong state even while its worksheet workbook is still active', async () => {
    const activeState = state()
    const wrongState = { ...state('session-2'), expectedWorkbookId: 'file-sha-two' }
    const activeCore = { getUnitId: () => activeState.expectedWorkbookId }
    const active = {
      getId: () => activeState.expectedWorkbookId,
      getWorkbook: () => activeCore,
    }
    const runtime = { univerAPI: { getActiveWorkbook: () => active } }
    const worksheet = { getWorkbook: () => activeCore }
    const lazyRef = { current: wrongState }

    expect(
      isLazyWorkbookTargetCurrent(
        runtime as never,
        lazyRef as never,
        wrongState,
        worksheet as never,
      ),
    ).toBe(false)
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })
})
