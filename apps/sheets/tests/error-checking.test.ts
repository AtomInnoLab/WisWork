import { beforeEach, describe, expect, it, vi } from 'vitest'

import {
  ERROR_SCAN_DEFAULT_BUDGETS,
  pickNextWorkbookError,
  scanStreamedWorkbookErrors,
} from '../src/renderer/error-checking'
import { readSheetRangeMapped } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

vi.mock('../src/renderer/univer-sync', () => ({
  MappedRangeByteBudgetError: class MappedRangeByteBudgetError extends Error {},
  MappedRangeRequestBudgetError: class MappedRangeRequestBudgetError extends Error {},
  readSheetRangeMapped: vi.fn(),
  ensureLazyRangeLoaded: vi.fn(),
}))

const readMapped = vi.mocked(readSheetRangeMapped)

function mapped(cells: Array<{ row: number; column: number; value: unknown; formula?: string }>) {
  return {
    raw: { indexingComplete: true },
    indexedThroughScreen: 99,
    fileEndRow: 99,
    byteCount: 100,
    requestBatchCount: 1,
    requestedCellCount: 200,
    screen: { cells, rows: [], merges: [], hyperlinks: [] },
  } as unknown as Awaited<ReturnType<typeof readSheetRangeMapped>>
}

function state(): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [
        { id: 'visible', name: 'Visible', rowCount: 100, columnCount: 2, hidden: false },
        { id: 'hidden', name: 'Hidden', rowCount: 100, columnCount: 2, hidden: true },
      ],
    },
    expectedWorkbookId: 'book-1',
    generation: 7,
    loadedRanges: new Map(),
    flags: { preloadComplete: false },
    editJournal: {
      cells: new Map(),
      structuralOps: new Map(),
      sheets: { removed: new Set(), hidden: new Map(), added: new Map(), renamed: new Map() },
      filterDirty: new Set(),
    },
    filterOrigins: new Map(),
  } as unknown as LazyWorkbookState
}

function identity(lazy: LazyWorkbookState) {
  const sheets = new Map(lazy.file.sheets.map((sheet) => [sheet.id, {}]))
  return {
    isCurrent: (sheetId?: string, worksheet?: unknown) =>
      lazy.file.sessionId === 'session-1' &&
      lazy.generation === 7 &&
      (!sheetId || sheets.get(sheetId) === worksheet),
    worksheet: (sheetId: string) => sheets.get(sheetId),
  }
}

describe('scanStreamedWorkbookErrors', () => {
  beforeEach(() => readMapped.mockReset())

  it('scans unloaded visible and hidden sheets in stable order and emits safe codes', async () => {
    const lazy = state()
    const guard = identity(lazy)
    readMapped.mockImplementation(async (_state, sheetId) =>
      mapped(
        sheetId === 'visible'
          ? [
              { row: 9, column: 1, value: '#REF!', formula: '=SECRET()' },
              { row: 1, column: 0, value: '#N/A' },
            ]
          : [{ row: 0, column: 0, value: '#DIV/0!', formula: '=1/0' }],
      ),
    )
    const result = await scanStreamedWorkbookErrors(lazy, guard)
    expect(result.status).toBe('complete')
    expect(result.findings).toEqual([
      { sheetId: 'hidden', address: 'A1', errorCode: '#DIV/0!' },
      { sheetId: 'visible', address: 'B10', errorCode: '#REF!' },
    ])
    expect(JSON.stringify(result)).not.toContain('SECRET')
  })

  it('lets journal values shadow stale sidecar values and scans journal formulas', async () => {
    const lazy = state()
    lazy.editJournal.cells.set(
      'visible',
      new Map([
        ['a', { row: 3, column: 0, hasValue: true, value: 'fixed' }],
        ['b', { row: 4, column: 1, hasValue: true, value: null, formula: '=NA()' }],
        ['c', { row: 6, column: 0, hasValue: false, value: null }],
      ]),
    )
    const guard = identity(lazy)
    const visible = guard.worksheet('visible') as { getSheet?: () => unknown }
    visible.getSheet = () => ({
      getCellMatrix: () => ({
        getValue: (row: number, column: number) =>
          row === 4 && column === 1 ? { f: '=NA()', v: '#N/A' } : undefined,
      }),
    })
    readMapped.mockImplementation(async (_state, sheetId) =>
      mapped(
        sheetId === 'visible'
          ? [
              { row: 3, column: 0, value: '#VALUE!', formula: '=stale()' },
              { row: 6, column: 0, value: '#REF!', formula: '=kept()' },
            ]
          : [],
      ),
    )
    const result = await scanStreamedWorkbookErrors(lazy, guard)
    expect(result.findings).toEqual([
      { sheetId: 'visible', address: 'A7', errorCode: '#REF!' },
      { sheetId: 'visible', address: 'B5', errorCode: '#N/A' },
    ])
  })

  it('uses loaded Univer computed formula state before sidecar in both directions', async () => {
    const lazy = state()
    lazy.loadedRanges.set('visible', {
      startRow: 0,
      endRow: 0,
      startColumn: 0,
      endColumn: 0,
    })
    const visible = identity(lazy).worksheet('visible') as {
      getSheet?: () => unknown
    }
    visible.getSheet = () => ({
      getCellMatrix: () => ({
        getValue: (row: number, column: number) =>
          row === 0 && column === 0
            ? { f: '=1/0', v: '#DIV/0!' }
            : row === 0 && column === 1
              ? { f: '=fixed()', v: 42 }
              : row === 1 && column === 1
                ? { f: '=recalc()', v: '#VALUE!' }
                : undefined,
      }),
    })
    const guard = identity(lazy)
    const guardedVisible = guard.worksheet('visible') as typeof visible
    guardedVisible.getSheet = visible.getSheet
    readMapped.mockImplementation(async (_state, sheetId, range) =>
      mapped(
        sheetId === 'visible' && range.startColumn <= 1 && range.endColumn >= 1
          ? [
              { row: 0, column: 1, value: '#REF!', formula: '=stale()' },
              { row: 1, column: 1, value: 42, formula: '=stale-success()' },
            ]
          : [],
      ),
    )
    const result = await scanStreamedWorkbookErrors(lazy, guard)
    expect(result.findings).toContainEqual({
      sheetId: 'visible',
      address: 'A1',
      errorCode: '#DIV/0!',
    })
    expect(result.findings).not.toContainEqual(
      expect.objectContaining({ sheetId: 'visible', address: 'B1' }),
    )
    expect(result.findings).toContainEqual({
      sheetId: 'visible',
      address: 'B2',
      errorCode: '#VALUE!',
    })
  })

  it('recognizes the complete modern error-code set', async () => {
    const lazy = state()
    readMapped.mockImplementation(async (_state, sheetId) =>
      mapped(
        sheetId === 'visible'
          ? [
              { row: 0, column: 0, value: '#BUSY!', formula: '=busy()' },
              { row: 1, column: 0, value: '#PYTHON!', formula: '=PY()' },
            ]
          : [],
      ),
    )
    expect((await scanStreamedWorkbookErrors(lazy, identity(lazy))).findings).toEqual([
      { sheetId: 'visible', address: 'A1', errorCode: '#BUSY!' },
      { sheetId: 'visible', address: 'A2', errorCode: '#PYTHON!' },
    ])
  })

  it('scans every Univer sheet with zero sidecar reads after preload completes', async () => {
    const lazy = state()
    lazy.flags.preloadComplete = true
    const guard = identity(lazy)
    for (const sheetId of ['visible', 'hidden']) {
      const worksheet = guard.worksheet(sheetId) as { getSheet?: () => unknown }
      worksheet.getSheet = () => ({
        getCellMatrix: () => ({
          getValue: (row: number, column: number) =>
            row === 0 && column === 0
              ? { f: '=1/0', v: sheetId === 'visible' ? '#DIV/0!' : '#CALC!' }
              : undefined,
        }),
      })
    }
    const result = await scanStreamedWorkbookErrors(lazy, guard)
    expect(result.status).toBe('complete')
    expect(result.findings).toEqual([
      { sheetId: 'hidden', address: 'A1', errorCode: '#CALC!' },
      { sheetId: 'visible', address: 'A1', errorCode: '#DIV/0!' },
    ])
    expect(readMapped).not.toHaveBeenCalled()
  })

  it('scans a duplicated live sheet after preload without file metadata', async () => {
    const lazy = state()
    lazy.flags.preloadComplete = true
    lazy.editJournal.sheets.added.set('copy', { name: 'Copy', sourceSheetId: 'visible' })
    const guard = identity(lazy)
    const copy = {
      getMaxRows: () => 1,
      getMaxColumns: () => 1,
      getSheet: () => ({
        getCellMatrix: () => ({ getValue: () => ({ f: '=1/0', v: '#DIV/0!' }) }),
      }),
    }
    const result = await scanStreamedWorkbookErrors(lazy, {
      ...guard,
      worksheet: (sheetId) => (sheetId === 'copy' ? copy : guard.worksheet(sheetId)),
      isCurrent: (sheetId, worksheet) =>
        sheetId === 'copy' ? worksheet === copy : guard.isCurrent(sheetId, worksheet),
    })
    expect(result.findings).toContainEqual({
      sheetId: 'copy',
      address: 'A1',
      errorCode: '#DIV/0!',
    })
    expect(readMapped).not.toHaveBeenCalled()
  })

  it.each([
    ['cells', { maxCells: 1 }],
    ['findings', { maxFindings: 1 }],
    ['batches', { maxBatches: 1 }],
    ['bytes', { maxBytes: 1 }],
  ] as const)('reports explicit truncation at the %s cap', async (_name, budget) => {
    const lazy = state()
    readMapped.mockResolvedValue(mapped([{ row: 0, column: 0, value: '#REF!', formula: '=bad()' }]))
    const result = await scanStreamedWorkbookErrors(lazy, identity(lazy), {
      ...ERROR_SCAN_DEFAULT_BUDGETS,
      ...budget,
    })
    expect(result.status).toBe('truncated')
  })

  it('forwards the decreasing remaining byte acceptance budget', async () => {
    const lazy = state()
    readMapped.mockResolvedValue(mapped([]))
    await scanStreamedWorkbookErrors(lazy, identity(lazy), {
      ...ERROR_SCAN_DEFAULT_BUDGETS,
      batchCells: 100,
      maxBytes: 250,
    })
    expect(readMapped.mock.calls[0]?.[4]?.maxBytes).toBe(250)
    expect(readMapped.mock.calls[1]?.[4]?.maxBytes).toBe(150)
  })

  it('aborts on deadline and ignores a late batch', async () => {
    const lazy = state()
    let now = 0
    readMapped.mockImplementation(async () => {
      now = 10
      return mapped([{ row: 0, column: 0, value: '#REF!', formula: '=bad()' }])
    })
    const result = await scanStreamedWorkbookErrors(lazy, identity(lazy), {
      ...ERROR_SCAN_DEFAULT_BUDGETS,
      maxMs: 5,
      now: () => now,
    })
    expect(result).toMatchObject({ status: 'truncated', findings: [] })
  })

  it('fails closed when workbook, worksheet, filter, journal, or session drifts', async () => {
    const lazy = state()
    const guard = identity(lazy)
    readMapped.mockImplementationOnce(async () => {
      lazy.editJournal.filterDirty.add('visible')
      return mapped([{ row: 0, column: 0, value: '#REF!', formula: '=bad()' }])
    })
    const result = await scanStreamedWorkbookErrors(lazy, guard)
    expect(result).toMatchObject({ status: 'unavailable', findings: [] })
  })

  it('distinguishes an authoritative zero-error completion from unavailable', async () => {
    const lazy = state()
    readMapped.mockResolvedValue(mapped([]))
    expect(await scanStreamedWorkbookErrors(lazy, identity(lazy))).toMatchObject({
      status: 'complete',
      findings: [],
    })
    readMapped.mockRejectedValueOnce(new Error('raw sidecar detail'))
    expect(await scanStreamedWorkbookErrors(lazy, identity(lazy))).toMatchObject({
      status: 'unavailable',
      findings: [],
    })
  })
})

describe('pickNextWorkbookError', () => {
  it('orders by sheet/address and wraps', () => {
    const findings = [
      { sheetId: 'b', address: 'A1', errorCode: '#REF!' as const },
      { sheetId: 'a', address: 'B2', errorCode: '#N/A' as const },
    ]
    const order = new Map([
      ['a', 0],
      ['b', 1],
    ])
    expect(pickNextWorkbookError(findings, order, 'a', 1, 1)?.sheetId).toBe('b')
    expect(pickNextWorkbookError(findings, order, 'b', 0, 0)?.sheetId).toBe('a')
  })
})
