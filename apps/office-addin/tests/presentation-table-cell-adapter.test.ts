import { afterEach, describe, expect, it, vi } from 'vitest'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'

afterEach(() => vi.unstubAllGlobals())

function setup() {
  let text = 'old',
    writes = 0
  const cell = {
    isNullObject: false,
    rowIndex: 1,
    columnIndex: 0,
    rowCount: 1,
    columnCount: 1,
    load: vi.fn(),
    get text() {
      return text
    },
    set text(value: string) {
      writes++
      text = value
    },
  }
  const table = {
    rowCount: 2,
    columnCount: 2,
    load: vi.fn(),
    getCellOrNullObject: vi.fn(() => cell),
  }
  const shape = { id: 'table-1', type: 'Table', load: vi.fn(), getTable: vi.fn(() => table) }
  const slide = { id: 'slide-1', load: vi.fn(), shapes: { getItem: vi.fn(() => shape) } }
  const slides = {
    getItem: vi.fn(() => slide),
    getItemAt: vi.fn(() => {
      throw new Error('index forbidden')
    }),
  }
  const context = { presentation: { slides }, sync: vi.fn(async () => {}) }
  const supports = vi.fn(() => true)
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: supports } },
  })
  vi.stubGlobal('PowerPoint', {
    run: vi.fn(async (callback: (value: typeof context) => Promise<unknown>) => callback(context)),
  })
  return {
    adapter: new BrowserPowerPointAdapter(),
    cell,
    table,
    shape,
    slide,
    slides,
    context,
    supports,
    writes: () => writes,
    setText: (value: string) => {
      text = value
    },
  }
}

describe('native PowerPoint table cell adapter', () => {
  it('reads and writes one exact unmerged cell through PowerPointApi 1.8', async () => {
    const f = setup()
    expect(await f.adapter.readPresentationTableCell('slide-1', 'table-1', 1, 0)).toEqual({
      slideId: 'slide-1',
      shapeId: 'table-1',
      rowIndex: 1,
      columnIndex: 0,
      text: 'old',
      rowCount: 2,
      columnCount: 2,
    })
    await f.adapter.editPresentationTableCell('slide-1', 'table-1', 1, 0, 'new', 'old')
    expect(f.cell.text).toBe('new')
    expect(f.writes()).toBe(1)
    expect(f.slides.getItem).toHaveBeenCalledWith('slide-1')
    expect(f.slides.getItemAt).not.toHaveBeenCalled()
    expect(f.table.getCellOrNullObject).toHaveBeenCalledWith(1, 0)
    expect(f.supports).toHaveBeenCalledWith('PowerPointApi', '1.8')
  })
  it('rejects stale values, wrong types, invalid indices, and merged cells before writing', async () => {
    const f = setup()
    await expect(
      f.adapter.editPresentationTableCell('slide-1', 'table-1', 1, 0, 'new', 'stale'),
    ).rejects.toThrow('office_concurrent_change')
    await expect(
      f.adapter.editPresentationTableCell('slide-1', 'table-1', -1, 0, 'new', 'old'),
    ).rejects.toThrow('invalid_tool_input')
    await expect(
      f.adapter.editPresentationTableCell('slide-1', 'table-1', 2, 0, 'new', 'old'),
    ).rejects.toThrow('invalid_tool_input')
    f.cell.rowCount = 2
    await expect(
      f.adapter.editPresentationTableCell('slide-1', 'table-1', 1, 0, 'new', 'old'),
    ).rejects.toThrow('office_api_unsupported')
    f.cell.rowCount = 1
    f.cell.isNullObject = true
    await expect(f.adapter.readPresentationTableCell('slide-1', 'table-1', 1, 0)).rejects.toThrow(
      'office_api_unsupported',
    )
    f.cell.isNullObject = false
    f.shape.type = 'TextBox'
    await expect(f.adapter.readPresentationTableCell('slide-1', 'table-1', 1, 0)).rejects.toThrow(
      'office_api_unsupported',
    )
    expect(f.writes()).toBe(0)
  })
  it('reconciles a rejected sync against the observed cell text', async () => {
    const f = setup()
    let failed = false
    f.context.sync.mockImplementation(async () => {
      if (f.writes() && !failed) {
        failed = true
        f.setText('user')
        throw new Error('sync rejected')
      }
    })
    await expect(
      f.adapter.editPresentationTableCell('slide-1', 'table-1', 1, 0, 'new', 'old'),
    ).rejects.toThrow('office_concurrent_change')
    expect(f.cell.text).toBe('user')
    expect(f.writes()).toBe(1)
  })
})
