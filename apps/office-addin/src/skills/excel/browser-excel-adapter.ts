export const MAX_EXCEL_CELLS = 2_000
export const MAX_EXCEL_ROWS = 500
export const MAX_EXCEL_OBJECTS = 256

type RuntimeRecord = Record<string, any>
export interface RangeRequest {
  sheetId: number
  ranges: string[]
  includeStyles?: boolean
  cellLimit?: number
}
export interface CsvRequest {
  sheetId: number
  range: string
  includeHeaders?: boolean
  maxRows?: number
}
export interface SearchRequest {
  searchTerm: string
  sheetId?: number
  range?: string
  offset?: number
  options?: {
    matchCase?: boolean
    matchEntireCell?: boolean
    matchFormulas?: boolean
    useRegex?: boolean
    maxResults?: number
  }
}
export interface ObjectRequest {
  sheetId?: number
  id?: string
}
export interface ExcelMutation {
  tool: string
  input: Record<string, any>
  targets: string[]
}
export interface ExcelAdapter {
  getCellRanges(input: RangeRequest, signal?: AbortSignal): Promise<unknown>
  getRangeAsCsv(input: CsvRequest, signal?: AbortSignal): Promise<unknown>
  searchData(input: SearchRequest, signal?: AbortSignal): Promise<unknown>
  screenshotRange(
    input: { sheetId: number; range: string },
    signal?: AbortSignal,
  ): Promise<{ base64: string; mime: 'image/png' }>
  getAllObjects(input: ObjectRequest, signal?: AbortSignal): Promise<unknown>
  fingerprint(targets: string[], signal?: AbortSignal): Promise<string>
  setCellRange(input: Record<string, any>, signal?: AbortSignal): Promise<void>
  clearCellRange(input: Record<string, any>, signal?: AbortSignal): Promise<void>
  copyTo(input: Record<string, any>, signal?: AbortSignal): Promise<void>
  modifySheetStructure(input: Record<string, any>, signal?: AbortSignal): Promise<void>
  modifyWorkbookStructure(input: Record<string, any>, signal?: AbortSignal): Promise<void>
  resizeRange(input: Record<string, any>, signal?: AbortSignal): Promise<void>
  modifyObject(input: Record<string, any>, signal?: AbortSignal): Promise<void>
  verifyRanges(targets: string[], signal?: AbortSignal): Promise<unknown>
  verifyObjects(input: ObjectRequest, signal?: AbortSignal): Promise<unknown>
  verifyWorkbook(signal?: AbortSignal): Promise<unknown>
}

function cancelled(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('cancelled')
}
function runtime(): RuntimeRecord {
  const root = globalThis as unknown as RuntimeRecord
  const office = root.Office
  const excel = root.Excel
  const requirements = office?.context?.requirements
  if (
    !office ||
    !excel ||
    office.context?.host !== 'Excel' ||
    typeof requirements?.isSetSupported !== 'function' ||
    !requirements.isSetSupported('ExcelApi', '1.3') ||
    typeof excel.run !== 'function'
  )
    throw new Error('office_api_unsupported')
  return excel
}
async function sync(context: RuntimeRecord, signal?: AbortSignal) {
  cancelled(signal)
  await context.sync()
  cancelled(signal)
}
function sheet(context: RuntimeRecord, id: number): RuntimeRecord {
  return context.workbook.worksheets.getItem(String(id))
}
function cleanAddress(value: unknown): string {
  return (
    String(value ?? '')
      .split('!')
      .pop() ?? ''
  )
}
function csv(value: unknown): string {
  const text = value == null ? '' : String(value)
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}
function safe(value: unknown, max = 12_000): string {
  return typeof value === 'string' ? value.slice(0, max) : ''
}

export class BrowserExcelAdapter implements ExcelAdapter {
  private run<T>(callback: (context: RuntimeRecord) => Promise<T>): Promise<T> {
    return runtime().run(callback)
  }
  async getCellRanges(input: RangeRequest, signal?: AbortSignal): Promise<unknown> {
    cancelled(signal)
    return this.run(async (context) => {
      const ws = sheet(context, input.sheetId)
      ws.load('name')
      const limit = Math.min(input.cellLimit ?? MAX_EXCEL_CELLS, MAX_EXCEL_CELLS)
      const ranges: RuntimeRecord[] = input.ranges.map((address) => {
        const r = address === '*' ? ws.getUsedRangeOrNullObject() : ws.getRange(address)
        r.load('values,formulas,numberFormat,address,rowCount,columnCount')
        return r
      })
      await sync(context, signal)
      let remaining = limit
      let truncated = false
      const output = ranges.map((range) => {
        if (range.isNullObject)
          return {
            sheetId: input.sheetId,
            sheetName: safe(ws.name, 256),
            address: 'A1',
            rows: 0,
            columns: 0,
            cells: [],
          }
        const cells: unknown[] = []
        for (let row = 0; row < range.rowCount; row++)
          for (let column = 0; column < range.columnCount; column++) {
            if (remaining-- <= 0) {
              truncated = true
              break
            }
            cells.push({
              address: `${cleanAddress(range.address)}[${row},${column}]`,
              value: range.values?.[row]?.[column] ?? null,
              formula: range.formulas?.[row]?.[column] ?? null,
              numberFormat: range.numberFormat?.[row]?.[column] ?? null,
            })
          }
        return {
          sheetId: input.sheetId,
          sheetName: safe(ws.name, 256),
          address: cleanAddress(range.address),
          rows: range.rowCount,
          columns: range.columnCount,
          cells,
        }
      })
      return { ranges: output, hasMore: truncated }
    })
  }
  async getRangeAsCsv(input: CsvRequest, signal?: AbortSignal): Promise<unknown> {
    cancelled(signal)
    return this.run(async (context) => {
      const ws = sheet(context, input.sheetId)
      ws.load('name')
      const range = ws.getRange(input.range)
      range.load('values,rowCount,columnCount,address')
      await sync(context, signal)
      const start = input.includeHeaders === false ? 1 : 0
      const maximum = Math.min(input.maxRows ?? MAX_EXCEL_ROWS, MAX_EXCEL_ROWS)
      const end = Math.min(range.rowCount, start + maximum)
      const rows = range.values.slice(start, end).map((row: unknown[]) => row.map(csv).join(','))
      return {
        sheetId: input.sheetId,
        sheetName: safe(ws.name, 256),
        address: cleanAddress(range.address),
        csv: rows.join('\n'),
        rowCount: rows.length,
        columnCount: range.columnCount,
        hasMore: end < range.rowCount,
      }
    })
  }
  async searchData(input: SearchRequest, signal?: AbortSignal): Promise<unknown> {
    cancelled(signal)
    const maximum = Math.min(input.options?.maxResults ?? 500, 500)
    const offset = input.offset ?? 0
    const data = (await this.getCellRanges(
      {
        sheetId: input.sheetId ?? 1,
        ranges: [input.range ?? 'A1:XFD1048576'],
        cellLimit: Math.min(offset + maximum + 1, MAX_EXCEL_CELLS),
        includeStyles: false,
      },
      signal,
    )) as any
    const flags = input.options?.matchCase ? '' : 'i'
    let regex: RegExp
    try {
      regex = input.options?.useRegex
        ? new RegExp(input.searchTerm, flags)
        : new RegExp(input.searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags)
    } catch {
      throw new Error('invalid_tool_input')
    }
    const all = data.ranges
      .flatMap((r: any) => r.cells)
      .filter((c: any) => {
        const candidate = input.options?.matchFormulas ? c.formula : c.value
        const value = String(candidate ?? '')
        return input.options?.matchEntireCell
          ? input.options.matchCase
            ? value === input.searchTerm
            : value.toLowerCase() === input.searchTerm.toLowerCase()
          : regex.test(value)
      })
    const matches = all.slice(offset, offset + maximum)
    const hasMore = all.length > offset + maximum || data.hasMore
    return {
      matches,
      offset,
      returned: matches.length,
      hasMore,
      nextOffset: hasMore ? offset + matches.length : null,
    }
  }
  async screenshotRange(): Promise<{ base64: string; mime: 'image/png' }> {
    throw new Error('office_api_unsupported')
  }
  async getAllObjects(input: ObjectRequest, signal?: AbortSignal): Promise<unknown> {
    cancelled(signal)
    return this.run(async (context) => {
      const sheets =
        input.sheetId === undefined
          ? context.workbook.worksheets
          : { items: [sheet(context, input.sheetId)], load() {} }
      sheets.load?.({ $top: MAX_EXCEL_OBJECTS + 1 })
      await sync(context, signal)
      const objects: unknown[] = []
      for (const ws of (sheets.items ?? ([] as RuntimeRecord[])).slice(0, MAX_EXCEL_OBJECTS)) {
        ws.load('id,name')
        ws.charts.load({ $top: MAX_EXCEL_OBJECTS + 1 })
        ws.pivotTables.load({ $top: MAX_EXCEL_OBJECTS + 1 })
      }
      await sync(context, signal)
      for (const ws of sheets.items ?? []) {
        for (const object of [...(ws.charts?.items ?? []), ...(ws.pivotTables?.items ?? [])]) {
          object.load?.('id,name')
          objects.push({
            sheetId: ws.id,
            sheetName: safe(ws.name, 256),
            id: safe(object.id, 256),
            name: safe(object.name, 256),
            type: ws.charts?.items?.includes(object) ? 'chart' : 'pivotTable',
          })
        }
      }
      await sync(context, signal)
      const filtered = input.id ? objects.filter((item: any) => item.id === input.id) : objects
      return {
        objects: filtered.slice(0, MAX_EXCEL_OBJECTS),
        hasMore: filtered.length > MAX_EXCEL_OBJECTS,
      }
    })
  }
  async fingerprint(targets: string[], signal?: AbortSignal): Promise<string> {
    const values = await Promise.all(
      targets.map(async (target) => {
        if (target.startsWith('workbook:')) return this.verifyWorkbook(signal)
        const object = /^sheet:(\d+)!object:(.+)$/.exec(target)
        if (object)
          return this.verifyObjects(
            { sheetId: Number(object[1]), ...(object[2] === 'new' ? {} : { id: object[2] }) },
            signal,
          )
        return this.verifyRanges([target], signal)
      }),
    )
    return JSON.stringify(values).slice(0, 64 * 1024)
  }
  private async mutate(
    input: Record<string, any>,
    action: (ws: RuntimeRecord, context: RuntimeRecord) => void | Promise<void>,
    signal?: AbortSignal,
  ) {
    cancelled(signal)
    await this.run(async (context) => {
      const ws = sheet(context, input.sheetId)
      await action(ws, context)
      await sync(context, signal)
    })
  }
  setCellRange(input: Record<string, any>, signal?: AbortSignal) {
    return this.mutate(
      input,
      async (ws, context) => {
        const rows = input.cells.length
        const columns = input.cells[0].length
        const r = ws
          .getRange(input.range)
          .getCell(0, 0)
          .getResizedRange(rows - 1, columns - 1)
        if (!input.allow_overwrite) {
          r.load('values')
          await sync(context, signal)
          if (
            r.values.some((row: unknown[]) => row.some((value) => value !== null && value !== ''))
          )
            throw new Error('office_write_failed')
        }
        cancelled(signal)
        r.values = input.cells.map((row: any[]) => row.map((c) => c.value ?? null))
        cancelled(signal)
        r.formulas = input.cells.map((row: any[]) => row.map((c) => c.formula ?? null))
        for (let row = 0; row < rows; row++)
          for (let column = 0; column < columns; column++) {
            const source = input.cells[row][column]
            const target = r.getCell(row, column)
            if (source.cellStyles) {
              const style = source.cellStyles
              const assignments: Array<[RuntimeRecord, string, unknown]> = [
                [
                  target.format.font,
                  'bold',
                  style.fontWeight === undefined ? undefined : style.fontWeight === 'bold',
                ],
                [
                  target.format.font,
                  'italic',
                  style.fontStyle === undefined ? undefined : style.fontStyle === 'italic',
                ],
                [target.format.font, 'size', style.fontSize],
                [target.format.font, 'name', style.fontFamily],
                [target.format.font, 'color', style.fontColor],
                [target.format.fill, 'color', style.backgroundColor],
                [target.format, 'horizontalAlignment', style.horizontalAlignment],
                [
                  target,
                  'numberFormat',
                  style.numberFormat === undefined ? undefined : [[style.numberFormat]],
                ],
              ]
              for (const [object, property, value] of assignments)
                if (value !== undefined) {
                  cancelled(signal)
                  object[property] = value
                }
            }
          }
        if (input.copyToRange) {
          cancelled(signal)
          ws.getRange(input.copyToRange).copyFrom(r, 'All')
        }
        const format = (input.copyToRange ? ws.getRange(input.copyToRange) : r).format
        if (input.resizeWidth) {
          cancelled(signal)
          if (input.resizeWidth.type === 'standard') format.autofitColumns()
          else format.columnWidth = input.resizeWidth.value
        }
        if (input.resizeHeight) {
          cancelled(signal)
          if (input.resizeHeight.type === 'standard') format.autofitRows()
          else format.rowHeight = input.resizeHeight.value
        }
      },
      signal,
    )
  }
  clearCellRange(input: Record<string, any>, signal?: AbortSignal) {
    return this.mutate(
      input,
      (ws) => {
        cancelled(signal)
        ws.getRange(input.range).clear(input.clearType ?? 'Contents')
      },
      signal,
    )
  }
  copyTo(input: Record<string, any>, signal?: AbortSignal) {
    return this.mutate(
      input,
      (ws) => {
        cancelled(signal)
        ws.getRange(input.destinationRange).copyFrom(ws.getRange(input.sourceRange), 'All')
      },
      signal,
    )
  }
  modifySheetStructure(input: Record<string, any>, signal?: AbortSignal) {
    return this.mutate(
      input,
      (ws) => {
        const reference = input.reference
        const count = input.count ?? 1
        const numeric = Number(reference)
        const col = (letters: string) =>
          letters
            .toUpperCase()
            .split('')
            .reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0)
        const letters = (number: number) => {
          let result = ''
          for (let n = number; n > 0; n = Math.floor((n - 1) / 26))
            result = String.fromCharCode(((n - 1) % 26) + 65) + result
          return result
        }
        const start =
          input.position === 'after'
            ? input.dimension === 'rows'
              ? numeric + 1
              : letters(col(reference) + 1)
            : reference
        const end =
          input.dimension === 'rows' ? Number(start) + count - 1 : letters(col(start) + count - 1)
        const ref = `${start}:${end}`
        const r = ws.getRange(ref)
        cancelled(signal)
        if (input.operation === 'insert') r.insert(input.dimension === 'rows' ? 'Down' : 'Right')
        else if (input.operation === 'delete') r.delete(input.dimension === 'rows' ? 'Up' : 'Left')
        else if (input.operation === 'hide')
          r[input.dimension === 'rows' ? 'rowHidden' : 'columnHidden'] = true
        else if (input.operation === 'unhide')
          r[input.dimension === 'rows' ? 'rowHidden' : 'columnHidden'] = false
        else if (input.operation === 'freeze') {
          if (input.dimension === 'rows') ws.freezePanes.freezeRows(count)
          else ws.freezePanes.freezeColumns(count)
        } else if (input.operation === 'unfreeze') ws.freezePanes.unfreeze()
        else throw new Error('office_api_unsupported')
      },
      signal,
    )
  }
  async modifyWorkbookStructure(input: Record<string, any>, signal?: AbortSignal) {
    cancelled(signal)
    await this.run(async (context) => {
      const sheets = context.workbook.worksheets
      cancelled(signal)
      if (input.operation === 'create') sheets.add(input.sheetName)
      else {
        const ws = sheet(context, input.sheetId)
        cancelled(signal)
        if (input.operation === 'delete') ws.delete()
        else if (input.operation === 'rename') ws.name = input.newName
        else if (input.operation === 'duplicate') ws.copy('After', ws).name = input.newName
      }
      await sync(context, signal)
    })
  }
  resizeRange(input: Record<string, any>, signal?: AbortSignal) {
    return this.mutate(
      input,
      (ws) => {
        const format = ws.getRange(input.range ?? 'A:XFD').format
        if (input.width) {
          cancelled(signal)
          if (input.width.type === 'standard') format.autofitColumns()
          else format.columnWidth = input.width.value
        }
        if (input.height) {
          cancelled(signal)
          if (input.height.type === 'standard') format.autofitRows()
          else format.rowHeight = input.height.value
        }
      },
      signal,
    )
  }
  modifyObject(input: Record<string, any>, signal?: AbortSignal) {
    return this.mutate(
      input,
      (ws) => {
        const collection = input.objectType === 'chart' ? ws.charts : ws.pivotTables
        cancelled(signal)
        if (input.operation === 'delete') collection.getItem(input.id).delete()
        else if (input.operation === 'create' && input.objectType === 'chart') {
          const chart = collection.add(input.properties.chartType, input.properties.source)
          if (input.properties.anchor) {
            cancelled(signal)
            chart.setPosition(input.properties.anchor)
          }
          if (input.properties.name) {
            cancelled(signal)
            chart.name = input.properties.name
          }
          if (input.properties.title) {
            cancelled(signal)
            chart.title.text = input.properties.title
            chart.title.visible = true
          }
        } else if (input.operation === 'create' && input.objectType === 'pivotTable') {
          cancelled(signal)
          collection.add(input.properties.name, input.properties.source, input.properties.range)
        } else if (input.operation === 'update' && input.objectType === 'chart') {
          const chart = collection.getItem(input.id)
          if (input.properties.name) {
            cancelled(signal)
            chart.name = input.properties.name
          }
          if (input.properties.chartType) {
            cancelled(signal)
            chart.chartType = input.properties.chartType
          }
          if (input.properties.anchor) {
            cancelled(signal)
            chart.setPosition(input.properties.anchor)
          }
          if (input.properties.title) {
            cancelled(signal)
            chart.title.text = input.properties.title
            chart.title.visible = true
          }
        } else throw new Error('office_api_unsupported')
      },
      signal,
    )
  }
  async verifyRanges(targets: string[], signal?: AbortSignal): Promise<unknown> {
    const grouped = targets.map((target) => {
      const match = /^sheet:(\d+)!(.+)$/.exec(target)
      if (!match) throw new Error('office_verify_failed')
      return { sheetId: Number(match[1]), range: match[2] }
    })
    return Promise.all(
      grouped.map((item) =>
        this.getCellRanges(
          { sheetId: item.sheetId, ranges: [item.range], cellLimit: MAX_EXCEL_CELLS },
          signal,
        ),
      ),
    )
  }
  verifyObjects(input: ObjectRequest, signal?: AbortSignal) {
    return this.getAllObjects(input, signal)
  }
  async verifyWorkbook(signal?: AbortSignal): Promise<unknown> {
    cancelled(signal)
    return this.run(async (context) => {
      const sheets = context.workbook.worksheets
      sheets.load({ $top: MAX_EXCEL_OBJECTS + 1 })
      await sync(context, signal)
      for (const ws of (sheets.items ?? []).slice(0, MAX_EXCEL_OBJECTS))
        ws.load('id,name,visibility')
      await sync(context, signal)
      return {
        sheets: (sheets.items ?? []).slice(0, MAX_EXCEL_OBJECTS).map((ws: RuntimeRecord) => ({
          id: safe(ws.id, 256),
          name: safe(ws.name, 256),
          visibility: safe(ws.visibility, 32),
        })),
        hasMore: (sheets.items ?? []).length > MAX_EXCEL_OBJECTS,
      }
    })
  }
}
