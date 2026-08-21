export const MAX_EXCEL_CELLS = 2_000
export const MAX_EXCEL_ROWS = 500
export const MAX_EXCEL_OBJECTS = 256
const MAX_SCREENSHOT_CELLS = 10_000
const MAX_SCREENSHOT_DIMENSION = 8_192
const MAX_SCREENSHOT_AREA = 16_777_216

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
  captureMutation(tool: string, input: Record<string, any>, signal?: AbortSignal): Promise<unknown>
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
  verifyMutation(
    tool: string,
    input: Record<string, any>,
    beforeState: unknown,
    signal?: AbortSignal,
  ): Promise<boolean>
}

function cancelled(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('cancelled')
}
function runtime(version = '1.3'): RuntimeRecord {
  const root = globalThis as unknown as RuntimeRecord
  const office = root.Office
  const excel = root.Excel
  const requirements = office?.context?.requirements
  if (
    !office ||
    !excel ||
    office.context?.host !== 'Excel' ||
    typeof requirements?.isSetSupported !== 'function' ||
    !requirements.isSetSupported('ExcelApi', version) ||
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
  return context.workbook.worksheets.getItemAt(id - 1)
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
function cellValue(value: unknown): unknown {
  return typeof value === 'string' ? safe(value) : (value ?? null)
}
type Box = { row: number; column: number; rows: number; columns: number }
function columnNumber(value: string): number {
  return value.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1
}
function parseA1(value: string): Box {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i.exec(value)
  if (!match) throw new Error('invalid_tool_input')
  const row = Number(match[2]) - 1,
    column = columnNumber(match[1].toUpperCase())
  const endRow = Number(match[4] ?? match[2]) - 1,
    endColumn = columnNumber((match[3] ?? match[1]).toUpperCase())
  if (endRow < row || endColumn < column || endRow >= 1_048_576 || endColumn >= 16_384)
    throw new Error('invalid_tool_input')
  return { row, column, rows: endRow - row + 1, columns: endColumn - column + 1 }
}
function cellAddress(row: number, column: number): string {
  let letters = ''
  for (let n = column + 1; n > 0; n = Math.floor((n - 1) / 26))
    letters = String.fromCharCode(((n - 1) % 26) + 65) + letters
  return `${letters}${row + 1}`
}

export class BrowserExcelAdapter implements ExcelAdapter {
  private run<T>(callback: (context: RuntimeRecord) => Promise<T>, version = '1.3'): Promise<T> {
    return runtime(version).run(callback)
  }
  async getCellRanges(input: RangeRequest, signal?: AbortSignal): Promise<unknown> {
    cancelled(signal)
    const boxes = input.ranges.map(parseA1)
    return this.run(
      async (context) => {
        const ws = sheet(context, input.sheetId)
        ws.load('name')
        const limit = Math.min(input.cellLimit ?? MAX_EXCEL_CELLS, MAX_EXCEL_CELLS)
        let budget = limit
        const ranges = boxes.flatMap((box) => {
          if (budget <= 0) return []
          const columns = Math.min(box.columns, Math.max(1, budget))
          const rows = Math.min(box.rows, Math.max(1, Math.floor(budget / columns)))
          budget = Math.max(0, budget - rows * columns)
          const r = ws.getRangeByIndexes(box.row, box.column, rows, columns)
          r.load('values,formulas,numberFormat,address,rowCount,columnCount')
          return [{ range: r, box }]
        })
        await sync(context, signal)
        const styleCells: Array<{ cell: RuntimeRecord; output: RuntimeRecord }> = []
        const output = ranges.map(({ range, box }) => {
          const cells: unknown[] = []
          for (let row = 0; row < range.rowCount; row++)
            for (let column = 0; column < range.columnCount; column++) {
              const result: RuntimeRecord = {
                address: cellAddress(box.row + row, box.column + column),
                value: cellValue(range.values?.[row]?.[column]),
                formula: cellValue(range.formulas?.[row]?.[column]),
                numberFormat: safe(range.numberFormat?.[row]?.[column], 256),
              }
              cells.push(result)
              if (input.includeStyles !== false) {
                const cell = range.getCell(row, column)
                cell.load('style')
                cell.format.load('horizontalAlignment')
                cell.format.font.load('name,size,color,bold,italic,underline,strikethrough')
                cell.format.fill.load('color,pattern')
                cell.format.borders.load('items/sideIndex,items/style,items/weight,items/color')
                styleCells.push({ cell, output: result })
              }
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
        if (styleCells.length) {
          await sync(context, signal)
          for (const { cell, output: result } of styleCells)
            result.style = {
              styleName: safe(cell.style, 256),
              fontFamily: safe(cell.format.font.name, 256),
              fontSize: cell.format.font.size ?? null,
              fontColor: safe(cell.format.font.color, 64),
              bold: cell.format.font.bold ?? null,
              italic: cell.format.font.italic ?? null,
              underline: safe(cell.format.font.underline, 32),
              strikethrough: cell.format.font.strikethrough ?? null,
              backgroundColor: safe(cell.format.fill.color, 64),
              fillPattern: safe(cell.format.fill.pattern, 32),
              horizontalAlignment: safe(cell.format.horizontalAlignment, 32),
              borders: (cell.format.borders.items ?? []).map((border: RuntimeRecord) => ({
                side: safe(border.sideIndex, 32),
                style: safe(border.style, 32),
                weight: safe(border.weight, 32),
                color: safe(border.color, 64),
              })),
            }
        }
        const requested = boxes.reduce((sum, box) => sum + box.rows * box.columns, 0)
        return { ranges: output, hasMore: requested > limit }
      },
      input.includeStyles === false ? '1.3' : '1.4',
    )
  }
  async getRangeAsCsv(input: CsvRequest, signal?: AbortSignal): Promise<unknown> {
    cancelled(signal)
    const box = parseA1(input.range)
    const skip = input.includeHeaders === false ? 1 : 0
    const maximum = Math.min(input.maxRows ?? MAX_EXCEL_ROWS, MAX_EXCEL_ROWS)
    const loadedColumns = Math.min(box.columns, MAX_EXCEL_CELLS)
    const loadedRows = Math.max(
      0,
      Math.min(maximum, box.rows - skip, Math.floor(MAX_EXCEL_CELLS / loadedColumns)),
    )
    return this.run(async (context) => {
      const ws = sheet(context, input.sheetId)
      ws.load('name')
      const range = ws.getRangeByIndexes(
        box.row + skip,
        box.column,
        Math.max(loadedRows, 1),
        loadedColumns,
      )
      range.load('values,rowCount,columnCount,address')
      await sync(context, signal)
      const rows =
        loadedRows === 0
          ? []
          : range.values.slice(0, loadedRows).map((row: unknown[]) => row.map(csv).join(','))
      return {
        sheetId: input.sheetId,
        sheetName: safe(ws.name, 256),
        address: cleanAddress(range.address),
        csv: rows.join('\n'),
        rowCount: rows.length,
        columnCount: loadedColumns,
        hasMore: box.rows - skip > loadedRows || box.columns > loadedColumns,
      }
    })
  }
  async searchData(input: SearchRequest, signal?: AbortSignal): Promise<unknown> {
    cancelled(signal)
    const maximum = Math.min(input.options?.maxResults ?? 500, 500)
    const offset = input.offset ?? 0
    const explicit = input.range ? parseA1(input.range) : undefined
    const flags = input.options?.matchCase ? '' : 'i'
    let regex: RegExp
    try {
      regex = input.options?.useRegex
        ? new RegExp(input.searchTerm, flags)
        : new RegExp(input.searchTerm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags)
    } catch {
      throw new Error('invalid_tool_input')
    }
    return this.run(async (context) => {
      const ws = sheet(context, input.sheetId ?? 1)
      ws.load('name')
      let box = explicit
      if (!box) {
        const used = ws.getUsedRangeOrNullObject()
        used.load('address,rowCount,columnCount,isNullObject')
        await sync(context, signal)
        if (used.isNullObject)
          return { matches: [], offset, returned: 0, hasMore: false, nextOffset: null }
        box = parseA1(cleanAddress(used.address))
      }
      const total = box.rows * box.columns
      if (offset > total) throw new Error('invalid_tool_input')
      const scan = Math.min(MAX_EXCEL_CELLS, total - offset)
      if (scan <= 0) return { matches: [], offset, returned: 0, hasMore: false, nextOffset: null }
      const startRow = Math.floor(offset / box.columns),
        startColumn = offset % box.columns
      const firstRowCells = Math.min(scan, box.columns - startColumn)
      const ranges: Array<{
        range: RuntimeRecord
        rawOffset: number
        row: number
        column: number
      }> = []
      const first = ws.getRangeByIndexes(
        box.row + startRow,
        box.column + startColumn,
        1,
        firstRowCells,
      )
      first.load('values,formulas')
      ranges.push({
        range: first,
        rawOffset: offset,
        row: box.row + startRow,
        column: box.column + startColumn,
      })
      const remaining = scan - firstRowCells
      if (remaining > 0) {
        const fullRows = Math.floor(remaining / box.columns)
        if (fullRows > 0) {
          const middle = ws.getRangeByIndexes(
            box.row + startRow + 1,
            box.column,
            fullRows,
            box.columns,
          )
          middle.load('values,formulas')
          ranges.push({
            range: middle,
            rawOffset: offset + firstRowCells,
            row: box.row + startRow + 1,
            column: box.column,
          })
        }
        const tail = remaining % box.columns
        if (tail > 0) {
          const tailRange = ws.getRangeByIndexes(
            box.row + startRow + 1 + fullRows,
            box.column,
            1,
            tail,
          )
          tailRange.load('values,formulas')
          ranges.push({
            range: tailRange,
            rawOffset: offset + firstRowCells + fullRows * box.columns,
            row: box.row + startRow + 1 + fullRows,
            column: box.column,
          })
        }
      }
      await sync(context, signal)
      const matches: unknown[] = []
      let consumed = 0
      outer: for (const page of ranges)
        for (let row = 0; row < page.range.values.length; row++)
          for (let column = 0; column < page.range.values[row].length; column++) {
            consumed += 1
            const candidate = input.options?.matchFormulas
              ? page.range.formulas[row][column]
              : page.range.values[row][column]
            const value = String(candidate ?? '')
            const matched = input.options?.matchEntireCell
              ? input.options.matchCase
                ? value === input.searchTerm
                : value.toLowerCase() === input.searchTerm.toLowerCase()
              : regex.test(value)
            if (matched)
              matches.push({
                sheetId: input.sheetId ?? 1,
                sheetName: safe(ws.name, 256),
                address: cellAddress(page.row + row, page.column + column),
                value: page.range.values[row][column] ?? null,
                formula: page.range.formulas[row][column] ?? null,
              })
            if (matches.length >= maximum) break outer
          }
      const next = offset + consumed,
        hasMore = next < total
      return {
        matches,
        offset,
        returned: matches.length,
        hasMore,
        nextOffset: hasMore ? next : null,
      }
    })
  }
  async screenshotRange(
    input: { sheetId: number; range: string },
    signal?: AbortSignal,
  ): Promise<{ base64: string; mime: 'image/png' }> {
    cancelled(signal)
    const box = parseA1(input.range)
    if (box.rows * box.columns > MAX_SCREENSHOT_CELLS) throw new Error('invalid_tool_input')
    return this.run(async (context) => {
      const range = sheet(context, input.sheetId).getRange(input.range)
      range.load('rowCount,columnCount,width,height')
      await sync(context, signal)
      if (
        range.rowCount * range.columnCount > MAX_SCREENSHOT_CELLS ||
        !Number.isFinite(range.width) ||
        !Number.isFinite(range.height) ||
        range.width <= 0 ||
        range.height <= 0 ||
        range.width > MAX_SCREENSHOT_DIMENSION ||
        range.height > MAX_SCREENSHOT_DIMENSION ||
        range.width * range.height > MAX_SCREENSHOT_AREA
      )
        throw new Error('office_read_failed')
      cancelled(signal)
      const image = range.getImage()
      await sync(context, signal)
      const base64 = safe(image.value, 6 * 1024 * 1024).replace(/^data:image\/png;base64,/, '')
      if (!base64) throw new Error('office_read_failed')
      return { base64, mime: 'image/png' }
    }, '1.7')
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
      const selected = (sheets.items ?? ([] as RuntimeRecord[])).slice(0, MAX_EXCEL_OBJECTS)
      for (const ws of selected) {
        ws.load('id,name')
        ws.charts.load({ $top: MAX_EXCEL_OBJECTS + 1 })
        ws.pivotTables.load({ $top: MAX_EXCEL_OBJECTS + 1 })
      }
      await sync(context, signal)
      const pending: Array<{
        ws: RuntimeRecord
        sheetId: number
        object: RuntimeRecord
        type: string
      }> = selected.flatMap((ws: RuntimeRecord, index: number) => [
        ...(ws.charts?.items ?? []).map((object: RuntimeRecord) => ({
          ws,
          sheetId: input.sheetId ?? index + 1,
          object,
          type: 'chart',
        })),
        ...(ws.pivotTables?.items ?? []).map((object: RuntimeRecord) => ({
          ws,
          sheetId: input.sheetId ?? index + 1,
          object,
          type: 'pivotTable',
        })),
      ])
      for (const item of pending) {
        item.object.load(item.type === 'chart' ? 'id,name,chartType' : 'id,name')
        if (item.type === 'chart') item.object.title.load('text,visible')
      }
      await sync(context, signal)
      const objects = pending.map(({ ws, sheetId, object, type }) => ({
        sheetId,
        sheetName: safe(ws.name, 256),
        id: safe(object.name, 256),
        officeId: safe(object.id, 256),
        name: safe(object.name, 256),
        type,
        ...(type === 'chart' && {
          chartType: safe(object.chartType, 64),
          title: object.title?.visible ? safe(object.title.text, 256) : '',
        }),
      }))
      const filtered = input.id ? objects.filter((item: any) => item.id === input.id) : objects
      return {
        objects: filtered.slice(0, MAX_EXCEL_OBJECTS),
        hasMore: filtered.length > MAX_EXCEL_OBJECTS,
      }
    }, '1.8')
  }
  async fingerprint(targets: string[], signal?: AbortSignal): Promise<string> {
    const values = await Promise.all(
      targets.map(async (target) => {
        if (target.startsWith('workbook:')) return this.verifyWorkbook(signal)
        const structure = /^structure:(\d+):([^:]+):(rows|columns):([^:]*):(\d+)$/.exec(target)
        if (structure)
          return this.run(
            async (context) => {
              const ws = sheet(context, Number(structure[1]))
              const used = ws.getUsedRangeOrNullObject()
              used.load('address,isNullObject')
              const reference = structure[4] || '1'
              const range = ws.getRange(`${reference}:${reference}`)
              range.load('rowHidden,columnHidden')
              const frozen = ['freeze', 'unfreeze'].includes(structure[2])
                ? ws.freezePanes.getLocationOrNullObject()
                : undefined
              frozen?.load('address,isNullObject')
              await sync(context, signal)
              return {
                used: used.isNullObject ? null : cleanAddress(used.address),
                rowHidden: range.rowHidden ?? null,
                columnHidden: range.columnHidden ?? null,
                ...(frozen && {
                  frozen: frozen.isNullObject ? null : cleanAddress(frozen.address),
                }),
              }
            },
            ['freeze', 'unfreeze'].includes(structure[2]) ? '1.7' : '1.4',
          )
        const resize = /^resize:(\d+)!(.+)$/.exec(target)
        if (resize)
          return this.run(async (context) => {
            const range = sheet(context, Number(resize[1])).getRange(resize[2])
            range.format.load('columnWidth,rowHeight')
            await sync(context, signal)
            return {
              columnWidth: range.format.columnWidth ?? null,
              rowHeight: range.format.rowHeight ?? null,
            }
          })
        const object = /^sheet:(\d+)!object:(.+)$/.exec(target)
        if (object)
          return this.verifyObjects(
            { sheetId: Number(object[1]), ...(object[2] === 'new' ? {} : { id: object[2] }) },
            signal,
          )
        return this.verifyRanges([target], signal)
      }),
    )
    const json = JSON.stringify(values)
    let hash = 0x811c9dc5
    for (let index = 0; index < json.length; index++) {
      hash ^= json.charCodeAt(index)
      hash = Math.imul(hash, 0x01000193)
    }
    return `${json.length}:${(hash >>> 0).toString(16)}`
  }
  async captureMutation(
    tool: string,
    input: Record<string, any>,
    signal?: AbortSignal,
  ): Promise<unknown> {
    if (tool === 'modify_workbook_structure' && input.operation === 'delete') {
      const state = (await this.verifyWorkbook(signal)) as RuntimeRecord
      const item = state.sheets.find((candidate: RuntimeRecord) => candidate.id === input.sheetId)
      if (!item) throw new Error('office_verify_failed')
      return { officeId: item.officeId, name: item.name }
    }
    const mutationTargets =
      tool === 'copy_to'
        ? [
            `sheet:${input.sheetId}!${input.sourceRange}`,
            `sheet:${input.sheetId}!${input.destinationRange}`,
          ]
        : tool === 'set_cell_range'
          ? [
              `sheet:${input.sheetId}!${input.range}`,
              ...(input.copyToRange ? [`sheet:${input.sheetId}!${input.copyToRange}`] : []),
            ]
          : tool === 'modify_sheet_structure'
            ? [
                `structure:${input.sheetId}:${input.operation}:${input.dimension}:${input.reference ?? ''}:${input.count ?? 1}`,
              ]
            : tool === 'resize_range'
              ? [`resize:${input.sheetId}!${input.range ?? 'A1:XFD1048576'}`]
              : tool === 'modify_object'
                ? [`sheet:${input.sheetId}!object:${input.id ?? input.properties?.name ?? 'new'}`]
                : [`workbook:${input.operation}:${input.sheetId ?? input.sheetName ?? ''}`]
    return this.fingerprint(mutationTargets, signal)
  }
  private async mutate(
    input: Record<string, any>,
    action: (ws: RuntimeRecord, context: RuntimeRecord) => void | Promise<void>,
    signal?: AbortSignal,
    version = '1.3',
  ) {
    cancelled(signal)
    await this.run(async (context) => {
      const ws = sheet(context, input.sheetId)
      await action(ws, context)
      await sync(context, signal)
    }, version)
  }
  setCellRange(input: Record<string, any>, signal?: AbortSignal) {
    parseA1(input.range)
    if (input.copyToRange) parseA1(input.copyToRange)
    return this.mutate(
      input,
      async (ws, context) => {
        const rows = input.cells.length
        const columns = input.cells[0].length
        const box = parseA1(input.range)
        const r = ws
          .getRange(input.range)
          .getCell(0, 0)
          .getResizedRange(rows - 1, columns - 1)
        const notes = new Map<string, RuntimeRecord>()
        for (let row = 0; row < rows; row++)
          for (let column = 0; column < columns; column++)
            if (input.cells[row][column].note !== undefined) {
              const address = cellAddress(box.row + row, box.column + column)
              const note = ws.notes.getItemOrNullObject(address)
              note.load('isNullObject')
              notes.set(address, note)
            }
        if (!input.allow_overwrite) {
          r.load('values')
        }
        if (!input.allow_overwrite || notes.size) await sync(context, signal)
        if (!input.allow_overwrite) {
          if (
            r.values.some((row: unknown[]) => row.some((value) => value !== null && value !== ''))
          )
            throw new Error('office_write_failed')
        }
        for (let row = 0; row < rows; row++)
          for (let column = 0; column < columns; column++) {
            const source = input.cells[row][column]
            const target = r.getCell(row, column)
            if (Object.hasOwn(source, 'value')) {
              cancelled(signal)
              target.values = [[source.value]]
            }
            if (Object.hasOwn(source, 'formula')) {
              cancelled(signal)
              target.formulas = [[source.formula]]
            }
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
                [
                  target.format.font,
                  'underline',
                  style.fontLine === undefined
                    ? undefined
                    : style.fontLine === 'underline'
                      ? 'Single'
                      : 'None',
                ],
                [
                  target.format.font,
                  'strikethrough',
                  style.fontLine === undefined ? undefined : style.fontLine === 'line-through',
                ],
                [target.format.font, 'size', style.fontSize],
                [target.format.font, 'name', style.fontFamily],
                [target.format.font, 'color', style.fontColor],
                [target.format.fill, 'color', style.backgroundColor],
                [
                  target.format,
                  'horizontalAlignment',
                  style.horizontalAlignment === undefined
                    ? undefined
                    : ({ left: 'Left', center: 'Center', right: 'Right' } as RuntimeRecord)[
                        style.horizontalAlignment
                      ],
                ],
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
            if (source.borderStyles)
              for (const [side, config] of Object.entries(source.borderStyles) as Array<
                [string, RuntimeRecord]
              >) {
                const border = target.format.borders.getItem(
                  (
                    {
                      top: 'EdgeTop',
                      bottom: 'EdgeBottom',
                      left: 'EdgeLeft',
                      right: 'EdgeRight',
                    } as RuntimeRecord
                  )[side],
                )
                if (config.style) {
                  cancelled(signal)
                  border.style = (
                    {
                      solid: 'Continuous',
                      dashed: 'Dash',
                      dotted: 'Dot',
                      double: 'Double',
                    } as RuntimeRecord
                  )[config.style]
                }
                if (config.weight) {
                  cancelled(signal)
                  border.weight = (
                    { thin: 'Thin', medium: 'Medium', thick: 'Thick' } as RuntimeRecord
                  )[config.weight]
                }
                if (config.color) {
                  cancelled(signal)
                  border.color = config.color
                }
              }
            if (source.note) {
              cancelled(signal)
              const address = cellAddress(box.row + row, box.column + column)
              const existing = notes.get(address)
              if (existing && !existing.isNullObject) existing.content = source.note
              else ws.notes.add(address, source.note)
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
      input.cells.some((row: RuntimeRecord[]) => row.some((cell) => cell.note !== undefined))
        ? '1.18'
        : '1.4',
    )
  }
  clearCellRange(input: Record<string, any>, signal?: AbortSignal) {
    parseA1(input.range)
    return this.mutate(
      input,
      (ws) => {
        cancelled(signal)
        const types: RuntimeRecord = { contents: 'Contents', formats: 'Formats', all: 'All' }
        ws.getRange(input.range).clear(types[input.clearType ?? 'contents'])
      },
      signal,
      '1.4',
    )
  }
  copyTo(input: Record<string, any>, signal?: AbortSignal) {
    parseA1(input.sourceRange)
    parseA1(input.destinationRange)
    return this.mutate(
      input,
      (ws) => {
        cancelled(signal)
        ws.getRange(input.destinationRange).copyFrom(ws.getRange(input.sourceRange), 'All')
      },
      signal,
      '1.4',
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
      ['freeze', 'unfreeze'].includes(input.operation) ? '1.7' : '1.4',
    )
  }
  async modifyWorkbookStructure(input: Record<string, any>, signal?: AbortSignal) {
    cancelled(signal)
    await this.run(
      async (context) => {
        const sheets = context.workbook.worksheets
        cancelled(signal)
        if (input.operation === 'create') {
          const created = sheets.add(input.sheetName)
          if (input.tabColor) {
            cancelled(signal)
            created.tabColor = input.tabColor
          }
        } else {
          const ws = sheet(context, input.sheetId)
          cancelled(signal)
          if (input.operation === 'delete') ws.delete()
          else if (input.operation === 'rename') ws.name = input.newName
          else if (input.operation === 'duplicate') ws.copy('After', ws).name = input.newName
        }
        await sync(context, signal)
      },
      input.operation === 'duplicate' ? '1.7' : '1.4',
    )
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
      '1.4',
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
          const types: RuntimeRecord = {
            columnClustered: 'ColumnClustered',
            barClustered: 'BarClustered',
            line: 'Line',
            pie: 'Pie',
            scatter: 'XYScatter',
            area: 'Area',
            doughnut: 'Doughnut',
          }
          const chart = collection.add(
            types[input.properties.chartType],
            ws.getRange(input.properties.source),
          )
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
            chart.chartType = (
              {
                columnClustered: 'ColumnClustered',
                barClustered: 'BarClustered',
                line: 'Line',
                pie: 'Pie',
                scatter: 'XYScatter',
                area: 'Area',
                doughnut: 'Doughnut',
              } as RuntimeRecord
            )[input.properties.chartType]
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
        } else if (input.operation === 'update' && input.objectType === 'pivotTable') {
          const pivot = collection.getItem(input.id)
          if (input.properties.source && input.properties.range) {
            cancelled(signal)
            pivot.delete()
            const replacement = collection.add(
              input.properties.name ?? input.id,
              input.properties.source,
              input.properties.range,
            )
            cancelled(signal)
            replacement.refresh()
          } else {
            if (input.properties.name) {
              cancelled(signal)
              pivot.name = input.properties.name
            }
            cancelled(signal)
            pivot.refresh()
          }
        } else throw new Error('office_api_unsupported')
      },
      signal,
      '1.8',
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
        ws.load('id,name,visibility,tabColor')
      await sync(context, signal)
      return {
        sheets: (sheets.items ?? [])
          .slice(0, MAX_EXCEL_OBJECTS)
          .map((ws: RuntimeRecord, index: number) => ({
            id: index + 1,
            officeId: safe(ws.id, 256),
            name: safe(ws.name, 256),
            visibility: safe(ws.visibility, 32),
            tabColor: safe(ws.tabColor, 64),
          })),
        hasMore: (sheets.items ?? []).length > MAX_EXCEL_OBJECTS,
      }
    })
  }
  async verifyMutation(
    tool: string,
    input: Record<string, any>,
    beforeState: unknown,
    signal?: AbortSignal,
  ): Promise<boolean> {
    if (tool === 'set_cell_range') {
      const result = (await this.getCellRanges(
        {
          sheetId: input.sheetId,
          ranges: [input.range, ...(input.copyToRange ? [input.copyToRange] : [])],
          includeStyles: true,
          cellLimit: input.cells.length * input.cells[0].length * (input.copyToRange ? 2 : 1),
        },
        signal,
      )) as RuntimeRecord
      const styleMatches = (expected: RuntimeRecord, actual: RuntimeRecord) => {
        const style = expected.cellStyles
        if (style) {
          const normalized: RuntimeRecord = {
            fontWeight: actual.style?.bold ? 'bold' : 'normal',
            fontStyle: actual.style?.italic ? 'italic' : 'normal',
            fontLine: actual.style?.strikethrough
              ? 'line-through'
              : actual.style?.underline && actual.style.underline !== 'None'
                ? 'underline'
                : 'none',
            fontSize: actual.style?.fontSize,
            fontFamily: actual.style?.fontFamily,
            fontColor: actual.style?.fontColor,
            backgroundColor: actual.style?.backgroundColor,
            horizontalAlignment: String(actual.style?.horizontalAlignment ?? '').toLowerCase(),
            numberFormat: actual.numberFormat,
          }
          if (Object.entries(style).some(([key, value]) => normalized[key] !== value)) return false
        }
        const borderMap: RuntimeRecord = {
          EdgeTop: 'top',
          EdgeBottom: 'bottom',
          EdgeLeft: 'left',
          EdgeRight: 'right',
        }
        const styles: RuntimeRecord = {
          Continuous: 'solid',
          Dash: 'dashed',
          Dot: 'dotted',
          Double: 'double',
        }
        const weights: RuntimeRecord = { Thin: 'thin', Medium: 'medium', Thick: 'thick' }
        for (const [side, config] of Object.entries(expected.borderStyles ?? {}) as Array<
          [string, RuntimeRecord]
        >) {
          const actualBorder = actual.style?.borders?.find(
            (border: RuntimeRecord) => borderMap[border.side] === side,
          )
          if (!actualBorder) return false
          if (config.style && styles[actualBorder.style] !== config.style) return false
          if (config.weight && weights[actualBorder.weight] !== config.weight) return false
          if (config.color && actualBorder.color !== config.color) return false
        }
        return true
      }
      const valuesMatch = result.ranges.every((rangeResult: RuntimeRecord) =>
        input.cells
          .flat()
          .every(
            (expected: RuntimeRecord, index: number) =>
              (!Object.hasOwn(expected, 'value') ||
                rangeResult.cells[index]?.value === expected.value) &&
              (!Object.hasOwn(expected, 'formula') ||
                rangeResult.cells[index]?.formula === expected.formula) &&
              styleMatches(expected, rangeResult.cells[index] ?? {}),
          ),
      )
      if (!valuesMatch) return false
      const notes: Array<{ index: number; content: string }> = input.cells
        .flat()
        .flatMap((row: RuntimeRecord, index: number) =>
          row.note === undefined ? [] : [{ index, content: row.note }],
        )
      if (!notes.length) return true
      return this.run(async (context) => {
        const ws = sheet(context, input.sheetId)
        const loaded: Array<{ note: RuntimeRecord; content: string }> = [
          input.range,
          ...(input.copyToRange ? [input.copyToRange] : []),
        ].flatMap((address) => {
          const box = parseA1(address)
          return notes.map(({ index, content }) => {
            const note = ws.notes.getItemOrNullObject(
              cellAddress(
                box.row + Math.floor(index / input.cells[0].length),
                box.column + (index % input.cells[0].length),
              ),
            )
            note.load('content,isNullObject')
            return { note, content }
          })
        })
        await sync(context, signal)
        return loaded.every(({ note, content }) => !note.isNullObject && note.content === content)
      }, '1.18')
    }
    if (tool === 'clear_cell_range') {
      const result = (await this.getCellRanges(
        {
          sheetId: input.sheetId,
          ranges: [input.range],
          includeStyles: input.clearType !== 'contents',
          cellLimit: MAX_EXCEL_CELLS,
        },
        signal,
      )) as RuntimeRecord
      return result.ranges[0].cells.every((cell: RuntimeRecord) => {
        const contentsClear =
          (cell.value === null || cell.value === '') &&
          (cell.formula === null || cell.formula === '')
        const formatsClear =
          cell.style?.styleName === 'Normal' &&
          cell.numberFormat === 'General' &&
          !cell.style?.bold &&
          !cell.style?.italic &&
          (cell.style?.underline === 'None' || cell.style?.underline === '') &&
          !cell.style?.strikethrough &&
          typeof cell.style?.fontFamily === 'string' &&
          cell.style.fontFamily.length > 0 &&
          Number.isFinite(cell.style?.fontSize) &&
          cell.style.fontSize > 0 &&
          typeof cell.style?.fontColor === 'string' &&
          cell.style.fontColor.length > 0 &&
          typeof cell.style?.backgroundColor === 'string' &&
          cell.style.backgroundColor.length > 0 &&
          (cell.style?.fillPattern === 'None' || cell.style?.fillPattern === '') &&
          ['General', ''].includes(cell.style?.horizontalAlignment) &&
          (cell.style?.borders ?? []).every(
            (border: RuntimeRecord) => border.style === 'None' || border.style === '',
          )
        return input.clearType === 'formats'
          ? formatsClear
          : input.clearType === 'all'
            ? contentsClear && formatsClear
            : contentsClear
      })
    }
    if (tool === 'copy_to') {
      const result = (await this.getCellRanges(
        {
          sheetId: input.sheetId,
          ranges: [input.sourceRange, input.destinationRange],
          includeStyles: true,
          cellLimit: MAX_EXCEL_CELLS,
        },
        signal,
      )) as RuntimeRecord
      const comparable = (cell: RuntimeRecord) => ({
        value: cell.value,
        formula: cell.formula,
        numberFormat: cell.numberFormat,
        style: cell.style,
      })
      return (
        JSON.stringify(result.ranges[0]?.cells?.map(comparable)) ===
        JSON.stringify(result.ranges[1]?.cells?.map(comparable))
      )
    }
    if (tool === 'resize_range')
      return this.run(async (context) => {
        const format = sheet(context, input.sheetId).getRange(input.range ?? 'A:XFD').format
        format.load('columnWidth,rowHeight')
        await sync(context, signal)
        return (
          (!input.width ||
            (input.width.type === 'standard'
              ? Number.isFinite(format.columnWidth) && format.columnWidth > 0
              : Math.abs(format.columnWidth - input.width.value) < 0.01)) &&
          (!input.height ||
            (input.height.type === 'standard'
              ? Number.isFinite(format.rowHeight) && format.rowHeight > 0
              : Math.abs(format.rowHeight - input.height.value) < 0.01))
        )
      })
    if (tool === 'modify_sheet_structure')
      if (['insert', 'delete'].includes(input.operation))
        return (
          (await this.fingerprint(
            [
              `structure:${input.sheetId}:${input.operation}:${input.dimension}:${input.reference}:${input.count ?? 1}`,
            ],
            signal,
          )) !== beforeState
        )
      else
        return this.run(async (context) => {
          const ws = sheet(context, input.sheetId)
          if (['hide', 'unhide'].includes(input.operation)) {
            const range = ws.getRange(`${input.reference}:${input.reference}`)
            range.load('rowHidden,columnHidden')
            await sync(context, signal)
            return (
              range[input.dimension === 'rows' ? 'rowHidden' : 'columnHidden'] ===
              (input.operation === 'hide')
            )
          }
          const frozen = ws.freezePanes.getLocationOrNullObject()
          frozen.load('isNullObject')
          await sync(context, signal)
          return input.operation === 'unfreeze' ? frozen.isNullObject : !frozen.isNullObject
        }, '1.7')
    if (tool === 'modify_workbook_structure') {
      const state = (await this.verifyWorkbook(signal)) as RuntimeRecord
      const names = state.sheets.map((item: RuntimeRecord) => item.name)
      if (input.operation === 'create') return names.includes(input.sheetName)
      if (input.operation === 'rename' || input.operation === 'duplicate')
        return names.includes(input.newName)
      if (input.operation === 'delete')
        return (
          !!beforeState &&
          typeof beforeState === 'object' &&
          !state.sheets.some(
            (item: RuntimeRecord) =>
              item.officeId === (beforeState as RuntimeRecord).officeId ||
              (item.officeId === '' && item.name === (beforeState as RuntimeRecord).name),
          )
        )
    }
    if (tool === 'modify_object') {
      const result = (await this.getAllObjects(
        {
          sheetId: input.sheetId,
          id: input.operation === 'delete' ? input.id : (input.properties?.name ?? input.id),
        },
        signal,
      )) as RuntimeRecord
      if (input.operation === 'delete') return result.objects.length === 0
      const expectedName = input.properties?.name ?? input.id
      return result.objects.some(
        (object: RuntimeRecord) =>
          object.type === input.objectType &&
          (!expectedName || object.name === expectedName) &&
          (!input.properties?.chartType ||
            object.chartType ===
              (
                {
                  columnClustered: 'ColumnClustered',
                  barClustered: 'BarClustered',
                  line: 'Line',
                  pie: 'Pie',
                  scatter: 'XYScatter',
                  area: 'Area',
                  doughnut: 'Doughnut',
                } as RuntimeRecord
              )[input.properties.chartType]) &&
          (!input.properties?.title || object.title === input.properties.title),
      )
    }
    return true
  }
}
