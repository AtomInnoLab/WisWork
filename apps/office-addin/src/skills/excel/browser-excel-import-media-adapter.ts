import type { ExcelImageInsertion, ExcelImportMediaAdapter } from './excel-import-media.js'

type Runtime = Record<string, any>
function cancelled(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('cancelled')
}
function hash(value: unknown): string {
  const text = JSON.stringify(value)
  let result = 0x811c9dc5
  for (let index = 0; index < text.length; index += 1)
    result = Math.imul(result ^ text.charCodeAt(index), 0x01000193)
  return `${text.length}:${(result >>> 0).toString(16)}`
}
function runtime(): Runtime {
  const root = globalThis as Runtime
  const requirements = root.Office?.context?.requirements
  if (
    root.Office?.context?.host !== 'Excel' ||
    !requirements?.isSetSupported?.('ExcelApi', '1.9') ||
    typeof root.Excel?.run !== 'function'
  )
    throw new Error('office_api_unsupported')
  return root.Excel
}
function worksheet(context: Runtime, sheetId: number): Runtime {
  return context.workbook.worksheets.getItemAt(sheetId - 1)
}
async function sync(context: Runtime, signal?: AbortSignal) {
  cancelled(signal)
  await context.sync()
  cancelled(signal)
}

export function supportsExcelImportMedia(): boolean {
  try {
    runtime()
    return true
  } catch {
    return false
  }
}

export class BrowserExcelImportMediaAdapter implements ExcelImportMediaAdapter {
  private run<T>(callback: (context: Runtime) => Promise<T>): Promise<T> {
    return runtime().run(callback)
  }
  async fingerprintRange(sheetId: number, range: string, signal?: AbortSignal): Promise<string> {
    return this.run(async (context) => {
      const item = worksheet(context, sheetId).getRange(range)
      if (typeof item.load !== 'function') throw new Error('office_api_unsupported')
      item.load('address,values,formulas,rowCount,columnCount')
      await sync(context, signal)
      return hash({
        address: item.address,
        values: item.values,
        formulas: item.formulas,
        rows: item.rowCount,
        columns: item.columnCount,
      })
    })
  }
  async captureRangeState(
    sheetId: number,
    range: string,
    signal?: AbortSignal,
  ): Promise<{ fingerprint: string; snapshot: unknown }> {
    return this.run(async (context) => {
      const item = worksheet(context, sheetId).getRange(range)
      item.load('address,values,formulas,rowCount,columnCount')
      await sync(context, signal)
      if (item.rowCount * item.columnCount > 10_000 || !Array.isArray(item.formulas))
        throw new Error('office_api_unsupported')
      const snapshot = {
        address: item.address,
        formulas: item.formulas.map((row: unknown[]) => row.slice()),
      }
      return {
        fingerprint: hash({
          address: item.address,
          values: item.values,
          formulas: item.formulas,
          rows: item.rowCount,
          columns: item.columnCount,
        }),
        snapshot,
      }
    })
  }
  async restoreRange(sheetId: number, range: string, snapshot: unknown): Promise<void> {
    await this.run(async (context) => {
      const saved = snapshot as Runtime
      if (!Array.isArray(saved?.formulas)) throw new Error('office_recovery_failed')
      worksheet(context, sheetId).getRange(range).formulas = saved.formulas
      await sync(context)
    })
  }
  async verifyRangeSnapshot(sheetId: number, range: string, snapshot: unknown): Promise<boolean> {
    return this.run(async (context) => {
      const item = worksheet(context, sheetId).getRange(range)
      item.load('address,formulas')
      await sync(context)
      const saved = snapshot as Runtime
      return (
        item.address === saved.address &&
        JSON.stringify(item.formulas) === JSON.stringify(saved.formulas)
      )
    })
  }
  async readRangeValues(
    sheetId: number,
    range: string,
    signal?: AbortSignal,
  ): Promise<unknown[][]> {
    return this.run(async (context) => {
      const item = worksheet(context, sheetId).getRange(range)
      item.load('values,rowCount,columnCount')
      await sync(context, signal)
      if (
        !Array.isArray(item.values) ||
        item.rowCount > 500 ||
        item.columnCount > 100 ||
        item.rowCount * item.columnCount > 10_000
      )
        throw new Error('office_read_failed')
      return item.values.map((row: unknown[]) => row.slice())
    })
  }
  async writeRangeValues(
    sheetId: number,
    startCell: string,
    values: string[][],
    signal?: AbortSignal,
  ): Promise<void> {
    await this.run(async (context) => {
      const start = worksheet(context, sheetId).getRange(startCell)
      const target = start.getResizedRange(values.length - 1, values[0].length - 1)
      cancelled(signal)
      target.values = values
      await sync(context, signal)
    })
  }
  async verifyRangeValues(
    sheetId: number,
    startCell: string,
    values: string[][],
    signal?: AbortSignal,
  ): Promise<boolean> {
    return this.run(async (context) => {
      const target = worksheet(context, sheetId)
        .getRange(startCell)
        .getResizedRange(values.length - 1, values[0].length - 1)
      target.load('values,rowCount,columnCount')
      await sync(context, signal)
      return (
        target.rowCount === values.length &&
        target.columnCount === values[0].length &&
        JSON.stringify(target.values) === JSON.stringify(values)
      )
    })
  }
  async insertImage(
    input: ExcelImageInsertion,
    base64: string,
    signal?: AbortSignal,
  ): Promise<{ id: string }> {
    return this.run(async (context) => {
      const sheet = worksheet(context, input.sheetId)
      const shapes = sheet.shapes
      if (typeof shapes?.addImage !== 'function') throw new Error('office_api_unsupported')
      shapes.load('items/id')
      const anchor = sheet.getRange(input.cell)
      anchor.load('left,top')
      await sync(context, signal)
      const beforeIds = new Set((shapes.items as Runtime[]).map((shape) => String(shape.id)))
      cancelled(signal)
      let created: Runtime | undefined
      try {
        created = shapes.addImage(base64)
        if (typeof created?.delete !== 'function') throw new Error('office_api_unsupported')
        created.left = anchor.left
        created.top = anchor.top
        created.width = input.width
        created.height = input.height
        created.load('id')
        await sync(context, signal)
        if (!created.id) throw new Error('office_write_failed')
        return { id: String(created.id) }
      } catch (writeError) {
        try {
          if (created) {
            created.delete()
            await sync(context)
          }
          shapes.load('items/id')
          await sync(context)
          const recovered = new Set((shapes.items as Runtime[]).map((shape) => String(shape.id)))
          if (recovered.size !== beforeIds.size || [...recovered].some((id) => !beforeIds.has(id)))
            throw new Error('office_recovery_failed')
        } catch (recoveryError) {
          throw new Error('office_recovery_failed', { cause: recoveryError })
        }
        throw new Error('office_write_failed', { cause: writeError })
      }
    })
  }
  async verifyImage(
    input: ExcelImageInsertion,
    id: string,
    signal?: AbortSignal,
  ): Promise<boolean> {
    return this.run(async (context) => {
      const sheet = worksheet(context, input.sheetId)
      const anchor = sheet.getRange(input.cell)
      anchor.load('left,top')
      const shape = sheet.shapes.getItem(id)
      shape.load('id,left,top,width,height,type')
      await sync(context, signal)
      return (
        String(shape.id) === id &&
        String(shape.type).toLowerCase().includes('image') &&
        shape.left === anchor.left &&
        shape.top === anchor.top &&
        shape.width === input.width &&
        shape.height === input.height
      )
    })
  }
  async removeImage(input: ExcelImageInsertion, id: string): Promise<void> {
    await this.run(async (context) => {
      const shape = worksheet(context, input.sheetId).shapes.getItem(id)
      if (typeof shape.delete !== 'function') throw new Error('office_api_unsupported')
      shape.delete()
      await sync(context)
    })
  }
  async verifyImageAbsent(input: ExcelImageInsertion, id: string): Promise<boolean> {
    return this.run(async (context) => {
      const shapes = worksheet(context, input.sheetId).shapes
      shapes.load('items/id')
      await sync(context)
      return !(shapes.items as Runtime[]).some((shape) => String(shape.id) === id)
    })
  }
}
