import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import { exactObject, integerField, optionalField, stringField } from '../../agent/tool-schema.js'
import { exportSafeCsv, readBoundedCsv, readBoundedImage } from '../shared/import-media.js'
import type { InMemoryVfs } from '../shared/vfs.js'

export interface ExcelImportMediaAdapter {
  captureRange(sheetId: number, range: string, signal?: AbortSignal): Promise<unknown>
  restoreRange(
    sheetId: number,
    range: string,
    snapshot: unknown,
    signal?: AbortSignal,
  ): Promise<void>
  verifyRangeSnapshot(
    sheetId: number,
    range: string,
    snapshot: unknown,
    signal?: AbortSignal,
  ): Promise<boolean>
  fingerprintRange(sheetId: number, range: string, signal?: AbortSignal): Promise<string>
  readRangeValues(sheetId: number, range: string, signal?: AbortSignal): Promise<unknown[][]>
  writeRangeValues(
    sheetId: number,
    startCell: string,
    values: string[][],
    signal?: AbortSignal,
  ): Promise<void>
  verifyRangeValues(
    sheetId: number,
    startCell: string,
    values: string[][],
    signal?: AbortSignal,
  ): Promise<boolean>
  insertImage(
    input: ExcelImageInsertion,
    base64: string,
    signal?: AbortSignal,
  ): Promise<{ id: string }>
  verifyImage(input: ExcelImageInsertion, id: string, signal?: AbortSignal): Promise<boolean>
  removeImage(input: ExcelImageInsertion, id: string, signal?: AbortSignal): Promise<void>
  verifyImageAbsent(input: ExcelImageInsertion, id: string, signal?: AbortSignal): Promise<boolean>
}

export interface ExcelImageInsertion {
  sheetId: number
  cell: string
  width: number
  height: number
}

const path = stringField({ minLength: 1, maxLength: 512 })
const range = (value: unknown) => {
  const parsed = stringField({ minLength: 1, maxLength: 64 })(value)
  if (!/^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/i.test(parsed))
    throw new Error('invalid_tool_input')
  return parsed.toUpperCase().replaceAll('$', '')
}
const point = (value: unknown) => {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < 1 || value > 2_000)
    throw new Error('invalid_tool_input')
  return value
}
const explanation = optionalField(stringField({ maxLength: 100 }))
const csvToSheet = exactObject({
  path,
  sheetId: integerField({ min: 1, max: 10_000 }),
  startCell: range,
  explanation,
})
const sheetToCsv = exactObject({
  sheetId: integerField({ min: 1, max: 10_000 }),
  range,
  path,
})
const imageToSheet = exactObject({
  path,
  sheetId: integerField({ min: 1, max: 10_000 }),
  cell: range,
  width: point,
  height: point,
  explanation,
})

const tools = [
  {
    name: 'csv-to-sheet',
    description: 'Propose importing a bounded VFS CSV into an Excel range.',
    inputSchema: schema(
      {
        path: stringSchema(512),
        sheetId: integerSchema(),
        startCell: stringSchema(64),
        explanation: stringSchema(100),
      },
      ['path', 'sheetId', 'startCell'],
    ),
  },
  {
    name: 'sheet-to-csv',
    description: 'Export a bounded Excel range to formula-injection-safe CSV in the VFS.',
    inputSchema: schema(
      { sheetId: integerSchema(), range: stringSchema(64), path: stringSchema(512) },
      ['sheetId', 'range', 'path'],
    ),
  },
  {
    name: 'image-to-sheet',
    description: 'Propose inserting a bounded VFS PNG or JPEG into an Excel sheet.',
    inputSchema: schema(
      {
        path: stringSchema(512),
        sheetId: integerSchema(),
        cell: stringSchema(64),
        width: numberSchema(),
        height: numberSchema(),
        explanation: stringSchema(100),
      },
      ['path', 'sheetId', 'cell', 'width', 'height'],
    ),
  },
]

function schema(properties: Record<string, unknown>, required: string[]) {
  return { type: 'object', properties, required, additionalProperties: false }
}
function stringSchema(maxLength: number) {
  return { type: 'string', maxLength }
}
function integerSchema() {
  return { type: 'integer', minimum: 1, maximum: 10_000 }
}
function numberSchema() {
  return { type: 'number', minimum: 1, maximum: 2_000 }
}
function cancelled(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('cancelled')
}
function destinationRange(startCell: string, rows: number, columns: number): string {
  const match = /^([A-Z]{1,3})([1-9]\d*)$/.exec(startCell)
  if (!match) throw new Error('invalid_tool_input')
  let column = 0
  for (const character of match[1]) column = column * 26 + character.charCodeAt(0) - 64
  const endColumn = column + columns - 1
  const endRow = Number(match[2]) + rows - 1
  if (endColumn > 16_384 || endRow > 1_048_576) throw new Error('import_limit')
  let letters = ''
  for (let value = endColumn; value > 0; value = Math.floor((value - 1) / 26))
    letters = String.fromCharCode(((value - 1) % 26) + 65) + letters
  return `${startCell}:${letters}${endRow}`
}
function failed(name: string, error: unknown): ToolExecution {
  const code =
    error instanceof Error &&
    [
      'invalid_tool_input',
      'import_limit',
      'export_limit',
      'invalid_csv',
      'image_limit',
      'image_mime_unsupported',
      'invalid_image',
      'vfs_not_found',
      'vfs_limit',
      'office_api_unsupported',
      'cancelled',
    ].includes(error.message)
      ? error.message
      : 'office_operation_failed'
  return { output: code, isError: true, mutated: false, summary: name }
}

export function createExcelImportMediaSkill(options: {
  adapter: ExcelImportMediaAdapter
  proposals: StructuredProposalController
  vfs: InMemoryVfs
}): AgentSkill {
  return {
    id: 'office-excel-import-media',
    systemPrompt:
      'CSV and image imports use bounded VFS data and require confirmation. CSV exports neutralize spreadsheet formulas.',
    tools,
    async executeTool(call, signal) {
      if (call.inputError || call.truncated)
        return failed(call.name, new Error('invalid_tool_input'))
      try {
        cancelled(signal)
        if (call.name === 'sheet-to-csv') {
          const input = sheetToCsv(call.input)
          const values = await options.adapter.readRangeValues(input.sheetId, input.range, signal)
          cancelled(signal)
          options.vfs.writeFile(input.path, exportSafeCsv(values))
          return {
            output: JSON.stringify({
              path: input.path,
              rows: values.length,
              columns: values[0]?.length ?? 0,
            }),
            mutated: false,
            summary: 'Exported safe CSV',
          }
        }
        if (call.name === 'csv-to-sheet') {
          const input = csvToSheet(call.input)
          const values = readBoundedCsv(options.vfs, input.path)
          const targetRange = destinationRange(input.startCell, values.length, values[0].length)
          const before = await options.adapter.fingerprintRange(input.sheetId, targetRange, signal)
          const snapshot = await options.adapter.captureRange(input.sheetId, targetRange, signal)
          cancelled(signal)
          const recover = async () => {
            try {
              await options.adapter.restoreRange(input.sheetId, targetRange, snapshot)
              if (
                !(await options.adapter.verifyRangeSnapshot(input.sheetId, targetRange, snapshot))
              )
                throw new Error('office_recovery_failed')
            } catch {
              throw new Error('office_recovery_failed')
            }
          }
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: input.explanation || 'Import CSV into sheet',
            preview: {
              path: input.path,
              startCell: input.startCell,
              targetRange,
              rows: values.length,
              columns: values[0].length,
              sample: values
                .slice(0, 5)
                .map((row) => row.slice(0, 8).map((cell) => cell.slice(0, 256))),
            },
            impact: {
              host: 'Excel',
              targets: [`sheet:${input.sheetId}!${targetRange}`],
              count: values.length * values[0].length,
            },
            fingerprint: selectionFingerprint(before),
            before,
            validate: async (s) =>
              (await options.adapter.fingerprintRange(input.sheetId, targetRange, s)) === before,
            execute: async (s) => {
              cancelled(s)
              try {
                await options.adapter.writeRangeValues(input.sheetId, input.startCell, values, s)
              } catch {
                await recover()
                if (s?.aborted) throw new Error('cancelled')
                throw new Error('office_write_failed')
              }
              if (s?.aborted) {
                await recover()
                throw new Error('cancelled')
              }
            },
            verify: async (s) => {
              if (
                !(await options.adapter.verifyRangeValues(
                  input.sheetId,
                  input.startCell,
                  values,
                  s,
                ))
              ) {
                await recover()
                throw new Error('office_verify_failed')
              }
            },
          })
          return {
            output: JSON.stringify({ proposalId: proposal.id, mutated: false }),
            mutated: false,
            summary: 'Proposed CSV import',
          }
        }
        if (call.name === 'image-to-sheet') {
          const input = imageToSheet(call.input)
          const image = readBoundedImage(options.vfs, input.path)
          const insertion = {
            sheetId: input.sheetId,
            cell: input.cell,
            width: input.width,
            height: input.height,
          }
          const before = await options.adapter.fingerprintRange(input.sheetId, input.cell, signal)
          let id: string | undefined
          const recover = async () => {
            if (!id) throw new Error('office_recovery_failed')
            try {
              await options.adapter.removeImage(insertion, id)
              if (!(await options.adapter.verifyImageAbsent(insertion, id)))
                throw new Error('office_recovery_failed')
            } catch {
              throw new Error('office_recovery_failed')
            }
          }
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: input.explanation || 'Insert image into sheet',
            preview: {
              path: input.path,
              mime: image.mime,
              bytes: image.bytes,
              sourceWidth: image.width,
              sourceHeight: image.height,
              ...insertion,
            },
            impact: { host: 'Excel', targets: [`sheet:${input.sheetId}!${input.cell}`], count: 1 },
            fingerprint: selectionFingerprint(`${before}:${image.fingerprint}`),
            before,
            validate: async (s) =>
              (await options.adapter.fingerprintRange(input.sheetId, input.cell, s)) === before,
            execute: async (s) => {
              cancelled(s)
              try {
                id = (await options.adapter.insertImage(insertion, image.base64, s)).id
                if (s?.aborted) {
                  await recover()
                  throw new Error('cancelled')
                }
              } catch (error) {
                if (error instanceof Error && error.message === 'office_recovery_failed')
                  throw error
                if (s?.aborted) throw new Error('cancelled')
                throw new Error('office_write_failed')
              }
            },
            verify: async (s) => {
              if (!id || !(await options.adapter.verifyImage(insertion, id, s))) {
                await recover()
                throw new Error('office_verify_failed')
              }
            },
          })
          return {
            output: JSON.stringify({ proposalId: proposal.id, mutated: false }),
            mutated: false,
            summary: 'Proposed Excel image insertion',
          }
        }
        throw new Error('invalid_tool_input')
      } catch (error) {
        return failed(call.name, error)
      }
    },
  }
}
