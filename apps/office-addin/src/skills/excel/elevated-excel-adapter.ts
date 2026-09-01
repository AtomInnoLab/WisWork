import type {
  ElevatedOfficeAdapter,
  ElevatedOfficeAuthority,
  ElevatedOfficeProgram,
  ElevatedOfficeSnapshot,
} from '../shared/elevated-office-program.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import type { ExcelAdapter } from './browser-excel-adapter.js'

export interface ExcelElevatedAuthority {
  captureAuthority(): ElevatedOfficeAuthority
  snapshot(program: ElevatedOfficeProgram, signal?: AbortSignal): Promise<ElevatedOfficeSnapshot>
  validateSnapshot(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<boolean>
  execute(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<void>
  readback(
    program: ElevatedOfficeProgram,
    snapshot: ElevatedOfficeSnapshot,
    signal?: AbortSignal,
  ): Promise<{ verified: boolean; output?: unknown }>
  rollback(snapshot: ElevatedOfficeSnapshot, signal?: AbortSignal): Promise<void>
}

export const createExcelElevatedAdapter = (
  authority: ExcelElevatedAuthority,
): ElevatedOfficeAdapter => Object.freeze({ host: 'excel', ...authority })

type ExcelState = { fingerprint: string; tool: string; input: Record<string, any>; before: unknown }
const excelState = (snapshot: ElevatedOfficeSnapshot) => snapshot.state as ExcelState
function excelOperation(program: ElevatedOfficeProgram) {
  if (program.kind !== 'office_js_ast' || program.operations.length !== 1)
    throw new Error('office_api_unsupported')
  const operation = program.operations[0]
  if (operation.call === 'range.setValues') {
    const values = operation.args.values as unknown[][]
    return {
      tool: 'set_cell_range',
      input: {
        sheetId: operation.args.sheetId,
        range: operation.args.range,
        cells: values.map((row) => row.map((value) => ({ value }))),
        allow_overwrite: true,
      },
    }
  }
  return {
    tool: 'clear_cell_range',
    input: {
      sheetId: operation.args.sheetId,
      range: operation.args.range,
      clearType: operation.args.clearType,
    },
  }
}

export function createBrowserExcelElevatedAdapter(options: {
  adapter: ExcelAdapter
  authority(): ElevatedOfficeAuthority
}): ElevatedOfficeAdapter {
  return createExcelElevatedAdapter({
    captureAuthority: options.authority,
    async snapshot(program, signal) {
      const operation = excelOperation(program)
      const target = `sheet:${operation.input.sheetId}!${operation.input.range}`
      const fingerprint = await options.adapter.fingerprint([target], signal)
      const before = await options.adapter.captureMutation(operation.tool, operation.input, signal)
      return {
        id: `history_${selectionFingerprint(`${fingerprint}:${JSON.stringify(before)}`).replace(':', '_')}`,
        state: { fingerprint, ...operation, before } satisfies ExcelState,
      }
    },
    async validateSnapshot(_program, snapshot, signal) {
      const captured = excelState(snapshot)
      const target = `sheet:${captured.input.sheetId}!${captured.input.range}`
      return (await options.adapter.fingerprint([target], signal)) === captured.fingerprint
    },
    async execute(_program, snapshot, signal) {
      const captured = excelState(snapshot)
      if (captured.tool === 'set_cell_range')
        await options.adapter.setCellRange(captured.input, signal)
      else await options.adapter.clearCellRange(captured.input, signal)
    },
    async readback(_program, snapshot, signal) {
      const captured = excelState(snapshot)
      return {
        verified: await options.adapter.verifyMutation(
          captured.tool,
          captured.input,
          captured.before,
          signal,
        ),
      }
    },
    async rollback(snapshot, signal) {
      const captured = excelState(snapshot)
      const outcome = await options.adapter.recoverMutation?.(
        captured.tool,
        captured.input,
        captured.before,
        signal,
      )
      if (outcome !== 'restored') throw new Error('office_state_uncertain')
    },
  })
}
