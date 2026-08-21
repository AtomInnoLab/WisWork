import type { AgentSkill, ToolExecution } from '@wiswork/agent-core'
import type { StructuredProposalController } from '../../agent/proposal-controller.js'
import { selectionFingerprint } from '../../agent/proposal-controller.js'
import { parseDeclarativeProgram } from '../shared/declarative-program.js'
import type { ExcelAdapter } from './browser-excel-adapter.js'

const MAX_RANGE = 128,
  MAX_RANGES = 32,
  MAX_CODE = 32 * 1024,
  MAX_RESULT = 256 * 1024
const MAX_SCREENSHOT_BYTES = 4 * 1024 * 1024
type Json = Record<string, any>
type ExcelProgramOperation = { op: string; input: Json }
const descriptors = [
  [
    'get_cell_ranges',
    {
      sheetId: 'integer',
      ranges: 'ranges',
      includeStyles: 'boolean?',
      cellLimit: 'limit?',
      explanation: 'explanation?',
    },
  ],
  [
    'get_range_as_csv',
    {
      sheetId: 'integer',
      range: 'range',
      includeHeaders: 'boolean?',
      maxRows: 'rows?',
      explanation: 'explanation?',
    },
  ],
  [
    'search_data',
    {
      searchTerm: 'search',
      sheetId: 'integer?',
      range: 'range?',
      offset: 'offset?',
      options: 'searchOptions?',
      explanation: 'explanation?',
    },
  ],
  ['screenshot_range', { sheetId: 'integer', range: 'range', explanation: 'explanation?' }],
  ['get_all_objects', { sheetId: 'integer?', id: 'id?', explanation: 'explanation?' }],
  [
    'set_cell_range',
    {
      sheetId: 'integer',
      range: 'range',
      cells: 'cells',
      copyToRange: 'range?',
      resizeWidth: 'size?',
      resizeHeight: 'size?',
      allow_overwrite: 'boolean?',
      explanation: 'explanation?',
    },
  ],
  [
    'clear_cell_range',
    { sheetId: 'integer', range: 'range', clearType: 'clear?', explanation: 'explanation?' },
  ],
  [
    'copy_to',
    {
      sheetId: 'integer',
      sourceRange: 'range',
      destinationRange: 'range',
      explanation: 'explanation?',
    },
  ],
  [
    'modify_sheet_structure',
    {
      sheetId: 'integer',
      operation: 'sheetOp',
      dimension: 'dimension',
      reference: 'reference?',
      count: 'count?',
      position: 'position?',
      explanation: 'explanation?',
    },
  ],
  [
    'modify_workbook_structure',
    {
      operation: 'bookOp',
      sheetId: 'integer?',
      sheetName: 'sheetName?',
      newName: 'sheetName?',
      tabColor: 'color?',
      explanation: 'explanation?',
    },
  ],
  [
    'resize_range',
    {
      sheetId: 'integer',
      range: 'range?',
      width: 'size?',
      height: 'size?',
      explanation: 'explanation?',
    },
  ],
  [
    'modify_object',
    {
      operation: 'objectOp',
      sheetId: 'integer',
      objectType: 'objectType',
      id: 'id?',
      properties: 'properties?',
      explanation: 'explanation?',
    },
  ],
  ['eval_officejs', { code: 'code', explanation: 'codeExplanation?' }],
] as const

const enums: Record<string, readonly string[]> = {
  clear: ['contents', 'formats', 'all'],
  sheetOp: ['insert', 'delete', 'hide', 'unhide', 'freeze', 'unfreeze'],
  dimension: ['rows', 'columns'],
  position: ['before', 'after'],
  bookOp: ['create', 'delete', 'rename', 'duplicate'],
  objectOp: ['create', 'update', 'delete'],
  objectType: ['pivotTable', 'chart'],
}

function schemaFor(fields: Record<string, string>) {
  const properties: Json = {}
  const required: string[] = []
  for (const [name, kind0] of Object.entries(fields)) {
    const optional = kind0.endsWith('?')
    const kind = kind0.replace('?', '')
    if (!optional) required.push(name)
    const sizeSchema = {
      type: 'object',
      properties: {
        type: { type: 'string', enum: ['points', 'standard'] },
        value: { type: 'number', minimum: 0, maximum: 1_000 },
      },
      required: ['type', 'value'],
      additionalProperties: false,
    }
    const styleSchema = {
      type: 'object',
      properties: {
        fontWeight: { type: 'string', enum: ['normal', 'bold'] },
        fontStyle: { type: 'string', enum: ['normal', 'italic'] },
        fontLine: { type: 'string', enum: ['none', 'underline', 'line-through'] },
        fontSize: { type: 'number', minimum: 1, maximum: 409 },
        fontFamily: { type: 'string', maxLength: 256 },
        fontColor: { type: 'string', maxLength: 64 },
        backgroundColor: { type: 'string', maxLength: 64 },
        horizontalAlignment: { type: 'string', enum: ['left', 'center', 'right'] },
        numberFormat: { type: 'string', maxLength: 256 },
      },
      required: [],
      additionalProperties: false,
    }
    const border = {
      type: 'object',
      properties: {
        style: { type: 'string', enum: ['solid', 'dashed', 'dotted', 'double'] },
        weight: { type: 'string', enum: ['thin', 'medium', 'thick'] },
        color: { type: 'string', maxLength: 64 },
      },
      required: [],
      additionalProperties: false,
    }
    const borderSchema = {
      type: 'object',
      properties: { top: border, bottom: border, left: border, right: border },
      required: [],
      additionalProperties: false,
    }
    const cellSchema = {
      type: 'object',
      properties: {
        value: {},
        formula: { type: 'string', maxLength: 8_000 },
        note: { type: 'string', maxLength: 8_000 },
        cellStyles: styleSchema,
        borderStyles: borderSchema,
      },
      required: [],
      additionalProperties: false,
    }
    properties[name] =
      kind === 'integer'
        ? { type: 'integer', minimum: 1, maximum: 1_000_000 }
        : kind === 'boolean'
          ? { type: 'boolean' }
          : kind === 'ranges'
            ? {
                type: 'array',
                minItems: 1,
                maxItems: MAX_RANGES,
                items: { type: 'string', minLength: 1, maxLength: MAX_RANGE },
              }
            : kind === 'cells'
              ? {
                  type: 'array',
                  minItems: 1,
                  maxItems: 500,
                  items: {
                    type: 'array',
                    minItems: 1,
                    maxItems: 500,
                    items: cellSchema,
                  },
                }
              : kind === 'size'
                ? sizeSchema
                : kind === 'searchOptions'
                  ? {
                      type: 'object',
                      properties: {
                        matchCase: { type: 'boolean' },
                        matchEntireCell: { type: 'boolean' },
                        matchFormulas: { type: 'boolean' },
                        useRegex: { type: 'boolean' },
                        maxResults: { type: 'integer', minimum: 1, maximum: 500 },
                      },
                      required: [],
                      additionalProperties: false,
                    }
                  : Object.hasOwn(enums, kind)
                    ? { type: 'string', enum: enums[kind] }
                    : kind === 'properties'
                      ? {
                          type: 'object',
                          properties: {
                            name: { type: 'string', maxLength: 256 },
                            source: { type: 'string', maxLength: 512 },
                            range: { type: 'string', maxLength: MAX_RANGE },
                            anchor: { type: 'string', maxLength: MAX_RANGE },
                            title: { type: 'string', maxLength: 256 },
                            chartType: {
                              type: 'string',
                              enum: [
                                'columnClustered',
                                'barClustered',
                                'line',
                                'pie',
                                'scatter',
                                'area',
                                'doughnut',
                              ],
                            },
                          },
                          required: [],
                          additionalProperties: false,
                        }
                      : kind === 'code'
                        ? { type: 'string', minLength: 1, maxLength: MAX_CODE }
                        : kind === 'explanation'
                          ? { type: 'string', maxLength: 50 }
                          : kind === 'codeExplanation'
                            ? { type: 'string', maxLength: 100 }
                            : kind === 'limit'
                              ? { type: 'integer', minimum: 1, maximum: 2_000 }
                              : kind === 'rows'
                                ? { type: 'integer', minimum: 1, maximum: 500 }
                                : kind === 'offset'
                                  ? { type: 'integer', minimum: 0, maximum: 1_000_000 }
                                  : { type: 'string' }
  }
  return { type: 'object', properties, required, additionalProperties: false }
}
const tools = descriptors.map(([name, fields]) => ({
  name,
  description:
    name === 'eval_officejs'
      ? 'Execute a confirmation-gated bounded declarative Excel JSON program; JavaScript syntax and ambient authority are rejected.'
      : `Bounded Excel operation: ${name}.`,
  inputSchema: schemaFor(fields),
}))

function invalid(): never {
  throw new Error('invalid_tool_input')
}
function string(value: unknown, max: number, min = 0) {
  if (typeof value !== 'string' || value.length < min || value.length > max) invalid()
  return value as string
}
function integer(value: unknown, min: number, max: number) {
  if (!Number.isInteger(value) || (value as number) < min || (value as number) > max) invalid()
  return value as number
}
function range(value: unknown): string {
  const text = string(value, MAX_RANGE, 1)
  if (
    !/^\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/i.test(text) &&
    !/^(?:[A-Z]{1,3}:[A-Z]{1,3}|[1-9]\d*:[1-9]\d*)$/i.test(text)
  )
    invalid()
  return text
}
function exact(input: unknown, fields: Record<string, string>): Json {
  if (!input || typeof input !== 'object' || Array.isArray(input)) invalid()
  const source = input as Json
  if (Object.keys(source).some((key) => !(key in fields))) invalid()
  const result: Json = {}
  for (const [key, kind0] of Object.entries(fields)) {
    const optional = kind0.endsWith('?')
    const kind = kind0.replace('?', '')
    const value = source[key]
    if (value === undefined) {
      if (!optional) invalid()
      continue
    }
    if (kind === 'integer') result[key] = integer(value, 1, 1_000_000)
    else if (kind === 'boolean') {
      if (typeof value !== 'boolean') invalid()
      result[key] = value
    } else if (kind === 'range') result[key] = range(value)
    else if (kind === 'ranges') {
      if (!Array.isArray(value) || value.length < 1 || value.length > MAX_RANGES) invalid()
      result[key] = value.map((item) => {
        const parsed = range(item)
        if (/^[A-Z]+:[A-Z]+$|^\d+:\d+$/i.test(parsed)) invalid()
        return parsed
      })
    } else if (kind === 'limit') result[key] = integer(value, 1, 2_000)
    else if (kind === 'rows') result[key] = integer(value, 1, 500)
    else if (kind === 'offset') result[key] = integer(value, 0, 1_000_000)
    else if (kind === 'explanation') result[key] = string(value, 50)
    else if (kind === 'codeExplanation') result[key] = string(value, 100)
    else if (kind === 'code') result[key] = string(value, MAX_CODE, 1)
    else if (kind === 'search') result[key] = string(value, 512, 1)
    else if (kind === 'id' || kind === 'sheetName' || kind === 'reference' || kind === 'color')
      result[key] = string(value, 256, 1)
    else if (kind === 'count') result[key] = integer(value, 1, 1_000)
    else if (enums[kind]) {
      if (typeof value !== 'string' || !enums[kind].includes(value)) invalid()
      result[key] = value
    } else if (kind === 'size') {
      if (
        !value ||
        typeof value !== 'object' ||
        Array.isArray(value) ||
        Object.keys(value).some((k) => !['type', 'value'].includes(k))
      )
        invalid()
      const item = value as Json
      if (
        !['points', 'standard'].includes(item.type) ||
        typeof item.value !== 'number' ||
        !Number.isFinite(item.value) ||
        item.value < 0 ||
        item.value > 1_000
      )
        invalid()
      result[key] = { type: item.type, value: item.value }
    } else if (kind === 'searchOptions') {
      if (
        !value ||
        typeof value !== 'object' ||
        Array.isArray(value) ||
        Object.keys(value).some(
          (k) =>
            !['matchCase', 'matchEntireCell', 'matchFormulas', 'useRegex', 'maxResults'].includes(
              k,
            ),
        )
      )
        invalid()
      const item = value as Json
      for (const k of ['matchCase', 'matchEntireCell', 'matchFormulas', 'useRegex'])
        if (item[k] !== undefined && typeof item[k] !== 'boolean') invalid()
      if (item.maxResults !== undefined) integer(item.maxResults, 1, 500)
      result[key] = { ...item }
    } else if (kind === 'cells') result[key] = parseCells(value)
    else if (kind === 'properties') result[key] = parseProperties(value)
    else invalid()
  }
  return result
}
function parseCells(value: unknown): Json[][] {
  if (!Array.isArray(value) || value.length < 1 || value.length > 500) invalid()
  let width = -1,
    count = 0
  return value.map((row) => {
    if (
      !Array.isArray(row) ||
      row.length < 1 ||
      row.length > 500 ||
      (width >= 0 && width !== row.length)
    )
      invalid()
    width = row.length
    count += width
    if (count > 2_000) invalid()
    return row.map((cell) => {
      if (!cell || typeof cell !== 'object' || Array.isArray(cell)) invalid()
      const item = cell as Json
      if (Object.hasOwn(item, 'value') && Object.hasOwn(item, 'formula')) invalid()
      if (
        Object.keys(item).some(
          (k) => !['value', 'formula', 'note', 'cellStyles', 'borderStyles'].includes(k),
        )
      )
        invalid()
      if (item.formula !== undefined) string(item.formula, 8_000)
      if (item.note !== undefined) string(item.note, 8_000)
      if (item.cellStyles !== undefined) {
        if (
          !item.cellStyles ||
          typeof item.cellStyles !== 'object' ||
          Array.isArray(item.cellStyles)
        )
          invalid()
        const allowed = [
          'fontWeight',
          'fontStyle',
          'fontLine',
          'fontSize',
          'fontFamily',
          'fontColor',
          'backgroundColor',
          'horizontalAlignment',
          'numberFormat',
        ]
        if (Object.keys(item.cellStyles).some((key) => !allowed.includes(key))) invalid()
        const styles = item.cellStyles as Json
        if (styles.fontWeight !== undefined && !['normal', 'bold'].includes(styles.fontWeight))
          invalid()
        if (styles.fontStyle !== undefined && !['normal', 'italic'].includes(styles.fontStyle))
          invalid()
        if (
          styles.fontLine !== undefined &&
          !['none', 'underline', 'line-through'].includes(styles.fontLine)
        )
          invalid()
        if (
          styles.horizontalAlignment !== undefined &&
          !['left', 'center', 'right'].includes(styles.horizontalAlignment)
        )
          invalid()
        if (
          styles.fontSize !== undefined &&
          (typeof styles.fontSize !== 'number' ||
            !Number.isFinite(styles.fontSize) ||
            styles.fontSize < 1 ||
            styles.fontSize > 409)
        )
          invalid()
        for (const key of ['fontFamily', 'fontColor', 'backgroundColor', 'numberFormat'])
          if (styles[key] !== undefined)
            string(styles[key], key === 'fontFamily' || key === 'numberFormat' ? 256 : 64, 1)
      }
      if (item.borderStyles !== undefined) {
        if (
          !item.borderStyles ||
          typeof item.borderStyles !== 'object' ||
          Array.isArray(item.borderStyles) ||
          Object.keys(item.borderStyles).some(
            (key) => !['top', 'bottom', 'left', 'right'].includes(key),
          )
        )
          invalid()
        for (const border of Object.values(item.borderStyles)) {
          if (
            !border ||
            typeof border !== 'object' ||
            Array.isArray(border) ||
            Object.keys(border as Json).some((key) => !['style', 'weight', 'color'].includes(key))
          )
            invalid()
          const config = border as Json
          if (
            config.style !== undefined &&
            !['solid', 'dashed', 'dotted', 'double'].includes(config.style)
          )
            invalid()
          if (config.weight !== undefined && !['thin', 'medium', 'thick'].includes(config.weight))
            invalid()
          if (config.color !== undefined) string(config.color, 64, 1)
        }
      }
      return JSON.parse(JSON.stringify(item))
    })
  })
}
function parseProperties(value: unknown): Json {
  if (!value || typeof value !== 'object' || Array.isArray(value)) invalid()
  const item = value as Json
  const allowed = ['name', 'source', 'range', 'anchor', 'title', 'chartType']
  if (Object.keys(item).some((k) => !allowed.includes(k))) invalid()
  for (const key of ['name', 'source', 'range', 'anchor', 'title', 'chartType'])
    if (item[key] !== undefined) string(item[key], key === 'source' ? 512 : 256, 1)
  if (
    item.chartType !== undefined &&
    !['columnClustered', 'barClustered', 'line', 'pie', 'scatter', 'area', 'doughnut'].includes(
      item.chartType,
    )
  )
    invalid()
  for (const key of ['source', 'range', 'anchor']) if (item[key] !== undefined) range(item[key])
  if (new TextEncoder().encode(JSON.stringify(item)).byteLength > 16 * 1024) invalid()
  return JSON.parse(JSON.stringify(item))
}
function fail(name: string, code: string): ToolExecution {
  return { output: code, isError: true, mutated: false, summary: name }
}
function check(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('cancelled')
}
function bounded(value: unknown) {
  const text = JSON.stringify(value)
  if (new TextEncoder().encode(text).byteLength > MAX_RESULT) throw new Error('office_read_failed')
  return text
}
function safeError(error: unknown, write = false) {
  const code = error instanceof Error ? error.message : ''
  return [
    'invalid_tool_input',
    'office_api_unsupported',
    'cancelled',
    'office_read_failed',
  ].includes(code)
    ? code
    : write
      ? 'office_write_failed'
      : 'office_read_failed'
}
function targets(name: string, input: Json): string[] {
  if (name === 'copy_to')
    return [
      `sheet:${input.sheetId}!${input.sourceRange}`,
      `sheet:${input.sheetId}!${input.destinationRange}`,
    ]
  if (name === 'set_cell_range')
    return [`sheet:${input.sheetId}!${input.copyToRange ?? input.range}`]
  if (name === 'modify_sheet_structure')
    return [
      `structure:${input.sheetId}:${input.operation}:${input.dimension}:${input.reference ?? ''}:${input.count ?? 1}`,
    ]
  if (name === 'resize_range') return [`resize:${input.sheetId}!${input.range ?? 'A1:XFD1048576'}`]
  if (name === 'modify_workbook_structure')
    return [`workbook:${input.operation}:${input.sheetId ?? input.sheetName ?? ''}`]
  if (name === 'modify_object')
    return [`sheet:${input.sheetId}!object:${input.id ?? input.properties?.name ?? 'new'}`]
  return [`sheet:${input.sheetId}!${input.range ?? '*'}`]
}
function cellCount(name: string, input: Json) {
  if (name === 'set_cell_range') return input.cells.reduce((n: number, r: any[]) => n + r.length, 0)
  return 1
}
function semantics(name: string, input: Json) {
  if (name === 'modify_workbook_structure') {
    if (input.operation === 'create' && !input.sheetName) invalid()
    if (input.operation !== 'create' && input.sheetId === undefined) invalid()
    if (['rename', 'duplicate'].includes(input.operation) && !input.newName) invalid()
  }
  if (name === 'modify_sheet_structure' && input.operation !== 'unfreeze' && !input.reference)
    invalid()
  if (
    name === 'modify_sheet_structure' &&
    input.reference &&
    (input.dimension === 'rows'
      ? !/^[1-9]\d*$/.test(input.reference)
      : !/^[A-Z]{1,3}$/i.test(input.reference))
  )
    invalid()
  if (name === 'resize_range' && !input.width && !input.height) invalid()
  if (name === 'modify_object') {
    if (input.operation !== 'create' && !input.id) invalid()
    if (input.operation !== 'delete' && !input.properties) invalid()
    if (
      input.operation === 'create' &&
      input.objectType === 'chart' &&
      (!input.properties.chartType || !input.properties.source)
    )
      invalid()
    if (
      input.operation === 'create' &&
      input.objectType === 'pivotTable' &&
      (!input.properties.name || !input.properties.source || !input.properties.range)
    )
      invalid()
    if (input.operation === 'update' && Object.keys(input.properties).length === 0) invalid()
    if (
      input.operation === 'update' &&
      input.objectType === 'pivotTable' &&
      ((input.properties.source && !input.properties.range) ||
        (!input.properties.source && input.properties.range))
    )
      invalid()
  }
}

const mutationNames = new Set([
  'set_cell_range',
  'clear_cell_range',
  'copy_to',
  'modify_sheet_structure',
  'modify_workbook_structure',
  'resize_range',
  'modify_object',
])

function parseExcelOperation(value: unknown): ExcelProgramOperation {
  if (!value || typeof value !== 'object' || Array.isArray(value)) invalid()
  const operation = value as Json
  if (Object.keys(operation).some((key) => !['op', 'input'].includes(key))) invalid()
  if (typeof operation.op !== 'string' || !mutationNames.has(operation.op)) invalid()
  const descriptor = descriptors.find(([name]) => name === operation.op)
  if (!descriptor) invalid()
  const input = exact(operation.input, descriptor[1])
  semantics(operation.op, input)
  return { op: operation.op, input }
}

export function createExcelSkill(options: {
  adapter: ExcelAdapter
  proposals: StructuredProposalController
}): AgentSkill {
  return {
    id: 'office-excel',
    systemPrompt:
      'Excel reads and screenshots are bounded. Every mutation is only a proposal until confirmed and semantically verified. eval_officejs uses the shared declarative execution runtime when available.',
    tools,
    async executeTool(call, signal) {
      if (call.inputError || call.truncated) return fail(call.name, 'invalid_tool_input')
      const descriptor = descriptors.find(([name]) => name === call.name)
      if (!descriptor) return fail(call.name, 'invalid_tool_input')
      try {
        check(signal)
        const input = exact(call.input, descriptor[1])
        semantics(call.name, input)
        const methods: Record<string, (data: Json, signal?: AbortSignal) => Promise<void>> = {
          set_cell_range: options.adapter.setCellRange.bind(options.adapter),
          clear_cell_range: options.adapter.clearCellRange.bind(options.adapter),
          copy_to: options.adapter.copyTo.bind(options.adapter),
          modify_sheet_structure: options.adapter.modifySheetStructure.bind(options.adapter),
          modify_workbook_structure: options.adapter.modifyWorkbookStructure.bind(options.adapter),
          resize_range: options.adapter.resizeRange.bind(options.adapter),
          modify_object: options.adapter.modifyObject.bind(options.adapter),
        }
        if (call.name === 'eval_officejs') {
          const program = parseDeclarativeProgram(input.code, parseExcelOperation)
          const affected = [
            ...new Set(program.operations.flatMap((item) => targets(item.op, item.input))),
          ]
          const before = await options.adapter.fingerprint(affected, signal)
          check(signal)
          let expectedFingerprint: string | undefined
          const proposal = options.proposals.propose({
            operation: call.name,
            toolName: call.name,
            title: input.explanation || 'Execute declarative Excel operations',
            preview: { version: program.version, operations: program.operations },
            impact: { host: 'Excel', targets: affected, count: program.operations.length },
            fingerprint: selectionFingerprint(before),
            before,
            code: input.code,
            validate: async (confirmSignal) =>
              (await options.adapter.fingerprint(affected, confirmSignal)) === before,
            execute: async (confirmSignal) => {
              for (const operation of program.operations) {
                check(confirmSignal)
                await methods[operation.op](operation.input, confirmSignal)
              }
              check(confirmSignal)
              expectedFingerprint = await options.adapter.fingerprint(affected, confirmSignal)
            },
            verify: async (confirmSignal) => {
              for (const operation of program.operations)
                if (
                  !(await options.adapter.verifyMutation(
                    operation.op,
                    operation.input,
                    before,
                    confirmSignal,
                  ))
                )
                  throw new Error('office_verify_failed')
              if (
                !expectedFingerprint ||
                (await options.adapter.fingerprint(affected, confirmSignal)) !== expectedFingerprint
              )
                throw new Error('office_verify_failed')
            },
          })
          return {
            output: JSON.stringify({ proposalId: proposal.id, mutated: false }),
            mutated: false,
            summary: 'Proposed declarative Excel execution',
          }
        }
        const reads: Record<string, () => Promise<unknown>> = {
          get_cell_ranges: () => options.adapter.getCellRanges(input as any, signal),
          get_range_as_csv: () => options.adapter.getRangeAsCsv(input as any, signal),
          search_data: () => options.adapter.searchData(input as any, signal),
          screenshot_range: () => options.adapter.screenshotRange(input as any, signal),
          get_all_objects: () => options.adapter.getAllObjects(input as any, signal),
        }
        if (reads[call.name]) {
          const result = await reads[call.name]()
          check(signal)
          if (call.name === 'screenshot_range') {
            const image = result as any
            const padding =
              typeof image.base64 === 'string' && image.base64.endsWith('==')
                ? 2
                : typeof image.base64 === 'string' && image.base64.endsWith('=')
                  ? 1
                  : 0
            const bytes =
              typeof image.base64 === 'string' ? (image.base64.length / 4) * 3 - padding : Infinity
            if (
              image.mime !== 'image/png' ||
              typeof image.base64 !== 'string' ||
              image.base64.length % 4 !== 0 ||
              !/^[A-Za-z0-9+/]*={0,2}$/.test(image.base64) ||
              !image.base64.startsWith('iVBORw0KGgo') ||
              bytes > MAX_SCREENSHOT_BYTES
            )
              throw new Error('office_read_failed')
            return {
              output: JSON.stringify({ mime: image.mime }),
              display: {
                kind: 'images',
                items: [{ url: `data:${image.mime};base64,${image.base64}` }],
              },
              mutated: false,
              summary: 'Captured Excel range',
            }
          }
          return { output: bounded(result), mutated: false, summary: call.name }
        }
        const affected = targets(call.name, input)
        const before = await options.adapter.fingerprint(affected, signal)
        check(signal)
        let expectedFingerprint: string | undefined
        const proposal = options.proposals.propose({
          operation: call.name,
          toolName: call.name,
          title: `Confirm ${call.name}`,
          preview: { input },
          impact: { host: 'Excel', targets: affected, count: cellCount(call.name, input) },
          fingerprint: selectionFingerprint(before),
          before,
          validate: async (s) => {
            check(s)
            const current = await options.adapter.fingerprint(affected, s)
            check(s)
            return current === before
          },
          execute: async (s) => {
            check(s)
            try {
              await methods[call.name](input, s)
              check(s)
              expectedFingerprint = await options.adapter.fingerprint(affected, s)
              check(s)
            } catch (error) {
              // Do not attach the raw Office exception: it can contain document data.
              // eslint-disable-next-line preserve-caught-error
              throw new Error(safeError(error, true))
            }
          },
          verify: async (s) => {
            check(s)
            try {
              if (!(await options.adapter.verifyMutation(call.name, input, before, s)))
                throw new Error('office_verify_failed')
              if (call.name === 'modify_object')
                await options.adapter.verifyObjects({ sheetId: input.sheetId, id: input.id }, s)
              else if (call.name === 'modify_workbook_structure')
                await options.adapter.verifyWorkbook(s)
              else if (call.name === 'modify_sheet_structure' || call.name === 'resize_range')
                await options.adapter.fingerprint(affected, s)
              else await options.adapter.verifyRanges(affected, s)
              check(s)
              if (
                !expectedFingerprint ||
                (await options.adapter.fingerprint(affected, s)) !== expectedFingerprint
              )
                throw new Error('office_verify_failed')
            } catch {
              throw new Error('office_verify_failed')
            }
          },
        })
        return {
          output: JSON.stringify({ proposalId: proposal.id, operation: call.name, mutated: false }),
          mutated: false,
          summary: `Proposed ${call.name}`,
        }
      } catch (error) {
        return fail(call.name, safeError(error))
      }
    },
  }
}
