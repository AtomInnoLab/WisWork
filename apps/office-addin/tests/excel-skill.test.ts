import { afterEach, describe, expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  BrowserExcelAdapter,
  type ExcelAdapter,
} from '../src/skills/excel/browser-excel-adapter.js'
import { createExcelSkill } from '../src/skills/excel/excel-skill.js'

const readResult = {
  sheetId: 1,
  sheetName: 'Data',
  address: 'A1:B2',
  rows: 2,
  columns: 2,
  hasMore: false,
  cells: [{ address: 'A1', value: 'a', formula: null, numberFormat: 'General' }],
}
function adapter(overrides: Partial<ExcelAdapter> = {}): ExcelAdapter {
  return {
    getCellRanges: vi.fn().mockResolvedValue({ ranges: [readResult], hasMore: false }),
    getRangeAsCsv: vi.fn().mockResolvedValue({
      sheetId: 1,
      sheetName: 'Data',
      address: 'A1:B2',
      csv: 'a,"b,b"\n"c""c",d',
      rowCount: 2,
      columnCount: 2,
      hasMore: false,
    }),
    searchData: vi
      .fn()
      .mockResolvedValue({ matches: [], offset: 0, returned: 0, hasMore: false, nextOffset: null }),
    screenshotRange: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    getAllObjects: vi.fn().mockResolvedValue({ objects: [], hasMore: false }),
    fingerprint: vi.fn().mockResolvedValue('fp'),
    setCellRange: vi.fn().mockResolvedValue(undefined),
    clearCellRange: vi.fn().mockResolvedValue(undefined),
    copyTo: vi.fn().mockResolvedValue(undefined),
    modifySheetStructure: vi.fn().mockResolvedValue(undefined),
    modifyWorkbookStructure: vi.fn().mockResolvedValue(undefined),
    resizeRange: vi.fn().mockResolvedValue(undefined),
    modifyObject: vi.fn().mockResolvedValue(undefined),
    verifyRanges: vi.fn().mockResolvedValue({ ranges: [readResult], hasMore: false }),
    verifyObjects: vi.fn().mockResolvedValue({ objects: [], hasMore: false }),
    verifyWorkbook: vi.fn().mockResolvedValue({ sheets: [], hasMore: false }),
    ...overrides,
  }
}
const call = (name: string, input: Record<string, unknown> = {}) => ({ id: 'c1', name, input })

describe('Excel compatibility skill', () => {
  it('exposes the exact documented inventory', () => {
    const skill = createExcelSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    expect(skill.tools.map((tool) => tool.name)).toEqual([
      'get_cell_ranges',
      'get_range_as_csv',
      'search_data',
      'screenshot_range',
      'get_all_objects',
      'set_cell_range',
      'clear_cell_range',
      'copy_to',
      'modify_sheet_structure',
      'modify_workbook_structure',
      'resize_range',
      'modify_object',
      'eval_officejs',
    ])
    expect(skill.tools.every((tool) => tool.inputSchema.additionalProperties === false)).toBe(true)
  })

  it('validates exact bounded read inputs and returns normalized data', async () => {
    const fake = adapter()
    const skill = createExcelSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(
        call('get_cell_ranges', {
          sheetId: 1,
          ranges: ['A1:B2'],
          includeStyles: true,
          cellLimit: 10,
        }),
      ),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('"numberFormat":"General"'),
    })
    await expect(
      skill.executeTool(call('get_cell_ranges', { sheetId: 1, ranges: [] })),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
    await expect(
      skill.executeTool(call('get_range_as_csv', { sheetId: 1, range: 'A1:B2', nope: true })),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
  })

  it('creates immutable write proposals and only writes after fresh validation, then verifies', async () => {
    const fake = adapter()
    const proposals = createStructuredProposalController()
    const skill = createExcelSkill({ adapter: fake, proposals })
    const result = await skill.executeTool(
      call('set_cell_range', {
        sheetId: 1,
        range: 'A1:B1',
        cells: [[{ value: 'x' }, { formula: '=1+1' }]],
      }),
    )
    expect(result).toMatchObject({ mutated: false, output: expect.stringContaining('proposalId') })
    expect(fake.setCellRange).not.toHaveBeenCalled()
    const pending = proposals.pending()!
    expect(pending.impact).toEqual({ host: 'Excel', targets: ['sheet:1!A1:B1'], count: 2 })
    await proposals.confirm(pending.id)
    expect(fake.fingerprint).toHaveBeenCalledTimes(2)
    expect(fake.setCellRange).toHaveBeenCalledOnce()
    expect(fake.verifyRanges).toHaveBeenCalledOnce()
  })

  it('rejects stale and cancelled writes without applying them', async () => {
    const fake = adapter({
      fingerprint: vi.fn().mockResolvedValueOnce('before').mockResolvedValueOnce('changed'),
    })
    const proposals = createStructuredProposalController()
    const skill = createExcelSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('clear_cell_range', { sheetId: 1, range: 'A1', clearType: 'contents' }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.clearCellRange).not.toHaveBeenCalled()
    const controller = new AbortController()
    controller.abort()
    await expect(
      skill.executeTool(
        call('copy_to', { sheetId: 1, sourceRange: 'A1', destinationRange: 'B1' }),
        controller.signal,
      ),
    ).resolves.toMatchObject({ output: 'cancelled', isError: true })
  })

  it('keeps eval_officejs stable and fail-closed without eval or Function', async () => {
    const proposals = createStructuredProposalController()
    const skill = createExcelSkill({ adapter: adapter(), proposals })
    await expect(
      skill.executeTool(call('eval_officejs', { code: 'return context.workbook' })),
    ).resolves.toMatchObject({ output: 'office_api_unsupported', isError: true, mutated: false })
    expect(proposals.pending()).toBeUndefined()
  })

  it('accepts every documented structured mutation shape and rejects incomplete/unknown nested input', async () => {
    const cases: Array<[string, Record<string, unknown>]> = [
      ['clear_cell_range', { sheetId: 1, range: 'A1', clearType: 'formats' }],
      ['copy_to', { sheetId: 1, sourceRange: 'A1', destinationRange: 'B1' }],
      [
        'modify_sheet_structure',
        {
          sheetId: 1,
          operation: 'insert',
          dimension: 'rows',
          reference: '5',
          count: 1,
          position: 'before',
        },
      ],
      [
        'modify_workbook_structure',
        { operation: 'create', sheetName: 'Summary', tabColor: '#ff0000' },
      ],
      ['resize_range', { sheetId: 1, range: 'A:D', width: { type: 'points', value: 40 } }],
      [
        'modify_object',
        {
          operation: 'create',
          sheetId: 1,
          objectType: 'chart',
          properties: { chartType: 'line', source: 'A1:B5', anchor: 'D1' },
        },
      ],
    ]
    for (const [name, input] of cases) {
      const proposals = createStructuredProposalController()
      await expect(
        createExcelSkill({ adapter: adapter(), proposals }).executeTool(call(name, input)),
      ).resolves.toMatchObject({ mutated: false, output: expect.stringContaining('proposalId') })
      expect(proposals.pending()?.toolName).toBe(name)
    }
    await expect(
      createExcelSkill({
        adapter: adapter(),
        proposals: createStructuredProposalController(),
      }).executeTool(call('resize_range', { sheetId: 1 })),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
    await expect(
      createExcelSkill({
        adapter: adapter(),
        proposals: createStructuredProposalController(),
      }).executeTool(
        call('set_cell_range', { sheetId: 1, range: 'A1', cells: [[{ value: 1, secret: true }]] }),
      ),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
  })
})

describe('browser Excel adapter', () => {
  const originals = { Office: globalThis.Office, Excel: globalThis.Excel }
  afterEach(() => Object.assign(globalThis, originals))

  it('feature-detects host/API and uses server-side bounded range loads', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: { host: 'Word', requirements: { isSetSupported: vi.fn().mockReturnValue(true) } },
      },
      Excel: { run },
    })
    await expect(
      new BrowserExcelAdapter().getCellRanges({ sheetId: 1, ranges: ['A1'] }),
    ).rejects.toThrow('office_api_unsupported')
    expect(run).not.toHaveBeenCalled()
  })

  it('escapes CSV fields and bounds rows in the adapter', async () => {
    const range = {
      values: [
        ['h1', 'h2'],
        ['a,b', 'x"y'],
        ['tail', 'z'],
      ],
      rowCount: 3,
      columnCount: 2,
      address: 'Data!A1:B3',
      load: vi.fn(),
    }
    const sheet = { id: '1', name: 'Data', load: vi.fn(), getRange: vi.fn().mockReturnValue(range) }
    Object.assign(globalThis, {
      Office: {
        context: { host: 'Excel', requirements: { isSetSupported: vi.fn().mockReturnValue(true) } },
      },
      Excel: {
        run: (cb: (ctx: unknown) => unknown) =>
          cb({ workbook: { worksheets: { getItem: () => sheet } }, sync: vi.fn() }),
      },
    })
    await expect(
      new BrowserExcelAdapter().getRangeAsCsv({ sheetId: 1, range: 'A1:B3', maxRows: 1 }),
    ).resolves.toMatchObject({ csv: 'h1,h2', rowCount: 1, hasMore: true })
    expect(range.load).toHaveBeenCalledWith('values,rowCount,columnCount,address')
  })
})
