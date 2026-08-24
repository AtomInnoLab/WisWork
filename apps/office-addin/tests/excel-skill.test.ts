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
    captureMutation: vi.fn().mockResolvedValue('operation-fp'),
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
    verifyMutation: vi.fn().mockResolvedValue(true),
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
    expect(fake.verifyRanges).not.toHaveBeenCalled()
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

  it('revalidates the stable Office worksheet identity after an ordinal reorder', async () => {
    const resolveWorksheetIdentity = vi.fn().mockResolvedValue('office-sheet-data')
    const validateWorksheetIdentity = vi.fn().mockResolvedValue(false)
    const fake = adapter({
      resolveWorksheetIdentity,
      validateWorksheetIdentity,
    } as any)
    const proposals = createStructuredProposalController()
    const skill = createExcelSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('set_cell_range', {
        sheetId: 1,
        range: 'A1',
        cells: [[{ value: 'safe' }]],
        allow_overwrite: true,
      }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(resolveWorksheetIdentity).toHaveBeenCalledWith(1, undefined)
    expect(validateWorksheetIdentity).toHaveBeenCalledWith(
      1,
      'office-sheet-data',
      expect.any(AbortSignal),
    )
    expect(fake.setCellRange).not.toHaveBeenCalled()
  })

  it('accepts a mutation after bounded delayed readback convergence', async () => {
    const verifyMutation = vi
      .fn()
      .mockResolvedValueOnce(false)
      .mockResolvedValueOnce(false)
      .mockResolvedValueOnce(true)
    const fake = adapter({ verifyMutation })
    const proposals = createStructuredProposalController()
    const skill = createExcelSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('clear_cell_range', { sheetId: 1, range: 'A1', clearType: 'contents' }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(verifyMutation).toHaveBeenCalledTimes(3)
  })

  it('never restores over a concurrent third state after verification failure', async () => {
    const recoverMutation = vi.fn().mockResolvedValue('concurrent')
    const fake = adapter({
      verifyMutation: vi.fn().mockResolvedValue(false),
      recoverMutation,
    } as any)
    const proposals = createStructuredProposalController()
    const skill = createExcelSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('copy_to', { sheetId: 1, sourceRange: 'A1', destinationRange: 'B1' }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow(
      'office_concurrent_change',
    )
    expect(recoverMutation).toHaveBeenCalledOnce()
  })

  it('reconciles a sync rejection and reports a proved restoration as a failed write', async () => {
    const recoverMutation = vi.fn().mockResolvedValue('restored')
    const fake = adapter({
      setCellRange: vi.fn().mockRejectedValue(new Error('InvalidArgument')),
      recoverMutation,
    } as any)
    const proposals = createStructuredProposalController()
    const skill = createExcelSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('set_cell_range', {
        sheetId: 1,
        range: 'A1',
        cells: [[{ value: 'new' }]],
        allow_overwrite: true,
      }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('office_write_failed')
    expect(recoverMutation).toHaveBeenCalledOnce()
  })

  it('classifies third-state cell edits without dispatching a restore', async () => {
    const excel = new BrowserExcelAdapter()
    const beforeCell = {
      address: 'A1',
      value: 'before',
      formula: 'before',
      formulaR1C1: 'before',
      numberFormat: 'General',
      style: {},
    }
    vi.spyOn(excel, 'getCellRanges').mockResolvedValue({
      ranges: [
        {
          address: 'A1',
          rows: 1,
          columns: 1,
          cells: [{ ...beforeCell, value: 'user edit', formula: 'user edit' }],
        },
      ],
    })
    const run = vi.spyOn(excel as any, 'run')

    await expect(
      excel.recoverMutation(
        'set_cell_range',
        {
          sheetId: 1,
          worksheetOfficeId: 'stable-sheet',
          range: 'A1',
          cells: [[{ value: 'agent edit' }]],
        },
        {
          kind: 'cells',
          state: {
            ranges: [{ address: 'A1', rows: 1, columns: 1, cells: [beforeCell] }],
          },
        },
      ),
    ).resolves.toBe('concurrent')
    expect(run).not.toHaveBeenCalled()
  })

  it('restores only a cell state composed of the approved write and captured prestate', async () => {
    const excel = new BrowserExcelAdapter()
    const beforeCell = {
      address: 'A1',
      value: 'before',
      formula: 'before',
      formulaR1C1: 'before',
      numberFormat: 'General',
      style: {},
    }
    const state = {
      ranges: [{ address: 'A1', rows: 1, columns: 1, cells: [beforeCell] }],
    }
    vi.spyOn(excel, 'getCellRanges')
      .mockResolvedValueOnce({
        ranges: [
          {
            address: 'A1',
            rows: 1,
            columns: 1,
            cells: [{ ...beforeCell, value: 'agent edit', formula: 'agent edit' }],
          },
        ],
      })
      .mockResolvedValueOnce(state)
    const run = vi.spyOn(excel as any, 'run').mockResolvedValue(undefined)

    await expect(
      excel.recoverMutation(
        'set_cell_range',
        {
          sheetId: 1,
          worksheetOfficeId: 'stable-sheet',
          range: 'A1',
          cells: [[{ value: 'agent edit' }]],
        },
        { kind: 'cells', state },
      ),
    ).resolves.toBe('restored')
    expect(run).toHaveBeenCalledOnce()
  })

  it('fails structural recovery closed when Office cannot provide a bounded inverse', async () => {
    const excel = new BrowserExcelAdapter()
    await expect(
      excel.recoverMutation('modify_sheet_structure', { sheetId: 1 }, 'before'),
    ).resolves.toBe('uncertain')
  })

  it('fails confirmation when request-semantic post-write verification detects a no-op', async () => {
    const fake = adapter({ verifyMutation: vi.fn().mockResolvedValue(false) })
    const proposals = createStructuredProposalController()
    const skill = createExcelSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('clear_cell_range', { sheetId: 1, range: 'A1', clearType: 'contents' }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('office_verify_failed')
    expect(fake.verifyMutation).toHaveBeenCalledWith(
      'clear_cell_range',
      expect.objectContaining({ range: 'A1' }),
      'operation-fp',
      expect.any(AbortSignal),
    )
  })

  it('rejects JavaScript but proposes an allowlisted declarative eval_officejs program', async () => {
    const proposals = createStructuredProposalController()
    const fake = adapter()
    const skill = createExcelSkill({ adapter: fake, proposals })
    await expect(
      skill.executeTool(call('eval_officejs', { code: 'return context.workbook' })),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true, mutated: false })
    await expect(
      skill.executeTool(
        call('eval_officejs', {
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'set_cell_range',
                input: {
                  sheetId: 1,
                  range: 'A1',
                  cells: [[{ value: 'safe' }]],
                  allow_overwrite: true,
                },
              },
            ],
          }),
        }),
      ),
    ).resolves.toMatchObject({ output: expect.stringContaining('proposalId'), mutated: false })
    await proposals.confirm(proposals.pending()!.id)
    expect(fake.setCellRange).toHaveBeenCalledOnce()
  })

  it('rejects every multi-operation declarative program before host reads or writes', async () => {
    const fake = adapter()
    const proposals = createStructuredProposalController()
    const result = await createExcelSkill({ adapter: fake, proposals }).executeTool(
      call('eval_officejs', {
        code: JSON.stringify({
          version: 1,
          operations: [
            {
              op: 'set_cell_range',
              input: {
                sheetId: 1,
                range: '$A$1',
                cells: [[{ value: 1 }, { value: 2 }]],
                allow_overwrite: true,
              },
            },
            {
              op: 'clear_cell_range',
              input: { sheetId: 1, range: 'B1:C1', clearType: 'all' },
            },
          ],
        }),
      }),
    )
    expect(result).toMatchObject({ output: 'office_api_unsupported', isError: true })
    expect(proposals.pending()).toBeUndefined()
    expect(fake.fingerprint).not.toHaveBeenCalled()
    expect(fake.captureMutation).not.toHaveBeenCalled()
    expect(fake.setCellRange).not.toHaveBeenCalled()
    await expect(
      createExcelSkill({ adapter: fake, proposals }).executeTool(
        call('eval_officejs', {
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'clear_cell_range',
                input: { sheetId: 1, range: 'A1', clearType: 'contents' },
              },
              {
                op: 'clear_cell_range',
                input: { sheetId: 1, range: 'B1', clearType: 'contents' },
              },
            ],
          }),
        }),
        AbortSignal.abort(),
      ),
    ).resolves.toMatchObject({ output: 'cancelled', isError: true })
    expect(fake.clearCellRange).not.toHaveBeenCalled()
  })

  it('rejects declarative programs exceeding the aggregate cell budget before host reads', async () => {
    const fake = adapter()
    const cells = Array.from({ length: 3 }, () => Array.from({ length: 334 }, () => ({ value: 1 })))
    const result = await createExcelSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
    }).executeTool(
      call('eval_officejs', {
        code: JSON.stringify({
          version: 1,
          operations: [
            {
              op: 'set_cell_range',
              input: { sheetId: 1, range: 'A1', cells, allow_overwrite: true },
            },
            {
              op: 'set_cell_range',
              input: { sheetId: 1, range: 'A10', cells, allow_overwrite: true },
            },
          ],
        }),
      }),
    )
    expect(result).toMatchObject({ output: 'office_api_unsupported', isError: true })
    expect(fake.fingerprint).not.toHaveBeenCalled()
  })

  it('rejects non-batchable declarative combinations before proposal or mutation', async () => {
    const fake = adapter()
    const proposals = createStructuredProposalController()
    const result = await createExcelSkill({ adapter: fake, proposals }).executeTool(
      call('eval_officejs', {
        code: JSON.stringify({
          version: 1,
          operations: [
            {
              op: 'set_cell_range',
              input: { sheetId: 1, range: 'A1', cells: [[{ value: 1 }]], allow_overwrite: true },
            },
            {
              op: 'resize_range',
              input: { sheetId: 1, range: 'A:A', width: { type: 'standard', value: 0 } },
            },
          ],
        }),
      }),
    )
    expect(result).toMatchObject({ output: 'office_api_unsupported', isError: true })
    expect(proposals.pending()).toBeUndefined()
    expect(fake.fingerprint).not.toHaveBeenCalled()
    expect(fake.setCellRange).not.toHaveBeenCalled()
  })

  it('returns a bounded model-visible PNG for screenshot_range', async () => {
    const png = 'iVBORw0KGgoAAA=='
    const skill = createExcelSkill({
      adapter: adapter({
        screenshotRange: vi.fn().mockResolvedValue({ mime: 'image/png', base64: png }),
      }),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(call('screenshot_range', { sheetId: 1, range: 'A1:B2' })),
    ).resolves.toMatchObject({
      mutated: false,
      display: { kind: 'images', items: [{ url: `data:image/png;base64,${png}` }] },
      modelContent: [{ type: 'image', image: { mime: 'image/png', base64: png } }],
    })
  })

  it('snapshots both source and destination for set_cell_range copyToRange', async () => {
    const fake = adapter()
    const proposals = createStructuredProposalController()
    await createExcelSkill({ adapter: fake, proposals }).executeTool(
      call('set_cell_range', {
        sheetId: 1,
        range: 'A1:B1',
        cells: [[{ value: 1 }, { value: 2 }]],
        copyToRange: 'D1:E1',
        allow_overwrite: true,
      }),
    )
    expect(proposals.pending()?.impact.targets).toEqual(['sheet:1!A1:B1', 'sheet:1!D1:E1'])
    expect(fake.fingerprint).toHaveBeenCalledWith(['sheet:1!A1:B1', 'sheet:1!D1:E1'], undefined)
  })

  it.each([
    [
      'styled range with note, border, copy, and autofit',
      'set_cell_range',
      {
        sheetId: 1,
        range: 'A1',
        cells: [
          [
            {
              value: 'x',
              note: 'reviewed',
              cellStyles: { fontWeight: 'bold', backgroundColor: '#ffeeaa' },
              borderStyles: { bottom: { style: 'double', weight: 'medium', color: '#112233' } },
            },
          ],
        ],
        copyToRange: 'B1',
        resizeWidth: { type: 'standard', value: 0 },
        resizeHeight: { type: 'standard', value: 0 },
        allow_overwrite: true,
      },
    ],
    ['clear formats', 'clear_cell_range', { sheetId: 1, range: 'A1', clearType: 'formats' }],
    ['clear all', 'clear_cell_range', { sheetId: 1, range: 'A1', clearType: 'all' }],
    ['copy range', 'copy_to', { sheetId: 1, sourceRange: 'A1:B2', destinationRange: 'D1:E2' }],
    [
      'insert rows',
      'modify_sheet_structure',
      { sheetId: 1, operation: 'insert', dimension: 'rows', reference: '5', count: 2 },
    ],
    [
      'delete columns',
      'modify_sheet_structure',
      { sheetId: 1, operation: 'delete', dimension: 'columns', reference: 'C', count: 2 },
    ],
    [
      'standard autofit',
      'resize_range',
      {
        sheetId: 1,
        range: 'A:D',
        width: { type: 'standard', value: 0 },
        height: { type: 'standard', value: 0 },
      },
    ],
    [
      'create chart',
      'modify_object',
      {
        operation: 'create',
        sheetId: 1,
        objectType: 'chart',
        properties: { name: 'Sales', source: 'A1:B4', chartType: 'line', title: 'Sales' },
      },
    ],
    [
      'update chart',
      'modify_object',
      {
        operation: 'update',
        sheetId: 1,
        objectType: 'chart',
        id: 'Sales',
        properties: { chartType: 'area', title: 'Updated' },
      },
    ],
    [
      'create pivot',
      'modify_object',
      {
        operation: 'create',
        sheetId: 1,
        objectType: 'pivotTable',
        properties: { name: 'Pivot', source: 'A1:B4', range: 'D1' },
      },
    ],
    [
      'update pivot',
      'modify_object',
      {
        operation: 'update',
        sheetId: 1,
        objectType: 'pivotTable',
        id: 'Pivot',
        properties: { name: 'Pivot2' },
      },
    ],
  ])('accepts documented %s payloads', async (_label, name, input) => {
    const proposals = createStructuredProposalController()
    await expect(
      createExcelSkill({ adapter: adapter(), proposals }).executeTool(call(name, input)),
    ).resolves.toMatchObject({ output: expect.stringContaining('proposalId'), mutated: false })
  })

  it('accepts every documented structured mutation shape and rejects incomplete/unknown nested input', async () => {
    const cases: Array<[string, Record<string, unknown>]> = [
      ['clear_cell_range', { sheetId: 1, range: 'A1', clearType: 'contents' }],
      [
        'modify_sheet_structure',
        {
          sheetId: 1,
          operation: 'hide',
          dimension: 'rows',
          reference: '5',
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
          operation: 'delete',
          sheetId: 1,
          objectType: 'chart',
          id: 'Chart 1',
        },
      ],
    ]
    for (const [name, input] of cases) {
      const proposals = createStructuredProposalController()
      const outcome = await createExcelSkill({ adapter: adapter(), proposals }).executeTool(
        call(name, input),
      )
      expect(outcome, name).toMatchObject({
        mutated: false,
        output: expect.stringContaining('proposalId'),
      })
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
    const sheet = {
      id: '1',
      name: 'Data',
      load: vi.fn(),
      getRange: vi.fn().mockReturnValue(range),
      getRangeByIndexes: vi.fn(() => {
        throw new Error('unsupported on this host')
      }),
    }
    const worksheets = {
      items: [sheet],
      load: vi.fn(),
      getItemAt: vi.fn(() => {
        throw new Error('non-standard API must not be called')
      }),
    }
    Object.assign(globalThis, {
      Office: {
        context: { host: 'Excel', requirements: { isSetSupported: vi.fn().mockReturnValue(true) } },
      },
      Excel: {
        run: (cb: (ctx: unknown) => unknown) => cb({ workbook: { worksheets }, sync: vi.fn() }),
      },
    })
    await expect(
      new BrowserExcelAdapter().getRangeAsCsv({ sheetId: 1, range: 'A1:B3', maxRows: 1 }),
    ).resolves.toMatchObject({ csv: 'h1,h2', rowCount: 1, hasMore: true })
    expect(sheet.getRange).toHaveBeenCalledWith('A1:B1')
    expect(sheet.getRangeByIndexes).not.toHaveBeenCalled()
    expect(worksheets.load).toHaveBeenCalledWith({ $skip: 0, $top: 1 })
    expect(worksheets.getItemAt).not.toHaveBeenCalled()
    expect(range.load).toHaveBeenCalledWith('values,rowCount,columnCount,address')
  })

  it('rejects malformed ranges before entering Excel.run', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: { host: 'Excel', requirements: { isSetSupported: vi.fn().mockReturnValue(true) } },
      },
      Excel: { run },
    })
    await expect(
      new BrowserExcelAdapter().getRangeAsCsv({ sheetId: 1, range: '*' }),
    ).rejects.toThrow('invalid_tool_input')
    expect(run).not.toHaveBeenCalled()
  })

  it('preserves bounded Office identifiers for diagnostics without exposing raw error text', async () => {
    const officeError = Object.assign(new Error('Workbook contains secret customer data'), {
      name: 'RichApi.Error',
      code: 'InvalidArgument',
      debugInfo: { errorLocation: 'Worksheet.getRange' },
    })
    const result = await createExcelSkill({
      adapter: adapter({ getCellRanges: vi.fn().mockRejectedValue(officeError) }),
      proposals: createStructuredProposalController(),
    }).executeTool(call('get_cell_ranges', { sheetId: 1, ranges: ['A1'] }))

    expect(result).toMatchObject({ output: 'office_read_failed', isError: true })
    expect(result).toHaveProperty('diagnosticError', officeError)
    expect(JSON.stringify(result)).not.toContain('secret customer data')
  })

  it('captures a range image using the exact ExcelApi gate and checks cancellation before sync', async () => {
    const result = { value: 'iVBORw0KGgoAAA==' }
    const range = {
      rowCount: 2,
      columnCount: 2,
      width: 100,
      height: 40,
      load: vi.fn(),
      getImage: vi.fn().mockReturnValue(result),
    }
    const isSetSupported = vi.fn().mockReturnValue(true)
    Object.assign(globalThis, {
      Office: { context: { host: 'Excel', requirements: { isSetSupported } } },
      Excel: {
        run: (cb: (ctx: unknown) => unknown) =>
          cb({
            workbook: { worksheets: { getItemAt: () => ({ getRange: () => range }) } },
            sync: vi.fn(),
          }),
      },
    })
    await expect(
      new BrowserExcelAdapter().screenshotRange({ sheetId: 1, range: 'A1:B2' }),
    ).resolves.toEqual({ mime: 'image/png', base64: result.value })
    expect(isSetSupported).toHaveBeenCalledWith('ExcelApi', '1.7')
    expect(range.getImage).toHaveBeenCalledOnce()
  })

  it('rejects oversized screenshot ranges before Excel.run and pixel-heavy ranges before getImage', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: { host: 'Excel', requirements: { isSetSupported: vi.fn().mockReturnValue(true) } },
      },
      Excel: { run },
    })
    await expect(
      new BrowserExcelAdapter().screenshotRange({ sheetId: 1, range: 'A1:XFD1048576' }),
    ).rejects.toThrow('invalid_tool_input')
    expect(run).not.toHaveBeenCalled()

    const getImage = vi.fn()
    const range = {
      width: 20_000,
      height: 20_000,
      rowCount: 2,
      columnCount: 2,
      load: vi.fn(),
      getImage,
    }
    Object.assign(globalThis, {
      Excel: {
        run: (cb: (ctx: unknown) => unknown) =>
          cb({
            workbook: { worksheets: { getItemAt: () => ({ getRange: () => range }) } },
            sync: vi.fn(),
          }),
      },
    })
    await expect(
      new BrowserExcelAdapter().screenshotRange({ sheetId: 1, range: 'A1:B2' }),
    ).rejects.toThrow('office_read_failed')
    expect(getImage).not.toHaveBeenCalled()
  })

  it('applies styles, borders, notes, copy and standard sizing behind the notes API gate', async () => {
    const border = {}
    const target = {
      values: undefined,
      formulas: undefined,
      numberFormat: undefined,
      load: vi.fn(),
      address: 'Data!A1',
      getResizedRange: vi.fn(),
      format: {
        font: {},
        fill: {},
        borders: { getItem: vi.fn().mockReturnValue(border) },
      },
    }
    const format = {
      autofitColumns: vi.fn(),
      autofitRows: vi.fn(),
    }
    const source = {
      format,
      getCell: vi.fn().mockReturnValue(target),
      getResizedRange: vi.fn(),
    }
    source.getResizedRange.mockReturnValue(source)
    target.getResizedRange.mockReturnValue(source)
    const destination = { format, copyFrom: vi.fn() }
    const notes = {
      add: vi.fn(),
      getItemOrNullObject: vi.fn().mockReturnValue({ isNullObject: true, load: vi.fn() }),
    }
    const worksheet = {
      notes,
      getRange: vi.fn((address: string) => (address === 'B1' ? destination : source)),
    }
    const isSetSupported = vi.fn().mockReturnValue(true)
    Object.assign(globalThis, {
      Office: { context: { host: 'Excel', requirements: { isSetSupported } } },
      Excel: {
        run: (cb: (ctx: unknown) => unknown) =>
          cb({
            workbook: { worksheets: { getItemAt: () => worksheet } },
            sync: vi.fn(),
          }),
      },
    })
    await new BrowserExcelAdapter().setCellRange({
      sheetId: 1,
      range: 'A1',
      cells: [
        [
          {
            value: 'x',
            note: 'note',
            cellStyles: { fontWeight: 'bold', backgroundColor: '#fff' },
            borderStyles: { bottom: { style: 'double', weight: 'medium' } },
          },
        ],
      ],
      copyToRange: 'B1',
      resizeWidth: { type: 'standard', value: 0 },
      resizeHeight: { type: 'standard', value: 0 },
      allow_overwrite: true,
    })
    expect(isSetSupported).toHaveBeenCalledWith('ExcelApi', '1.18')
    expect(target.values).toEqual([['x']])
    expect(target.format.font).toMatchObject({ bold: true })
    expect(target.format.fill).toMatchObject({ color: '#fff' })
    expect(border).toMatchObject({ style: 'Double', weight: 'Medium' })
    expect(notes.add).toHaveBeenCalledWith('A1', 'note')
    expect(destination.copyFrom).toHaveBeenCalledWith(source, 'All')
    expect(format.autofitColumns).toHaveBeenCalledOnce()
    expect(format.autofitRows).toHaveBeenCalledOnce()
  })

  it('cancels before entering Excel.run for every write family', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: { host: 'Excel', requirements: { isSetSupported: vi.fn().mockReturnValue(true) } },
      },
      Excel: { run },
    })
    const signal = AbortSignal.abort()
    const excel = new BrowserExcelAdapter()
    await expect(
      excel.copyTo({ sheetId: 1, sourceRange: 'A1', destinationRange: 'B1' }, signal),
    ).rejects.toThrow('cancelled')
    await expect(
      excel.modifyObject(
        { operation: 'delete', sheetId: 1, objectType: 'chart', id: 'Chart' },
        signal,
      ),
    ).rejects.toThrow('cancelled')
    expect(run).not.toHaveBeenCalled()
  })

  it('semantically verifies idempotent styles, clear, copy, autofit, and object updates', async () => {
    const excel = new BrowserExcelAdapter()
    vi.spyOn(excel, 'getCellRanges')
      .mockResolvedValueOnce({
        ranges: [
          {
            cells: [
              {
                value: 'x',
                formula: null,
                numberFormat: 'General',
                style: {
                  bold: true,
                  italic: false,
                  underline: 'None',
                  strikethrough: false,
                  backgroundColor: '#ffeeaa',
                  borders: [
                    { side: 'EdgeBottom', style: 'Double', weight: 'Medium', color: '#112233' },
                  ],
                },
              },
            ],
          },
        ],
      })
      .mockResolvedValueOnce({
        ranges: [
          {
            cells: [
              {
                value: '',
                formula: '',
                numberFormat: 'General',
                style: {
                  styleName: 'Normal',
                  bold: false,
                  italic: false,
                  underline: 'None',
                  strikethrough: false,
                  fontFamily: 'Aptos',
                  fontSize: 11,
                  fontColor: '#000000',
                  backgroundColor: '#000000',
                  fillPattern: 'None',
                  horizontalAlignment: 'General',
                  borders: [{ side: 'EdgeTop', style: 'None', weight: 'Thin', color: '#000000' }],
                },
              },
            ],
          },
        ],
      })
      .mockResolvedValueOnce({
        ranges: [{ cells: [{ value: 1, formula: null, numberFormat: 'General', style: {} }] }],
      })
      .mockResolvedValueOnce({
        ranges: [{ cells: [{ value: 1, formula: null, numberFormat: 'General', style: {} }] }],
      })
    await expect(
      excel.verifyMutation(
        'set_cell_range',
        {
          sheetId: 1,
          range: 'A1',
          cells: [
            [
              {
                value: 'x',
                cellStyles: { fontWeight: 'bold', backgroundColor: '#ffeeaa' },
                borderStyles: {
                  bottom: { style: 'double', weight: 'medium', color: '#112233' },
                },
              },
            ],
          ],
        },
        'same-fingerprint',
      ),
    ).resolves.toBe(true)
    await expect(
      excel.verifyMutation(
        'clear_cell_range',
        { sheetId: 1, range: 'A1', clearType: 'all' },
        'same-fingerprint',
      ),
    ).resolves.toBe(true)
    await expect(
      excel.verifyMutation(
        'copy_to',
        { sheetId: 1, sourceRange: 'A1', destinationRange: 'B1' },
        'same-fingerprint',
      ),
    ).resolves.toBe(true)
    vi.spyOn(excel, 'getCellRanges').mockResolvedValueOnce({
      ranges: [
        {
          cells: [
            {
              value: '',
              formula: '',
              numberFormat: 'General',
              style: {
                styleName: 'Normal',
                bold: false,
                italic: false,
                underline: 'None',
                strikethrough: false,
                fontFamily: 'Aptos',
                fontSize: 11,
                fontColor: '#000000',
                fillPattern: 'None',
                horizontalAlignment: 'General',
                borders: [{ side: 'EdgeBottom', style: 'Continuous' }],
              },
            },
          ],
        },
      ],
    })
    await expect(
      excel.verifyMutation(
        'clear_cell_range',
        { sheetId: 1, range: 'A1', clearType: 'formats' },
        'same-fingerprint',
      ),
    ).resolves.toBe(false)
    vi.spyOn(excel, 'getAllObjects').mockResolvedValue({
      objects: [{ type: 'chart', name: 'Sales', chartType: 'Area', title: 'Updated' }],
    })
    await expect(
      excel.verifyMutation(
        'modify_object',
        {
          sheetId: 1,
          operation: 'update',
          objectType: 'chart',
          id: 'Sales',
          properties: { chartType: 'area', title: 'Updated' },
        },
        'same-fingerprint',
      ),
    ).resolves.toBe(true)
  })

  it('verifies copied relative formulas by formulaR1C1 semantics', async () => {
    const excel = new BrowserExcelAdapter()
    vi.spyOn(excel, 'getCellRanges').mockResolvedValue({
      ranges: [
        {
          cells: [
            {
              address: 'A1',
              value: 2,
              formula: '=B1+1',
              formulaR1C1: '=RC[1]+1',
              numberFormat: 'General',
              style: {},
            },
          ],
        },
      ],
    })
    ;(excel.getCellRanges as any)
      .mockResolvedValueOnce({
        ranges: [
          {
            cells: [
              {
                address: 'A1',
                value: 2,
                formula: '=B1+1',
                formulaR1C1: '=RC[1]+1',
                numberFormat: 'General',
                style: {},
              },
            ],
          },
        ],
      })
      .mockResolvedValueOnce({
        ranges: [
          {
            cells: [
              {
                address: 'D1',
                value: 2,
                formula: '=E1+1',
                formulaR1C1: '=RC[1]+1',
                numberFormat: 'General',
                style: {},
              },
            ],
          },
        ],
      })
    await expect(
      excel.verifyMutation(
        'copy_to',
        { sheetId: 1, sourceRange: 'A1', destinationRange: 'D1' },
        'prestate',
      ),
    ).resolves.toBe(true)
  })

  it('verifies workbook deletion by stable worksheet identity rather than regenerated position', async () => {
    const excel = new BrowserExcelAdapter()
    vi.spyOn(excel, 'verifyWorkbook').mockResolvedValue({
      sheets: [
        { id: 1, officeId: 'stable-next-sheet', name: 'Next' },
        { id: 2, officeId: 'stable-last-sheet', name: 'Last' },
      ],
      hasMore: false,
    })
    await expect(
      excel.verifyMutation(
        'modify_workbook_structure',
        { operation: 'delete', sheetId: 1 },
        { officeId: 'deleted-stable-id', name: 'Deleted' },
      ),
    ).resolves.toBe(true)
    await expect(
      excel.verifyMutation(
        'modify_workbook_structure',
        { operation: 'delete', sheetId: 1 },
        { officeId: 'stable-next-sheet', name: 'Next' },
      ),
    ).resolves.toBe(false)
  })

  it('paginates search by a monotonic raw-cell cursor without rescanning matches', async () => {
    const matrix = [
      ['x', 'no'],
      ['x', 'x'],
      ['tail', 'x'],
    ]
    const sheet = {
      name: 'Data',
      load: vi.fn(),
      getRange: vi.fn((address: string) => {
        const match = /^([A-Z])(\d+)(?::([A-Z])(\d+))?$/.exec(address)!
        const row = Number(match[2]) - 1
        const column = match[1].charCodeAt(0) - 65
        const endRow = Number(match[4] ?? match[2]) - 1
        const endColumn = (match[3] ?? match[1]).charCodeAt(0) - 65
        const rows = endRow - row + 1
        const columns = endColumn - column + 1
        return {
          values: Array.from({ length: rows }, (_, r) =>
            matrix[row + r].slice(column, column + columns),
          ),
          formulas: Array.from({ length: rows }, () => Array.from({ length: columns }, () => null)),
          load: vi.fn(),
        }
      }),
    }
    Object.assign(globalThis, {
      Office: {
        context: { host: 'Excel', requirements: { isSetSupported: vi.fn().mockReturnValue(true) } },
      },
      Excel: {
        run: (cb: (ctx: unknown) => unknown) =>
          cb({ workbook: { worksheets: { getItemAt: () => sheet } }, sync: vi.fn() }),
      },
    })
    const adapter = new BrowserExcelAdapter()
    const first = (await adapter.searchData({
      searchTerm: 'x',
      sheetId: 1,
      range: 'A1:B3',
      options: { maxResults: 1 },
    })) as { nextOffset: number; matches: Array<{ address: string }> }
    expect(first).toMatchObject({ nextOffset: 1, matches: [{ address: 'A1' }] })
    const second = (await adapter.searchData({
      searchTerm: 'x',
      sheetId: 1,
      range: 'A1:B3',
      offset: first.nextOffset,
      options: { maxResults: 1 },
    })) as { nextOffset: number; matches: Array<{ address: string }> }
    expect(second).toMatchObject({ nextOffset: 3, matches: [{ address: 'A2' }] })
  })
})
