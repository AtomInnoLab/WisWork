import { describe, expect, it, vi } from 'vitest'
import { createBrowserWordElevatedAdapter } from '../src/skills/word/elevated-word-adapter.js'
import { createBrowserExcelElevatedAdapter } from '../src/skills/excel/elevated-excel-adapter.js'
import { createBrowserPowerPointElevatedAdapter } from '../src/skills/powerpoint/elevated-powerpoint-adapter.js'

const authority = () => ({
  activeMode: 'enhanced' as const,
  signedIn: true,
  paired: true,
  hostEnabled: true,
  rawOfficeEnabled: true,
  rawOfficeJsEnabled: true,
  rawOfficeOoxmlEnabled: true,
  documentId: 'document_AAAAAAAAAAAAAAAA',
  sessionId: 'session_AAAAAAAAAAAAAAAA',
  generation: 1,
  revision: 'revision_AAAAAAAAAAAAAAAA',
})

describe('production elevated Office adapters', () => {
  it('Word compiles the closed AST into the existing Word transaction authority', async () => {
    const adapter = {
      getDocumentSnapshot: vi.fn(async () => ({ text: 'before', fingerprint: 'before-fp' })),
      getOoxml: vi.fn(async () => ({ xml: '<w:document/>', children: [] })),
      fingerprint: vi.fn(async () => 'before-fp'),
      executeOperations: vi.fn(async () => undefined),
      verifyOperations: vi.fn(async () => true),
    } as any
    const elevated = createBrowserWordElevatedAdapter({ adapter, authority })
    const program = {
      version: 1 as const,
      kind: 'office_js_ast' as const,
      operations: [{ call: 'body.insertText', args: { location: 'end', text: 'hello' } }],
    }
    const snapshot = await elevated.snapshot(program)
    expect(await elevated.validateSnapshot(program, snapshot)).toBe(true)
    await elevated.execute(program, snapshot)
    expect(adapter.executeOperations).toHaveBeenCalledWith(
      [{ op: 'insert_text', location: 'end', text: 'hello' }],
      undefined,
    )
    expect(await elevated.readback(program, snapshot)).toMatchObject({ verified: true })
    const sync = vi.fn(async () => undefined)
    ;(globalThis as any).Word = {
      run: async (run: (context: any) => Promise<void>) =>
        run({ document: { body: { insertOoxml: vi.fn() } }, sync }),
    }
    await elevated.rollback(snapshot)
    expect(sync).toHaveBeenCalledOnce()
    delete (globalThis as any).Word
  })

  it('Excel compiles values into the existing capture/write/verify/recovery authority and has no OOXML path', async () => {
    const adapter = {
      fingerprint: vi.fn(async () => 'excel-fp'),
      captureMutation: vi.fn(async () => ({ values: [['before']] })),
      setCellRange: vi.fn(async () => undefined),
      clearCellRange: vi.fn(async () => undefined),
      verifyMutation: vi.fn(async () => true),
      recoverMutation: vi.fn(async () => 'restored'),
    } as any
    const elevated = createBrowserExcelElevatedAdapter({ adapter, authority })
    const program = {
      version: 1 as const,
      kind: 'office_js_ast' as const,
      operations: [
        { call: 'range.setValues', args: { sheetId: 1, range: 'A1', values: [['after']] } },
      ],
    }
    const snapshot = await elevated.snapshot(program)
    await elevated.execute(program, snapshot)
    expect(adapter.setCellRange).toHaveBeenCalledWith(
      expect.objectContaining({
        sheetId: 1,
        range: 'A1',
        cells: [[{ value: 'after' }]],
        allow_overwrite: true,
      }),
      undefined,
    )
    expect(await elevated.readback(program, snapshot)).toMatchObject({ verified: true })
    await elevated.rollback(snapshot)
    expect(adapter.recoverMutation).toHaveBeenCalledTimes(1)
  })

  it('PowerPoint compiles shape text and proves it with host readback', async () => {
    const adapter = {
      exportSlidePackage: vi.fn(async () => ({
        slideId: 'slide-1',
        fingerprint: 'ppt-fp',
        base64: 'UEs=',
      })),
      snapshotSlide: vi.fn(async () => ({ slideId: 'slide-1', fingerprint: 'ppt-fp' })),
      executeDeclarative: vi.fn(async () => ({ createdShapeIds: [] })),
      readSlideText: vi.fn(async () => ({
        slideId: 'slide-1',
        shapeId: 'shape-1',
        text: 'after',
        paragraphs: ['after'],
      })),
      listSlideShapes: vi.fn(),
      replaceSlidePackage: vi.fn(),
    } as any
    const elevated = createBrowserPowerPointElevatedAdapter({ adapter, authority })
    const program = {
      version: 1 as const,
      kind: 'office_js_ast' as const,
      operations: [
        { call: 'shape.setText', args: { slideIndex: 0, shapeId: 'shape-1', text: 'after' } },
      ],
    }
    const snapshot = await elevated.snapshot(program)
    expect(await elevated.validateSnapshot(program, snapshot)).toBe(true)
    await elevated.execute(program, snapshot)
    expect(adapter.executeDeclarative).toHaveBeenCalledWith(
      [{ op: 'set_shape_text', slide_index: 0, shape_id: 'shape-1', text: 'after' }],
      undefined,
    )
    expect(await elevated.readback(program, snapshot)).toMatchObject({ verified: true })
    await elevated.rollback(snapshot)
    expect(adapter.replaceSlidePackage).toHaveBeenCalledWith(0, 'UEs=', false, undefined, undefined)
  })

  it('PowerPoint proves a new text box by its generated shape id, text, and geometry', async () => {
    const adapter = {
      exportSlidePackage: vi.fn(async () => ({
        slideId: 'slide-1',
        fingerprint: 'ppt-fp',
        base64: 'UEs=',
      })),
      snapshotSlide: vi.fn(async () => ({ slideId: 'slide-1', fingerprint: 'ppt-fp' })),
      executeDeclarative: vi.fn(async () => ({ createdShapeIds: ['created-1'] })),
      readSlideText: vi.fn(async () => ({
        slideId: 'slide-1',
        shapeId: 'created-1',
        text: 'hello',
        paragraphs: ['hello'],
      })),
      listSlideShapes: vi.fn(async () => ({
        slideId: 'slide-1',
        slideIndex: 0,
        shapes: [
          {
            id: 'created-1',
            name: 'WisWork Raw Text',
            type: 'TextBox',
            left: 1,
            top: 2,
            width: 3,
            height: 4,
          },
        ],
      })),
      readShapeTextStyle: vi.fn(async () => ({
        color: '#112233',
        fontFamily: 'Aptos',
        fontSize: 18,
        bold: false,
        italic: false,
      })),
      replaceSlidePackage: vi.fn(),
    } as any
    const elevated = createBrowserPowerPointElevatedAdapter({ adapter, authority })
    const program = {
      version: 1 as const,
      kind: 'office_js_ast' as const,
      operations: [
        {
          call: 'slide.addTextBox',
          args: {
            slideIndex: 0,
            text: 'hello',
            left: 1,
            top: 2,
            width: 3,
            height: 4,
            style: {
              color: '#112233',
              fontFamily: 'Aptos',
              fontSize: 18,
              bold: false,
              italic: false,
            },
          },
        },
      ],
    }
    const snapshot = await elevated.snapshot(program)
    await elevated.execute(program, snapshot)
    expect(await elevated.readback(program, snapshot)).toMatchObject({ verified: true })
    expect(adapter.readSlideText).toHaveBeenCalledWith(0, 'created-1', undefined)
  })
})
