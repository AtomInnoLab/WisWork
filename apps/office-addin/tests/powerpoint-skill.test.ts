import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  BrowserPowerPointAdapter,
  type PowerPointAdapter,
} from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import { benchmarkDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-benchmark'
import {
  editPowerPointPackage,
  presentationPackageDigest,
} from '../src/skills/powerpoint/powerpoint-package.js'

const png = 'iVBORw0KGgoAAAA='

function adapter(overrides: Partial<PowerPointAdapter> = {}): PowerPointAdapter {
  return {
    inspectSlideMasters: vi.fn().mockResolvedValue({
      masters: [
        {
          id: 'master-1',
          name: 'Main',
          background: { type: 'Solid', color: '#FFFFFF', transparency: 0 },
          themeColors: { Light1: '#FFFFFF', Dark1: '#000000' },
          layouts: [
            {
              id: 'layout-1',
              name: 'Title',
              isMasterBackgroundFollowed: true,
              areBackgroundGraphicsHidden: false,
              background: { type: 'Solid' },
            },
          ],
        },
      ],
    }),
    executeMasterOperations: vi.fn().mockResolvedValue(undefined),
    screenshotSlide: vi
      .fn()
      .mockResolvedValue({ slideId: 'host-slide-1', mime: 'image/png', base64: png }),
    listSlideShapes: vi.fn().mockResolvedValue({
      slideId: 'slide-1',
      slideIndex: 0,
      shapes: [
        { id: '2', name: 'Title', type: 'TextBox', left: 10, top: 20, width: 200, height: 40 },
      ],
    }),
    readSlideText: vi.fn().mockResolvedValue({
      slideId: 'slide-1',
      shapeId: '2',
      text: 'Hello',
      paragraphs: ['Hello'],
    }),
    readSlideTable: vi.fn().mockResolvedValue([
      ['方案', '结果'],
      ['甲', '120'],
    ]),
    verifySlides: vi.fn().mockResolvedValue({
      slideWidth: 960,
      slideHeight: 540,
      slides: [],
    }),
    snapshotSlide: vi.fn().mockResolvedValue({ slideId: 'slide-1', fingerprint: 'slide-1:1' }),
    editSlideText: vi.fn().mockResolvedValue(undefined),
    duplicateSlide: vi.fn().mockResolvedValue({ slideId: 'slide-copy' }),
    exportSlidePackage: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    replaceSlidePackage: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    executeDeclarative: vi.fn().mockRejectedValue(new Error('office_api_unsupported')),
    ...overrides,
  }
}

const call = (name: string, input: Record<string, unknown> = {}) => ({ id: 'call-1', name, input })

describe('PowerPoint compatibility skill', () => {
  it('exposes native master inspection and editing with exact schemas', () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    expect(skill.tools.map((tool) => tool.name)).toEqual([
      'inspect_slide_masters',
      'screenshot_slide',
      'list_slide_shapes',
      'read_slide_text',
      'verify_slides',
      'execute_office_js',
      'add_slide_ir_objects',
      'edit_slide_text',
      'edit_slide_xml',
      'edit_slide_chart',
      'edit_slide_master',
      'edit_slide_master_xml',
      'duplicate_slide',
    ])
    for (const tool of skill.tools) expect(tool.inputSchema.additionalProperties).toBe(false)
    expect(skill.tools.find((tool) => tool.name === 'edit_slide_text')?.inputSchema).toMatchObject({
      required: ['slide_index', 'shape_id', 'text'],
      properties: {
        slide_index: { type: 'integer', minimum: 0, maximum: 100000 },
        shape_id: { type: 'string', minLength: 1, maxLength: 256 },
        text: { type: 'string', maxLength: 12000 },
      },
    })
    for (const name of [
      'execute_office_js',
      'edit_slide_xml',
      'edit_slide_chart',
      'edit_slide_master',
      'edit_slide_master_xml',
    ]) {
      const schema = skill.tools.find((tool) => tool.name === name)?.inputSchema
      expect(schema?.properties).toHaveProperty('program')
      expect(schema?.properties).not.toHaveProperty('code')
      expect(schema?.required).toContain('program')
    }
    const executeSchema = skill.tools.find((tool) => tool.name === 'execute_office_js')
      ?.inputSchema as unknown as {
      properties: { program: { properties: { operations: { items: unknown } } } }
    }
    const programSchema = executeSchema.properties.program
    expect(programSchema.properties.operations.items).toHaveProperty('anyOf')
  })

  it('proposes one native master edit and verifies semantic readback', async () => {
    let state = await adapter().inspectSlideMasters()
    const fake = adapter({
      inspectSlideMasters: vi.fn().mockImplementation(() => Promise.resolve(state)),
      executeMasterOperations: vi.fn().mockImplementation(async () => {
        state = {
          masters: [
            {
              ...state.masters[0],
              background: { type: 'Solid', color: '#000000', transparency: 0 },
            },
          ],
        }
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await expect(
      skill.executeTool(
        call('edit_slide_master', {
          program: {
            version: 2,
            operations: [
              {
                op: 'set_master_background',
                master_id: 'master-1',
                fill: { type: 'solid', color: '#000000', transparency: 0 },
              },
            ],
          },
        }),
      ),
    ).resolves.toMatchObject({ mutated: false, summary: 'Proposed native PowerPoint master edit' })
    expect(proposals.pending()?.impact).toMatchObject({ count: 1, targets: ['master:master-1'] })
    await proposals.confirm(proposals.pending()!.id)
    expect(fake.executeMasterOperations).toHaveBeenCalledOnce()
  })

  it('derives native QA scope and blocks dependency drift after persistence', async () => {
    let slides = [{ slideId: 's1', masterId: 'master-1', layoutId: 'layout-1' }]
    const fake = adapter({ inspectStyleDependencies: vi.fn(async () => ({ slides })) })
    const proposals = createStructuredProposalController(undefined, {
      beforeWrite: async () => {
        slides = [...slides, { slideId: 's2', masterId: 'master-1', layoutId: 'layout-1' }]
      },
      afterWrite: vi.fn(),
    })
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('edit_slide_master', {
        program: {
          version: 2,
          operations: [
            {
              op: 'set_master_theme_color',
              master_id: 'master-1',
              theme_color: 'Dark1',
              color: '#112233',
            },
          ],
        },
      }),
    )
    expect(proposals.pending()?.preview).toMatchObject({
      qaScope: { basis: 'native_master_layout', hostSlideIds: ['s1'] },
    })
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.executeMasterOperations).not.toHaveBeenCalled()
  })

  it.each(['unsupported', 'invalid', 'empty', 'cancelled'])(
    'handles %s initial dependencies conservatively',
    async (mode) => {
      const inspectStyleDependencies = vi.fn(async () => {
        if (mode === 'unsupported') throw new Error('office_api_unsupported')
        if (mode === 'cancelled') throw new Error('cancelled')
        return mode === 'invalid' ? { slides: [{ slideId: 'x' }] } : { slides: [] }
      })
      const proposals = createStructuredProposalController()
      const fake = adapter({
        inspectStyleDependencies:
          inspectStyleDependencies as unknown as PowerPointAdapter['inspectStyleDependencies'],
      })
      const result = await createPowerPointSkill({ adapter: fake, proposals }).executeTool(
        call('edit_slide_master', {
          program: {
            version: 2,
            operations: [
              {
                op: 'set_master_theme_color',
                master_id: 'master-1',
                theme_color: 'Dark1',
                color: '#112233',
              },
            ],
          },
        }),
      )
      if (mode === 'cancelled') {
        expect(result).toMatchObject({ isError: true, output: 'cancelled' })
        expect(proposals.pending()).toBeUndefined()
      } else
        expect(proposals.pending()?.preview).toMatchObject({
          qaScope:
            mode === 'empty'
              ? { basis: 'native_master_layout', hostSlideIds: [] }
              : { basis: 'document' },
        })
    },
  )

  it.each(['drift', 'read failure'])('blocks %s at first confirmation validation', async (mode) => {
    const inspectStyleDependencies = vi.fn().mockResolvedValueOnce({
      slides: [{ slideId: 's1', masterId: 'master-1', layoutId: 'layout-1' }],
    })
    if (mode === 'drift') inspectStyleDependencies.mockResolvedValue({ slides: [] })
    else inspectStyleDependencies.mockRejectedValue(new Error('office_read_failed'))
    const fake = adapter({ inspectStyleDependencies })
    const beforeWrite = vi.fn()
    const proposals = createStructuredProposalController(undefined, {
      beforeWrite,
      afterWrite: vi.fn(),
    })
    await createPowerPointSkill({ adapter: fake, proposals }).executeTool(
      call('edit_slide_master', {
        program: {
          version: 2,
          operations: [
            {
              op: 'set_master_theme_color',
              master_id: 'master-1',
              theme_color: 'Dark1',
              color: '#112233',
            },
          ],
        },
      }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow()
    expect(beforeWrite).not.toHaveBeenCalled()
    expect(fake.executeMasterOperations).not.toHaveBeenCalled()
  })

  it('recovers an already verified native master operation when a later operation fails', async () => {
    const original = await adapter().inspectSlideMasters()
    const state = structuredClone(original)
    let writes = 0
    const fake = adapter({
      inspectSlideMasters: vi
        .fn()
        .mockImplementation(() => Promise.resolve(structuredClone(state))),
      executeMasterOperations: vi.fn().mockImplementation(async (operations) => {
        writes += 1
        const operation = operations[0]
        if (writes === 2) throw new Error('office_write_failed')
        if (operation.op === 'set_master_background' && operation.fill.type === 'solid')
          state.masters[0]!.background = {
            type: 'Solid',
            color: operation.fill.color,
            transparency: operation.fill.transparency,
          }
        if (operation.op === 'set_master_theme_color')
          state.masters[0]!.themeColors[operation.theme_color] = operation.color
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('edit_slide_master', {
        program: {
          version: 2,
          operations: [
            {
              op: 'set_master_background',
              master_id: 'master-1',
              fill: { type: 'solid', color: '#000000', transparency: 0 },
            },
            {
              op: 'set_master_theme_color',
              master_id: 'master-1',
              theme_color: 'Light1',
              color: '#EEEEEE',
            },
          ],
        },
      }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('office_write_failed')
    expect(state).toEqual(original)
    expect(fake.executeMasterOperations).toHaveBeenCalledTimes(3)
  })

  it('does not advertise or execute master package edits on PowerPoint for Mac', async () => {
    const fake = adapter()
    const skill = createPowerPointSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
      platform: 'Mac',
    })

    expect(skill.tools.map((tool) => tool.name)).toContain('edit_slide_master')
    expect(skill.tools.map((tool) => tool.name)).not.toContain('edit_slide_master_xml')
    expect(skill.systemPrompt).toContain('inspect_slide_masters')
    await expect(
      skill.executeTool(
        call('edit_slide_master_xml', {
          program: {
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slideMasters/slideMaster1.xml',
                xml: '<p:sldMaster xmlns:p="urn:p"/>',
              },
            ],
          },
        }),
      ),
    ).resolves.toMatchObject({ isError: true, output: 'office_api_unsupported' })
    expect(fake.exportSlidePackage).not.toHaveBeenCalled()
    expect(fake.replaceSlidePackage).not.toHaveBeenCalled()
  })

  it('normalizes reads, image display, and rejects unknown fields', async () => {
    const fake = adapter()
    const skill = createPowerPointSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(call('list_slide_shapes', { slide_index: 0 })),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('"id":"2"'),
    })
    await expect(
      skill.executeTool(call('read_slide_text', { slide_index: 0, shape_id: '2' })),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('Hello'),
    })
    await expect(
      skill.executeTool(call('screenshot_slide', { slide_index: 0 })),
    ).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('"visualAvailableToModel":true'),
      modelContent: [{ type: 'image', image: { mime: 'image/png', base64: png } }],
      display: { kind: 'images', items: [{ url: `data:image/png;base64,${png}` }] },
    })
    expect(
      JSON.parse((await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))).output),
    ).toMatchObject({ slideId: 'host-slide-1', slideIndex: 0 })
    await expect(skill.executeTool(call('verify_slides', { nope: true }))).resolves.toMatchObject({
      output: 'invalid_tool_input',
      isError: true,
    })
  })
  it('uses a labeled fallback screenshot and keeps a missing image in a waiting state', async () => {
    const failure = Object.assign(new Error('office_read_failed'), {
      code: 'office_screenshot_unavailable',
    })
    const fake = adapter({ screenshotSlide: vi.fn().mockRejectedValue(failure) })
    const screenshotFallback = vi.fn().mockResolvedValue({
      slideId: 'host-slide-1',
      mime: 'image/png',
      base64: png,
      renderer: 'libreoffice',
    })
    const skill = createPowerPointSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
      screenshotFallback,
    })
    const result = await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    expect(JSON.parse(result.output)).toMatchObject({
      slideId: 'host-slide-1',
      slideIndex: 0,
      renderer: 'libreoffice',
      visualAvailableToModel: true,
    })
    expect(screenshotFallback).toHaveBeenCalledWith(0, undefined, undefined)
    const boundFailure = Object.assign(new Error('busy'), {
      code: 'Timeout',
      targetSlideId: 'host-slide-1',
    })
    ;(fake.screenshotSlide as ReturnType<typeof vi.fn>).mockRejectedValueOnce(boundFailure)
    await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    expect(screenshotFallback).toHaveBeenLastCalledWith(0, undefined, 'host-slide-1')
    screenshotFallback.mockResolvedValueOnce({
      slideId: 'different-page',
      mime: 'image/png',
      base64: png,
    })
    ;(fake.screenshotSlide as ReturnType<typeof vi.fn>).mockRejectedValueOnce(boundFailure)
    await expect(
      skill.executeTool(call('screenshot_slide', { slide_index: 0 })),
    ).resolves.toMatchObject({
      isError: true,
      output: 'office_concurrent_change',
    })
    screenshotFallback.mockRejectedValue(new Error('renderer_unavailable'))
    const waiting = await skill.executeTool(call('screenshot_slide', { slide_index: 0 }))
    expect(JSON.parse(waiting.output)).toEqual({ status: 'waiting_screenshot', slideIndex: 0 })
    expect(waiting.modelContent).toBeUndefined()
  })

  it('gates text edits behind immutable stale-checked proposals and verifies after confirmation', async () => {
    const fake = adapter({
      readSlideText: vi
        .fn()
        .mockResolvedValueOnce({
          slideId: 'slide-1',
          shapeId: '2',
          text: 'Hello',
          paragraphs: ['Hello'],
        })
        .mockResolvedValueOnce({
          slideId: 'slide-1',
          shapeId: '2',
          text: 'Hello',
          paragraphs: ['Hello'],
        })
        .mockResolvedValue({
          slideId: 'slide-1',
          shapeId: '2',
          text: 'New',
          paragraphs: ['New'],
        }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const proposed = await skill.executeTool(
      call('edit_slide_text', {
        slide_index: 0,
        shape_id: '2',
        text: 'New',
        explanation: 'Update title',
      }),
    )
    expect(proposed).toMatchObject({ mutated: false, summary: 'Proposed PowerPoint text edit' })
    const pending = proposals.pending()!
    expect(pending).toMatchObject({
      toolName: 'edit_slide_text',
      preview: { shapeId: '2', before: 'Hello', after: 'New' },
      impact: { host: 'powerpoint', targets: ['slide-1/2'], count: 1 },
    })
    await proposals.confirm(pending.id)
    expect(fake.editSlideText).toHaveBeenCalledWith(0, '2', 'New', expect.any(AbortSignal))
    expect(fake.snapshotSlide).not.toHaveBeenCalled()
    expect(fake.verifySlides).toHaveBeenCalledOnce()
  })

  it('refuses stale or cancelled writes before mutation', async () => {
    const fake = adapter({
      listSlideShapes: vi.fn().mockImplementation((index: number) =>
        Promise.resolve({
          slideId: index === 1 ? 'slide-copy' : 'slide-1',
          slideIndex: index,
          shapes: [],
        }),
      ),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(call('duplicate_slide', { slide_index: 0 }))
    ;(fake.snapshotSlide as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      slideId: 'slide-1',
      fingerprint: 'changed',
    })
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.duplicateSlide).not.toHaveBeenCalled()

    const controller = new AbortController()
    controller.abort()
    await expect(
      skill.executeTool(
        call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'x' }),
        controller.signal,
      ),
    ).resolves.toMatchObject({
      output: 'cancelled',
      isError: true,
    })
  })

  it('rejects JavaScript syntax and unknown declarative authority without proposals', async () => {
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: adapter(), proposals })
    for (const input of [
      { code: 'return context.presentation' },
      { code: '{"version":1,"operations":[{"op":"fetch","url":"https://x"}]}' },
    ]) {
      await expect(skill.executeTool(call('execute_office_js', input))).resolves.toMatchObject({
        output: 'invalid_tool_input',
        isError: true,
        mutated: false,
      })
      expect(proposals.pending()).toBeUndefined()
    }
  })

  it('accepts direct structured programs without JSON string double encoding', async () => {
    const fake = adapter()
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const program = {
      version: 1,
      operations: [{ op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'Structured' }],
    }

    await expect(skill.executeTool(call('execute_office_js', { program }))).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('set_shape_text'),
    })
    expect(proposals.pending()?.preview).toEqual(program)
  })

  it('reports a content-free program location for invalid operation fields', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(
        call('execute_office_js', {
          program: { version: 1, operations: [{ op: 'add_text_box', slide_index: 0 }] },
        }),
      ),
    ).resolves.toMatchObject({
      output: 'invalid_tool_input',
      isError: true,
      diagnosticError: {
        code: 'InvalidToolInput',
        debugInfo: { errorLocation: 'program.operations' },
      },
    })
  })

  it('reports a content-free program location for malformed streamed tool JSON', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool({
        id: 'call-1',
        name: 'execute_office_js',
        input: {},
        inputError: 'raw malformed JSON must not be retained',
      }),
    ).resolves.toMatchObject({
      output: 'invalid_tool_input',
      isError: true,
      diagnosticError: {
        code: 'InvalidToolInput',
        debugInfo: { errorLocation: 'program' },
      },
    })
  })

  it('accepts direct structured XML programs for slide and master edits', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    const base64 = await zip.generateAsync({ type: 'base64' })
    const fake = adapter({
      exportSlidePackage: vi.fn().mockResolvedValue({
        slideId: 's1',
        base64,
        fingerprint: 'stable',
      }),
    })
    for (const [name, input] of [
      [
        'edit_slide_xml',
        {
          slide_index: 0,
          program: {
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slides/slide1.xml',
                xml: '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>',
              },
            ],
          },
        },
      ],
      [
        'edit_slide_master_xml',
        {
          program: {
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slideMasters/slideMaster1.xml',
                xml: '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
              },
            ],
          },
        },
      ],
    ] as const) {
      const controller = createStructuredProposalController()
      const scoped = createPowerPointSkill({ adapter: fake, proposals: controller })
      await expect(scoped.executeTool(call(name, input))).resolves.toMatchObject({
        mutated: false,
        summary: expect.stringContaining('Proposed'),
      })
      controller.reject()
    }
  })

  it('proposes and semantically verifies bounded XML, chart, and master package edits', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    zip.file('ppt/charts/chart1.xml', '<c:chart xmlns:c="urn:c"/>')
    zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    let current = await zip.generateAsync({ type: 'base64' })
    const fake = adapter({
      exportSlidePackage: vi.fn().mockImplementation(() =>
        Promise.resolve({
          slideId: 's1',
          base64: current,
          fingerprint: `${current.length}:${current.slice(-8)}`,
        }),
      ),
      replaceSlidePackage: vi.fn().mockImplementation((_index, base64) => {
        current = base64
        return Promise.resolve({ slideId: 's2' })
      }),
    })
    for (const [name, input] of [
      [
        'edit_slide_xml',
        {
          slide_index: 0,
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slides/slide1.xml',
                xml: '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>',
              },
            ],
          }),
        },
      ],
      [
        'edit_slide_chart',
        {
          slide_index: 0,
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/charts/chart1.xml',
                xml: '<c:chart xmlns:c="urn:c"><c:title/></c:chart>',
              },
            ],
          }),
        },
      ],
      [
        'edit_slide_master_xml',
        {
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'replace_xml',
                path: 'ppt/slideMasters/slideMaster1.xml',
                xml: '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
              },
            ],
          }),
        },
      ],
    ] as const) {
      const proposals = createStructuredProposalController()
      const skill = createPowerPointSkill({ adapter: fake, proposals })
      await expect(skill.executeTool(call(name, input))).resolves.toMatchObject({
        mutated: false,
        summary: expect.stringContaining('Proposed'),
      })
      await proposals.confirm(proposals.pending()!.id)
    }
    expect(fake.replaceSlidePackage).toHaveBeenCalledTimes(3)
  })

  it('confirms one synchronized chart value edit and rejects package drift', async () => {
    const book = new JSZip()
    book.file(
      'xl/workbook.xml',
      '<workbook><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>',
    )
    book.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
    )
    book.file(
      'xl/worksheets/sheet1.xml',
      '<worksheet><sheetData><row r="2"><c r="A2" t="inlineStr"><is><t>Q1</t></is></c><c r="B2"><v>1</v></c></row></sheetData></worksheet>',
    )
    const zip = new JSZip()
    zip.file(
      'ppt/slides/slide1.xml',
      '<p:sld><p:graphicFrame><p:cNvPr id="8"/><c:chart r:id="rId5"/></p:graphicFrame></p:sld>',
    )
    zip.file(
      'ppt/slides/_rels/slide1.xml.rels',
      '<Relationships><Relationship Id="rId5" Type="x/chart" Target="../charts/chart1.xml"/></Relationships>',
    )
    zip.file(
      'ppt/charts/chart1.xml',
      '<c:chartSpace><c:chart><c:plotArea><c:barChart><c:ser><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$2</c:f><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$2</c:f><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart><c:externalData r:id="rId9"/></c:chartSpace>',
    )
    zip.file(
      'ppt/charts/_rels/chart1.xml.rels',
      '<Relationships><Relationship Id="rId9" Type="x/package" Target="../embeddings/Book1.xlsx"/></Relationships>',
    )
    zip.file('ppt/embeddings/Book1.xlsx', await book.generateAsync({ type: 'uint8array' }))
    const original = await zip.generateAsync({ type: 'base64' })
    let current = original
    let activeSlideId = 's1'
    let savedBytes = new Uint8Array(0)
    let meta: Record<string, unknown> = {}
    const records = new Map<
      string,
      import('../src/skills/powerpoint/presentation-existing-chart.js').PresentationExistingChartChange
    >()
    const chartSavepoint = {
      documentId: async () => 'doc-1',
      readExistingChartChange: (id: string) => records.get(id),
      writeExistingChartChange: async (
        record: import('../src/skills/powerpoint/presentation-existing-chart.js').PresentationExistingChartChange,
      ) => {
        records.set(record.changeId, structuredClone(record))
      },
      request: async (body: unknown) => {
        const input = body as Record<string, unknown>
        if (input.operation === 'existing_page_backup_begin')
          meta = { ...input, status: 'uploading', receivedBytes: 0 }
        if (input.operation === 'existing_page_backup_chunk') {
          const bytes = Uint8Array.from(atob(input.base64 as string), (char) => char.charCodeAt(0))
          const next = new Uint8Array(savedBytes.length + bytes.length)
          next.set(savedBytes)
          next.set(bytes, savedBytes.length)
          savedBytes = next
          meta.receivedBytes = savedBytes.length
        }
        if (input.operation === 'existing_page_backup_finish') meta.status = 'ready'
        if (input.operation === 'existing_page_backup_release') {
          savedBytes = new Uint8Array(0)
          return new Response(JSON.stringify({ ...input, status: 'released' }))
        }
        if (input.operation === 'existing_page_backup_read') {
          const part = savedBytes.subarray(
            input.offset as number,
            (input.offset as number) + (input.length as number),
          )
          return new Response(
            JSON.stringify({
              backupId: meta.backupId,
              offset: input.offset,
              sizeBytes: meta.sizeBytes,
              sha256: meta.sha256,
              base64: btoa(String.fromCharCode(...part)),
            }),
          )
        }
        return new Response(JSON.stringify(meta))
      },
    }
    const fake = adapter({
      verifySlides: vi.fn().mockImplementation(() =>
        Promise.resolve({
          slideWidth: 960,
          slideHeight: 540,
          slides: [{ slideId: activeSlideId }],
        }),
      ),
      exportSlidePackage: vi
        .fn()
        .mockImplementation(() =>
          Promise.resolve({ slideId: activeSlideId, base64: current, fingerprint: activeSlideId }),
        ),
      replaceSlidePackage: vi.fn().mockImplementation((_index, base64) => {
        current = base64
        return Promise.resolve({ slideId: activeSlideId })
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals, chartSavepoint })
    await expect(
      skill.executeTool(
        call('update_slide_chart_values', { slide_index: 0, shape_id: '8', values: [['5']] }),
      ),
    ).resolves.toMatchObject({ mutated: false })
    await proposals.confirm(proposals.pending()!.id)
    expect(fake.replaceSlidePackage).toHaveBeenCalledOnce()
    const updated = await JSZip.loadAsync(current, { base64: true })
    expect(await updated.file('ppt/charts/chart1.xml')!.async('string')).toContain('<c:v>5</c:v>')
    const updatedBook = await JSZip.loadAsync(
      await updated.file('ppt/embeddings/Book1.xlsx')!.async('uint8array'),
    )
    expect(await updatedBook.file('xl/worksheets/sheet1.xml')!.async('string')).toContain(
      '<v>5</v>',
    )
    const changeId = [...records.keys()][0]!
    expect(records.get(changeId)?.state).toBe('applied')
    const appliedRecord = structuredClone(records.get(changeId)!)
    await expect(
      skill.executeTool(call('release_slide_chart_values_change', { change_id: changeId })),
    ).resolves.toMatchObject({ isError: true })
    const reopened = createPowerPointSkill({ adapter: fake, proposals, chartSavepoint })
    await expect(
      reopened.executeTool(call('inspect_slide_chart_values_change', { change_id: changeId })),
    ).resolves.toMatchObject({ mutated: false })
    await expect(
      reopened.executeTool(call('undo_slide_chart_values_change', { change_id: changeId })),
    ).resolves.toMatchObject({ mutated: false })
    await proposals.confirm(proposals.pending()!.id)
    expect(records.get(changeId)?.state).toBe('undone')
    await reopened.executeTool(call('release_slide_chart_values_change', { change_id: changeId }))
    await proposals.confirm(proposals.pending()!.id)
    expect(records.get(changeId)?.backupReleasedAt).toMatch(/^\d{4}-/)
    expect(savedBytes.length).toBe(0)
    expect(current).not.toBe('')
    const interruptedId = 'interrupted_chart'
    records.set(interruptedId, {
      ...appliedRecord,
      changeId: interruptedId,
      state: 'write_pending',
      newSlideId: undefined,
    })
    activeSlideId = 's1-restored'
    current = original
    const recovered = createPowerPointSkill({ adapter: fake, proposals, chartSavepoint })
    const inspection = await recovered.executeTool(
      call('inspect_slide_chart_values_change', { change_id: interruptedId }),
    )
    expect(inspection.output).toContain('"hostStatus":"before"')
    await recovered.executeTool(
      call('resume_slide_chart_values_change', { change_id: interruptedId }),
    )
    await proposals.confirm(proposals.pending()!.id)
    expect(records.get(interruptedId)?.state).toBe('cancelled')
    activeSlideId = 's1'
    const unavailable = createPowerPointSkill({
      adapter: fake,
      proposals: createStructuredProposalController(),
    })
    expect(unavailable.tools.some((tool) => tool.name === 'update_slide_chart_values')).toBe(false)
    current = original
    await skill.executeTool(
      call('update_slide_chart_values', { slide_index: 0, shape_id: '8', values: [['6']] }),
    )
    const drifted = await JSZip.loadAsync(current, { base64: true })
    drifted.file(
      'ppt/slides/slide1.xml',
      '<p:sld><p:graphicFrame><p:cNvPr id="8"/><c:chart r:id="rId5"/></p:graphicFrame><p:sp name="manual"/></p:sld>',
    )
    current = await drifted.generateAsync({ type: 'base64' })
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.replaceSlidePackage).toHaveBeenCalledTimes(2)
    current = original
    const failedProposals = createStructuredProposalController()
    const noBackup = createPowerPointSkill({
      adapter: fake,
      proposals: failedProposals,
      chartSavepoint: {
        ...chartSavepoint,
        request: async () => {
          throw new Error('backup_unavailable')
        },
      },
    })
    await noBackup.executeTool(
      call('update_slide_chart_values', { slide_index: 0, shape_id: '8', values: [['7']] }),
    )
    await expect(failedProposals.confirm(failedProposals.pending()!.id)).rejects.toThrow(
      'backup_unavailable',
    )
    expect(fake.replaceSlidePackage).toHaveBeenCalledTimes(2)
    const orphanProposals = createStructuredProposalController()
    const release = vi.fn(chartSavepoint.request)
    const journalFailure = createPowerPointSkill({
      adapter: fake,
      proposals: orphanProposals,
      chartSavepoint: {
        ...chartSavepoint,
        request: release,
        writeExistingChartChange: async () => {
          throw new Error('journal_unavailable')
        },
      },
    })
    await journalFailure.executeTool(
      call('update_slide_chart_values', { slide_index: 0, shape_id: '8', values: [['7']] }),
    )
    await expect(orphanProposals.confirm(orphanProposals.pending()!.id)).rejects.toThrow(
      'journal_unavailable',
    )
    expect(
      release.mock.calls.some(
        ([body]) => (body as Record<string, unknown>).operation === 'existing_page_backup_release',
      ),
    ).toBe(true)
    expect(savedBytes.length).toBe(0)
    expect(fake.replaceSlidePackage).toHaveBeenCalledTimes(2)
    const uncertainProposals = createStructuredProposalController()
    const uncertainRequest = vi.fn(chartSavepoint.request)
    const uncertainJournal = createPowerPointSkill({
      adapter: fake,
      proposals: uncertainProposals,
      chartSavepoint: {
        ...chartSavepoint,
        request: uncertainRequest,
        writeExistingChartChange: async (record) => {
          records.set(record.changeId, structuredClone(record))
          throw new Error('journal_ack_lost')
        },
      },
    })
    await uncertainJournal.executeTool(
      call('update_slide_chart_values', { slide_index: 0, shape_id: '8', values: [['7']] }),
    )
    await expect(uncertainProposals.confirm(uncertainProposals.pending()!.id)).rejects.toThrow(
      'journal_ack_lost',
    )
    expect(
      uncertainRequest.mock.calls.some(
        ([body]) => (body as Record<string, unknown>).operation === 'existing_page_backup_release',
      ),
    ).toBe(false)
    expect(savedBytes.length).toBeGreaterThan(0)
    expect(fake.replaceSlidePackage).toHaveBeenCalledTimes(2)
  })

  it('validates master edits from the targeted XML instead of volatile package bytes', async () => {
    const first = new JSZip()
    first.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    first.file('docProps/core.xml', '<core modified="one"/>')
    const second = new JSZip()
    second.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    second.file('docProps/core.xml', '<core modified="two"/>')
    let current = await first.generateAsync({ type: 'base64' })
    const confirmSnapshot = await second.generateAsync({ type: 'base64' })
    let exports = 0
    const fake = adapter({
      exportSlidePackage: vi.fn().mockImplementation(() => {
        exports += 1
        const base64 = exports === 2 || exports === 3 ? confirmSnapshot : current
        return Promise.resolve({ slideId: 's1', base64, fingerprint: `volatile-${exports}` })
      }),
      replaceSlidePackage: vi.fn().mockImplementation((_index, base64) => {
        current = base64
        return Promise.resolve({ slideId: 's1' })
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(
      call('edit_slide_master_xml', {
        program: {
          version: 1,
          operations: [
            {
              op: 'replace_xml',
              path: 'ppt/slideMasters/slideMaster1.xml',
              xml: '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
            },
          ],
        },
      }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.replaceSlidePackage).toHaveBeenCalledOnce()
    const appliedBase64 = vi.mocked(fake.replaceSlidePackage).mock.calls[0]?.[1]
    const applied = await JSZip.loadAsync(appliedBase64!, { base64: true })
    await expect(applied.file('docProps/core.xml')?.async('string')).resolves.toContain('two')
  })

  it('does not overwrite a target XML change between validation and execution', async () => {
    const original = new JSZip()
    original.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="urn:p"/>')
    const changed = new JSZip()
    changed.file(
      'ppt/slideMasters/slideMaster1.xml',
      '<p:sldMaster xmlns:p="urn:p"><p:changed-by-user/></p:sldMaster>',
    )
    const originalBase64 = await original.generateAsync({ type: 'base64' })
    const changedBase64 = await changed.generateAsync({ type: 'base64' })
    let exports = 0
    const fake = adapter({
      exportSlidePackage: vi.fn().mockImplementation(() => {
        exports += 1
        return Promise.resolve({
          slideId: 's1',
          base64: exports < 3 ? originalBase64 : changedBase64,
          fingerprint: `volatile-${exports}`,
        })
      }),
      replaceSlidePackage: vi.fn(),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('edit_slide_master_xml', {
        program: {
          version: 1,
          operations: [
            {
              op: 'replace_xml',
              path: 'ppt/slideMasters/slideMaster1.xml',
              xml: '<p:sldMaster xmlns:p="urn:p"><p:cSld/></p:sldMaster>',
            },
          ],
        },
      }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.replaceSlidePackage).not.toHaveBeenCalled()
  })

  it('executes only confirmed declarative PowerPoint operations and verifies text', async () => {
    const fake = adapter({
      exportSlidePackage: vi
        .fn()
        .mockResolvedValue({ slideId: 's1', base64: 'ppt', fingerprint: 'same' }),
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: [] }),
      readSlideText: vi
        .fn()
        .mockResolvedValue({ slideId: 's1', shapeId: '2', text: 'New', paragraphs: ['New'] }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const code =
      '{"version":1,"operations":[{"op":"set_shape_text","slide_index":0,"shape_id":"2","text":"New"}]}'
    await skill.executeTool(call('execute_office_js', { code }))
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
    await proposals.confirm(proposals.pending()!.id)
    expect(fake.executeDeclarative).toHaveBeenCalledWith(
      [{ op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'New' }],
      expect.any(AbortSignal),
    )
    expect(fake.snapshotSlide).toHaveBeenCalledTimes(2)
    expect(fake.exportSlidePackage).not.toHaveBeenCalled()
  })

  it('accepts strict declarative geometry, text-box creation, and shape deletion families', async () => {
    const fake = adapter({
      exportSlidePackage: vi.fn().mockResolvedValue({
        slideId: 's1',
        base64: 'ppt',
        fingerprint: 'same',
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const code = JSON.stringify({
      version: 1,
      operations: [
        {
          op: 'set_shape_geometry',
          slide_index: 0,
          shape_id: '2',
          left: 1,
          top: 2,
          width: 3,
          height: 4,
        },
        {
          op: 'add_text_box',
          slide_index: 0,
          name: 'Agent box',
          text: 'Hi',
          left: 5,
          top: 6,
          width: 70,
          height: 20,
        },
        { op: 'delete_shape', slide_index: 0, shape_id: '9' },
      ],
    })
    await expect(skill.executeTool(call('execute_office_js', { code }))).resolves.toMatchObject({
      mutated: false,
      output: expect.stringContaining('set_shape_geometry'),
    })
    expect(proposals.pending()?.impact.count).toBe(3)
    proposals.reject()
    await expect(
      skill.executeTool(
        call('execute_office_js', {
          code: JSON.stringify({
            version: 1,
            operations: [
              {
                op: 'set_shape_geometry',
                slide_index: 0,
                shape_id: '2',
                left: 1,
                top: 2,
                width: -1,
                height: 4,
              },
            ],
          }),
        }),
      ),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
  })

  it('accepts host-normalized PowerPoint geometry when verifying a created text box', async () => {
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: ['4'] }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({ slideId: 's1', slideIndex: 0, shapes: [] })
        .mockResolvedValue({
          slideId: 's1',
          slideIndex: 0,
          shapes: [
            {
              id: '4',
              name: 'Status',
              type: 'TextBox',
              left: 300.00003,
              top: 449.99997,
              width: 360.00003,
              height: 50.00003,
            },
          ],
        }),
      readSlideText: vi.fn().mockResolvedValue({
        slideId: 's1',
        shapeId: '4',
        text: 'PASS',
        paragraphs: ['PASS'],
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const code = JSON.stringify({
      version: 1,
      operations: [
        {
          op: 'add_text_box',
          slide_index: 0,
          name: 'Status',
          text: 'PASS',
          left: 300,
          top: 450,
          width: 360,
          height: 50,
        },
      ],
    })

    await skill.executeTool(call('execute_office_js', { code }))
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
  })

  it('waits for delayed PowerPoint text readback before rejecting an applied edit', async () => {
    const fake = adapter({
      readSlideText: vi
        .fn()
        .mockResolvedValueOnce({
          slideId: 's1',
          shapeId: '2',
          text: 'Old',
          paragraphs: ['Old'],
        })
        .mockResolvedValueOnce({
          slideId: 's1',
          shapeId: '2',
          text: 'Old',
          paragraphs: ['Old'],
        })
        .mockResolvedValueOnce({
          slideId: 's1',
          shapeId: '2',
          text: 'Old',
          paragraphs: ['Old'],
        })
        .mockResolvedValue({
          slideId: 's1',
          shapeId: '2',
          text: 'New',
          paragraphs: ['New'],
        }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'New' }))

    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.readSlideText).toHaveBeenCalledTimes(4)
  })

  it('waits for delayed geometry, delete, and duplicate readback', async () => {
    const beforeShape = {
      id: '2',
      name: 'Title',
      type: 'TextBox',
      left: 10,
      top: 20,
      width: 200,
      height: 40,
    }
    const afterShape = { ...beforeShape, left: 30 }
    const deletedShape = { ...beforeShape, id: '9', name: 'Remove me' }
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: [] }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({
          slideId: 's1',
          slideIndex: 0,
          shapes: [beforeShape, deletedShape],
        })
        .mockResolvedValueOnce({
          slideId: 's1',
          slideIndex: 0,
          shapes: [afterShape, deletedShape],
        })
        .mockResolvedValueOnce({
          slideId: 's1',
          slideIndex: 0,
          shapes: [afterShape, deletedShape],
        })
        .mockResolvedValueOnce({ slideId: 's1', slideIndex: 0, shapes: [afterShape] }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(
      call('execute_office_js', {
        code: JSON.stringify({
          version: 1,
          operations: [
            {
              op: 'set_shape_geometry',
              slide_index: 0,
              shape_id: '2',
              left: 30,
              top: 20,
              width: 200,
              height: 40,
            },
            { op: 'delete_shape', slide_index: 0, shape_id: '9' },
          ],
        }),
      }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
  })

  it('waits for a delayed duplicate-slide collection readback', async () => {
    const fake = adapter({
      duplicateSlide: vi.fn().mockResolvedValue({ slideId: 'copy' }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 1, shapes: [] })
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 1, shapes: [] })
        .mockResolvedValue({ slideId: 'copy', slideIndex: 1, shapes: [] }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(call('duplicate_slide', { slide_index: 0 }))
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.listSlideShapes).toHaveBeenCalledTimes(3)
  })

  it('requires the exact declarative duplicate receipt instead of accepting any following slide', async () => {
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({
        createdShapeIds: [],
        insertedSlideId: 'copy',
      } as never),
      listSlideShapes: vi.fn().mockResolvedValue({
        slideId: 'unrelated-existing-slide',
        slideIndex: 1,
        shapes: [],
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await skill.executeTool(
      call('execute_office_js', {
        code: JSON.stringify({
          version: 1,
          operations: [{ op: 'duplicate_slide', slide_index: 0 }],
        }),
      }),
    )

    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('office_verify_failed')
  })

  it('rejects a declarative mutation that edits and then deletes the same shape', async () => {
    const fake = adapter({ executeDeclarative: vi.fn() })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })

    await expect(
      skill.executeTool(
        call('execute_office_js', {
          code: JSON.stringify({
            version: 1,
            operations: [
              { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'temporary' },
              { op: 'delete_shape', slide_index: 0, shape_id: '2' },
            ],
          }),
        }),
      ),
    ).resolves.toMatchObject({ output: 'invalid_tool_input', isError: true })
    expect(proposals.pending()).toBeUndefined()
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
  })

  it('maps adapter internals to stable errors and validates screenshots', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter({
        listSlideShapes: vi.fn().mockRejectedValue(new Error('secret')),
        screenshotSlide: vi
          .fn()
          .mockResolvedValue({ slideId: 'host-slide-1', mime: 'image/png', base64: 'bad!' }),
      }),
      proposals: createStructuredProposalController(),
    })
    await expect(
      skill.executeTool(call('list_slide_shapes', { slide_index: 0 })),
    ).resolves.toMatchObject({ output: 'office_read_failed', isError: true })
    await expect(
      skill.executeTool(call('screenshot_slide', { slide_index: 0 })),
    ).resolves.toMatchObject({ output: 'office_read_failed', isError: true })
  })
})

describe('browser PowerPoint adapter', () => {
  it.each(['complete', 'empty', 'missing', 'duplicate', 'overflow', 'cancelled'])(
    'reads %s complete native style dependencies',
    async (mode) => {
      const slide = { id: 's1', slideMaster: { id: 'm1' }, layout: { id: 'l1' } }
      const items =
        mode === 'empty'
          ? []
          : mode === 'missing'
            ? [{ ...slide, layout: undefined }]
            : mode === 'duplicate'
              ? [slide, slide]
              : mode === 'overflow'
                ? Array.from({ length: 101 }, (_, index) => ({ ...slide, id: String(index) }))
                : [slide]
      const slides = { items, load: vi.fn() }
      const controller = new AbortController()
      const run = vi.fn(async (callback) =>
        callback({
          presentation: { slides },
          sync: async () => {
            if (mode === 'cancelled') controller.abort()
          },
        }),
      )
      Object.assign(globalThis, {
        Office: { context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } } },
        PowerPoint: { run },
      })
      const result = new BrowserPowerPointAdapter().inspectStyleDependencies(controller.signal)
      if (mode === 'complete' || mode === 'empty') {
        await expect(result).resolves.toEqual({
          slides: mode === 'empty' ? [] : [{ slideId: 's1', masterId: 'm1', layoutId: 'l1' }],
        })
        expect(slides.load).toHaveBeenCalledWith('items/id,items/slideMaster/id,items/layout/id')
      } else
        await expect(result).rejects.toThrow(
          mode === 'cancelled' ? 'cancelled' : 'office_read_failed',
        )
    },
  )

  it('rejects master package replacement on Mac before entering PowerPoint.run', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          platform: 'Mac',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run },
    })

    await expect(
      new BrowserPowerPointAdapter().replaceSlidePackage(0, 'ppt', true),
    ).rejects.toThrow('office_api_unsupported')
    expect(run).not.toHaveBeenCalled()
  })

  const originals = { Office: globalThis.Office, PowerPoint: globalThis.PowerPoint }
  afterEach(() => Object.assign(globalThis, originals))

  it('detects host/API support before PowerPoint.run', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: { host: 'Word', requirements: { isSetSupported: vi.fn().mockReturnValue(true) } },
      },
      PowerPoint: { run },
    })
    await expect(new BrowserPowerPointAdapter().listSlideShapes(0)).rejects.toThrow(
      'office_api_unsupported',
    )
    expect(run).not.toHaveBeenCalled()
  })

  it('requires the per-operation PowerPoint API set', async () => {
    const run = vi.fn()
    const supports = vi.fn((_name: string, version: string) => version === '1.4')
    Object.assign(globalThis, {
      Office: { context: { host: 'PowerPoint', requirements: { isSetSupported: supports } } },
      PowerPoint: { run },
    })
    await expect(new BrowserPowerPointAdapter().screenshotSlide(0)).rejects.toThrow(
      'office_api_unsupported',
    )
    await expect(new BrowserPowerPointAdapter().verifySlides()).rejects.toThrow(
      'office_api_unsupported',
    )
    expect(run).not.toHaveBeenCalled()
    expect(supports).toHaveBeenCalledWith('PowerPointApi', '1.8')
    expect(supports).toHaveBeenCalledWith('PowerPointApi', '1.10')
  })

  it('maps native master operations to PowerPointApi 1.10 objects', async () => {
    const setSolidFill = vi.fn()
    const setThemeColor = vi.fn()
    const layoutBackground: Record<string, unknown> = {}
    const layout = { background: layoutBackground }
    const master = {
      background: { fill: { setSolidFill } },
      themeColorScheme: { setThemeColor },
      layouts: { getItem: vi.fn().mockReturnValue(layout) },
    }
    const context = {
      presentation: { slideMasters: { getItem: vi.fn().mockReturnValue(master) } },
      sync: vi.fn().mockResolvedValue(undefined),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run: (callback: (value: typeof context) => unknown) => callback(context) },
    })

    await new BrowserPowerPointAdapter().executeMasterOperations([
      {
        op: 'set_master_background',
        master_id: 'm1',
        fill: { type: 'solid', color: '#000000', transparency: 0.2 },
      },
      {
        op: 'set_master_theme_color',
        master_id: 'm1',
        theme_color: 'Light1',
        color: '#FFFFFF',
      },
      {
        op: 'set_layout_background_following',
        master_id: 'm1',
        layout_id: 'l1',
        follow_master: true,
        show_master_graphics: false,
      },
    ])

    expect(setSolidFill).toHaveBeenCalledWith({ color: '#000000', transparency: 0.2 })
    expect(setThemeColor).toHaveBeenCalledWith('Light1', '#FFFFFF')
    expect(layoutBackground).toMatchObject({
      isMasterBackgroundFollowed: true,
      areBackgroundGraphicsHidden: true,
    })
  })

  it('routes declarative duplication through the transaction-safe duplicate primitive', async () => {
    const subject = new BrowserPowerPointAdapter()
    const duplicate = vi.spyOn(subject, 'duplicateSlide').mockResolvedValue({ slideId: 'copy' })

    await expect(
      subject.executeDeclarative([{ op: 'duplicate_slide', slide_index: 0 }]),
    ).resolves.toEqual({ createdShapeIds: [], insertedSlideId: 'copy' })
    expect(duplicate).toHaveBeenCalledWith(0, undefined)
  })

  it('defensively rejects edit-then-delete before entering PowerPoint.run', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'temporary' },
        { op: 'delete_shape', slide_index: 0, shape_id: '2' },
      ]),
    ).rejects.toThrow('invalid_tool_input')
    expect(run).not.toHaveBeenCalled()
  })

  it('creates a native geometric shape from a bounded declarative operation', async () => {
    const fill = { setSolidColor: vi.fn() }
    const created = { id: 'new-shape', name: '', fill, lineFormat: { color: '' }, load: vi.fn() }
    const addGeometricShape = vi.fn(() => created)
    const slide = { id: 's1', load: vi.fn(), shapes: { addGeometricShape } }
    const slides = { getCount: vi.fn(() => ({ value: 1 })), getItemAt: vi.fn(() => slide) }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync: vi.fn().mockResolvedValue(undefined) }),
      },
    })
    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        {
          op: 'add_geometric_shape',
          slide_index: 0,
          name: 'step',
          shape: 'roundRect',
          left: 72,
          top: 180,
          width: 216,
          height: 144,
          fill: '2255AA',
          lineColor: '2255AA',
        },
      ]),
    ).resolves.toEqual({ createdShapeIds: ['new-shape'] })
    expect(addGeometricShape).toHaveBeenCalledWith('RoundRectangle', {
      left: 72,
      top: 180,
      width: 216,
      height: 144,
    })
    expect(fill.setSolidColor).toHaveBeenCalledWith('#2255AA')
    expect(created.lineFormat.color).toBe('#2255AA')
  })

  it('applies shared text style and alignment when creating a native text box', async () => {
    const font: Record<string, unknown> = {}
    const paragraphFormat: Record<string, unknown> = {}
    const created = {
      id: 'new-text',
      name: '',
      load: vi.fn(),
      textFrame: { textRange: { font, paragraphFormat } },
    }
    const addTextBox = vi.fn(() => created)
    const slide = { id: 's1', load: vi.fn(), shapes: { addTextBox } }
    const slides = { getCount: vi.fn(() => ({ value: 1 })), getItemAt: vi.fn(() => slide) }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync: vi.fn().mockResolvedValue(undefined) }),
      },
    })
    await new BrowserPowerPointAdapter().executeDeclarative([
      {
        op: 'add_text_box',
        slide_index: 0,
        name: 'title',
        text: 'Centered',
        left: 72,
        top: 72,
        width: 720,
        height: 72,
        fontFace: 'Microsoft YaHei',
        fontSize: 32,
        color: '172033',
        bold: true,
        align: 'center',
        margin: 0,
        verticalAlignment: 'top',
      },
    ])
    expect(font).toMatchObject({ name: 'Microsoft YaHei', size: 32, color: '#172033', bold: true })
    expect(paragraphFormat.horizontalAlignment).toBe('Center')
    expect(created.textFrame).toMatchObject({
      leftMargin: 0,
      rightMargin: 0,
      topMargin: 0,
      bottomMargin: 0,
      verticalAlignment: 'Top',
    })
  })

  it('confirms and verifies a native geometric shape creation', async () => {
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: ['new-shape'] }),
      listSlideShapes: vi.fn().mockResolvedValue({
        slideId: 'slide-1',
        slideIndex: 0,
        shapes: [
          {
            id: 'new-shape',
            name: 'step',
            type: 'GeometricShape',
            left: 72,
            top: 180,
            width: 216,
            height: 144,
          },
        ],
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const operation = {
      op: 'add_geometric_shape',
      slide_index: 0,
      name: 'step',
      shape: 'roundRect',
      left: 72,
      top: 180,
      width: 216,
      height: 144,
      fill: '2255AA',
      lineColor: '2255AA',
    }
    await skill.executeTool(
      call('execute_office_js', { code: JSON.stringify({ version: 1, operations: [operation] }) }),
    )
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.executeDeclarative).toHaveBeenCalledWith([operation], expect.any(AbortSignal))
  })

  it('creates and verifies a native table with exact cell values', async () => {
    const created = { id: 'new-table', name: '', load: vi.fn() }
    const addTable = vi.fn(() => created)
    const slide = { id: 's1', load: vi.fn(), shapes: { addTable } }
    const slides = { getCount: vi.fn(() => ({ value: 1 })), getItemAt: vi.fn(() => slide) }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync: vi.fn().mockResolvedValue(undefined) }),
      },
    })
    const operation = {
      op: 'add_native_table' as const,
      slide_index: 0,
      name: 'table',
      rows: [
        ['方案', '结果'],
        ['甲', '120'],
      ],
      left: 72,
      top: 180,
      width: 576,
      height: 144,
      fontFace: 'Microsoft YaHei',
      fontSize: 12,
      color: '172033',
      borderColor: '2255AA',
      cellMargin: 2.88,
    }
    await expect(new BrowserPowerPointAdapter().executeDeclarative([operation])).resolves.toEqual({
      createdShapeIds: ['new-table'],
    })
    expect(addTable).toHaveBeenCalledWith(
      2,
      2,
      expect.objectContaining({ values: operation.rows, width: 576 }),
    )
    expect(addTable).toHaveBeenCalledWith(
      2,
      2,
      expect.objectContaining({
        uniformCellProperties: expect.objectContaining({
          borders: expect.objectContaining({
            top: { color: '#2255AA', weight: 1 },
            right: { color: '#2255AA', weight: 1 },
          }),
          margins: { top: 2.88, right: 2.88, bottom: 2.88, left: 2.88 },
        }),
      }),
    )
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: ['new-table'] }),
      listSlideShapes: vi.fn().mockResolvedValue({
        slideId: 'slide-1',
        slideIndex: 0,
        shapes: [
          {
            id: 'new-table',
            name: 'table',
            type: 'Table',
            left: 72,
            top: 180,
            width: 576,
            height: 144,
          },
        ],
      }),
      readSlideTable: vi.fn().mockResolvedValue(operation.rows),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('execute_office_js', { code: JSON.stringify({ version: 1, operations: [operation] }) }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.readSlideTable).toHaveBeenCalledWith(0, 'new-table', expect.any(AbortSignal))
  })

  it('confirms a complete supported SlideIR page and rejects unsupported pages before writing', async () => {
    const deck = benchmarkDeck()
    const fake = adapter({
      executeDeclarative: vi
        .fn()
        .mockResolvedValue({ createdShapeIds: ['title-host', 'shape-host', 'source-host'] }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 0, shapes: [] })
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 0, shapes: [] })
        .mockResolvedValue({
          slideId: 'slide-1',
          slideIndex: 0,
          shapes: [
            {
              id: 'title-host',
              name: 'title',
              type: 'TextBox',
              left: 72,
              top: 72,
              width: 720,
              height: 72,
            },
            {
              id: 'shape-host',
              name: 'step',
              type: 'GeometricShape',
              left: 72,
              top: 180,
              width: 216,
              height: 144,
            },
            {
              id: 'source-host',
              name: 'source-attribution',
              type: 'TextBox',
              left: 36,
              top: 507.6,
              width: 885.6,
              height: 21.6,
            },
          ],
        }),
      readSlideText: vi.fn().mockImplementation(async (_index, shapeId) => ({
        slideId: 'slide-1',
        shapeId,
        text: shapeId === 'source-host' ? '[source-1] 研究报告（合成基准） · 第 1 页' : '研究流程',
        paragraphs: [
          shapeId === 'source-host' ? '[source-1] 研究报告（合成基准） · 第 1 页' : '研究流程',
        ],
      })),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const proposed = await skill.executeTool(
      call('add_slide_ir_objects', {
        slide_index: 0,
        slide: deck.slides[3],
        style: deck.style,
        claims: deck.claims,
      }),
    )
    expect(proposed).toMatchObject({
      mutated: false,
      summary: 'Proposed declarative PowerPoint execution',
    })
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.executeDeclarative).toHaveBeenCalledWith(
      expect.arrayContaining([
        expect.objectContaining({ op: 'add_text_box', name: 'title' }),
        expect.objectContaining({ op: 'add_geometric_shape', name: 'step' }),
      ]),
      expect.any(AbortSignal),
    )
    const count = (fake.executeDeclarative as ReturnType<typeof vi.fn>).mock.calls.length
    await expect(
      skill.executeTool(
        call('add_slide_ir_objects', {
          slide_index: 0,
          slide: deck.slides[2],
          style: deck.style,
          claims: deck.claims,
        }),
      ),
    ).resolves.toMatchObject({ isError: true })
    expect((fake.executeDeclarative as ReturnType<typeof vi.fn>).mock.calls).toHaveLength(count)
  })

  it('requires a declared resolved font before proposing direct SlideIR writes', async () => {
    const deck = benchmarkDeck()
    deck.style.fontFace = 'WisWork Benchmark Display 2026'
    deck.style.fontFallbacks = ['Noto Sans CJK SC']
    const fake = adapter()
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const input = {
      slide_index: 0,
      slide: deck.slides[3],
      style: deck.style,
      claims: deck.claims,
    }
    expect(await skill.executeTool(call('add_slide_ir_objects', input))).toMatchObject({
      isError: true,
    })
    expect(
      await skill.executeTool(
        call('add_slide_ir_objects', { ...input, resolved_font_face: 'Unlisted Font' }),
      ),
    ).toMatchObject({ isError: true })
    const result = await skill.executeTool(
      call('add_slide_ir_objects', { ...input, resolved_font_face: 'Noto Sans CJK SC' }),
    )
    expect(result.isError).not.toBe(true)
    expect(proposals.pending()).toBeTruthy()
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
  })

  it('routes a SlideIR table page through native table readback', async () => {
    const deck = benchmarkDeck()
    const rows = (deck.slides[5]!.elements[1] as { rows: string[][] }).rows
    const fake = adapter({
      executeDeclarative: vi
        .fn()
        .mockResolvedValue({ createdShapeIds: ['title-host', 'table-host', 'source-host'] }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 0, shapes: [] })
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 0, shapes: [] })
        .mockResolvedValue({
          slideId: 'slide-1',
          slideIndex: 0,
          shapes: [
            {
              id: 'title-host',
              name: 'title',
              type: 'TextBox',
              left: 72,
              top: 72,
              width: 720,
              height: 72,
            },
            {
              id: 'table-host',
              name: 'table',
              type: 'Table',
              left: 72,
              top: 180,
              width: 576,
              height: 144,
            },
            {
              id: 'source-host',
              name: 'source-attribution',
              type: 'TextBox',
              left: 36,
              top: 507.6,
              width: 885.6,
              height: 21.6,
            },
          ],
        }),
      readSlideText: vi.fn().mockImplementation(async (_index, shapeId) => ({
        slideId: 'slide-1',
        shapeId,
        text: shapeId === 'source-host' ? '[source-1] 研究报告（合成基准） · 第 1 页' : '实验表格',
        paragraphs: [
          shapeId === 'source-host' ? '[source-1] 研究报告（合成基准） · 第 1 页' : '实验表格',
        ],
      })),
      readSlideTable: vi.fn().mockResolvedValue(rows),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    const proposed = await skill.executeTool(
      call('add_slide_ir_objects', {
        slide_index: 0,
        slide: deck.slides[5],
        style: deck.style,
        claims: deck.claims,
      }),
    )
    expect(proposed.mutated).toBe(false)
    await expect(proposals.confirm(proposals.pending()!.id)).resolves.toBeUndefined()
    expect(fake.readSlideTable).toHaveBeenCalledWith(0, 'table-host', expect.any(AbortSignal))
  })

  it('rejects a SlideIR page when the host already has an object with a planned name', async () => {
    const deck = benchmarkDeck()
    const fake = adapter({
      listSlideShapes: vi.fn().mockResolvedValue({
        slideId: 'slide-1',
        slideIndex: 0,
        shapes: [
          { id: 'old', name: 'title', type: 'TextBox', left: 0, top: 0, width: 100, height: 20 },
        ],
      }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await expect(
      skill.executeTool(
        call('add_slide_ir_objects', {
          slide_index: 0,
          slide: deck.slides[3],
          style: deck.style,
          claims: deck.claims,
        }),
      ),
    ).resolves.toMatchObject({ isError: true, output: 'office_concurrent_change' })
    expect(proposals.pending()).toBeUndefined()
    expect(fake.executeDeclarative).not.toHaveBeenCalled()
  })

  it('does not accept a created SlideIR shape whose host name differs from the planned name', async () => {
    const deck = benchmarkDeck()
    const slide = {
      ...deck.slides[0]!,
      claimIds: [],
      elements: [deck.slides[0]!.elements[0]!],
    }
    const fake = adapter({
      executeDeclarative: vi.fn().mockResolvedValue({ createdShapeIds: ['created'] }),
      listSlideShapes: vi
        .fn()
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 0, shapes: [] })
        .mockResolvedValueOnce({ slideId: 'slide-1', slideIndex: 0, shapes: [] })
        .mockResolvedValue({
          slideId: 'slide-1',
          slideIndex: 0,
          shapes: [
            {
              id: 'created',
              name: 'wrong-name',
              type: 'TextBox',
              left: 72,
              top: 72,
              width: 720,
              height: 72,
            },
          ],
        }),
    })
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('add_slide_ir_objects', {
        slide_index: 0,
        slide,
        style: deck.style,
        claims: deck.claims,
      }),
    )
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('office_verify_failed')
  })

  it('reads native table values with bounded dimensions', async () => {
    const table = {
      values: [
        ['方案', '结果'],
        ['甲', '120'],
      ],
      rowCount: 2,
      columnCount: 2,
      load: vi.fn(),
    }
    const shape = { getTable: vi.fn(() => table) }
    const slide = { id: 's1', load: vi.fn(), shapes: { getItem: vi.fn(() => shape) } }
    const slides = { getCount: vi.fn(() => ({ value: 1 })), getItemAt: vi.fn(() => slide) }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync: vi.fn().mockResolvedValue(undefined) }),
      },
    })
    const subject = new BrowserPowerPointAdapter()
    await expect(subject.readSlideTable(0, 'table')).resolves.toEqual(table.values)
    table.values = [['bad']]
    await expect(subject.readSlideTable(0, 'table')).rejects.toThrow('office_read_failed')
  })

  it('returns stable IDs/geometry and verifies negative, overflow, and overlap geometry', async () => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const shapes = {
      load: vi.fn(),
      items: [
        { id: '2', name: 'A', type: 'TextBox', left: -5, top: 10, width: 100, height: 50 },
        { id: '3', name: 'B', type: 'TextBox', left: 50, top: 20, width: 950, height: 530 },
      ],
    }
    const slides = {
      load: vi.fn(),
      items: [{ id: 's1', shapes, load: vi.fn() }],
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn((i) => slides.items[i]),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({
            presentation: {
              slides,
              pageSetup: { slideWidth: 960, slideHeight: 540, load: vi.fn() },
            },
            sync,
          }),
      },
    })
    await expect(new BrowserPowerPointAdapter().listSlideShapes(0)).resolves.toMatchObject({
      slideId: 's1',
      shapes: [
        { id: '2', left: -5 },
        { id: '3', left: 50 },
      ],
    })
    await expect(new BrowserPowerPointAdapter().verifySlides()).resolves.toMatchObject({
      slides: [
        {
          overflows: expect.arrayContaining([
            expect.objectContaining({ shapeId: '2', edge: 'left' }),
            expect.objectContaining({ shapeId: '3', edge: 'right' }),
            expect.objectContaining({ shapeId: '3', edge: 'bottom' }),
          ]),
          overlaps: [{ shapeAId: '2', shapeBId: '3', overlapX: 45, overlapY: 40 }],
        },
      ],
    })
    expect(sync).toHaveBeenCalled()
  })

  it('checks cancellation before every write/sync and implements text edit and duplicate', async () => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const slidePackage = await packageZip.generateAsync({ type: 'base64' })
    const sync = vi.fn().mockResolvedValue(undefined)
    const textRange = { text: 'Old', load: vi.fn() }
    const shape = {
      id: '2',
      name: 'Title',
      type: 'Placeholder',
      left: 10,
      top: 20,
      width: 200,
      height: 40,
      textFrame: { hasText: true, load: vi.fn(), textRange },
    }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [shape], getItem: vi.fn(() => shape) },
      exportAsBase64: vi.fn(() => ({ value: slidePackage })),
    }
    const slides = {
      load: vi.fn(),
      items: [slide],
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
    }
    const insertSlidesFromBase64 = vi.fn(() => {
      slides.items.splice(1, 0, { ...slide, id: 's2' })
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })
    const subject = new BrowserPowerPointAdapter()
    await expect(subject.snapshotSlide(0)).resolves.toMatchObject({
      slideId: 's1',
      fingerprint: expect.stringMatching(/^s1:\d+:[0-9a-f]{8}$/),
    })
    await subject.editSlideText(0, '2', 'New')
    expect(textRange.text).toBe('New')
    await expect(subject.duplicateSlide(0)).resolves.toEqual({ slideId: 's2' })
    expect(insertSlidesFromBase64).toHaveBeenCalledWith(slidePackage, { targetSlideId: 's1' })

    const controller = new AbortController()
    controller.abort()
    await expect(subject.editSlideText(0, '2', 'No', controller.signal)).rejects.toThrow(
      'cancelled',
    )
    expect(textRange.text).toBe('New')
  })

  it('reconciles a text sync rejection that committed and never overwrites a third state', async () => {
    const textRange = { text: 'Old', load: vi.fn() }
    const shape = { textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (textRange.text === 'New' && !rejected) {
        rejected = true
        throw new Error('host rejected after commit')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().editSlideText(0, '2', 'New'),
    ).resolves.toBeUndefined()
    expect(textRange.text).toBe('New')

    textRange.text = 'Old'
    sync.mockClear()
    rejected = false
    sync.mockImplementation(async () => {
      if (textRange.text === 'New' && !rejected) {
        rejected = true
        textRange.text = 'User edit'
        throw new Error('host rejected after third-party edit')
      }
    })
    await expect(new BrowserPowerPointAdapter().editSlideText(0, '2', 'New')).rejects.toThrow(
      'office_concurrent_change',
    )
    expect(textRange.text).toBe('User edit')
  })

  it('uses bounded convergence when a rejected text sync is initially read back stale', async () => {
    let visible = 'Old'
    let assigned = 'Old'
    let reconciliationReads = 0
    const textRange = {
      load: vi.fn(),
      get text() {
        return visible
      },
      set text(value: string) {
        assigned = value
      },
    }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => ({ textFrame: { textRange } })) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (assigned === 'New' && !rejected) {
        rejected = true
        throw new Error('host rejected after commit')
      }
      if (rejected && visible !== assigned && ++reconciliationReads >= 3) visible = assigned
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().editSlideText(0, '2', 'New'),
    ).resolves.toBeUndefined()
    expect(visible).toBe('New')
    expect(reconciliationReads).toBe(3)
  })

  it('uses bounded convergence after a rejected declarative sync', async () => {
    let visible = 'Old'
    let assigned = 'Old'
    let reconciliationReads = 0
    const textRange = {
      load: vi.fn(),
      get text() {
        return visible
      },
      set text(value: string) {
        assigned = value
      },
    }
    const shape = { id: '2', load: vi.fn(), textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (assigned === 'New' && !rejected) {
        rejected = true
        throw new Error('host rejected after declarative commit')
      }
      if (rejected && visible !== assigned && ++reconciliationReads >= 3) visible = assigned
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'New' },
      ]),
    ).resolves.toEqual({ createdShapeIds: [] })
    expect(visible).toBe('New')
  })

  it('reports an attributable text commit as applied when cancellation races after sync', async () => {
    const controller = new AbortController()
    const textRange = { text: 'Old', load: vi.fn() }
    const shape = { textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let aborted = false
    const sync = vi.fn().mockImplementation(async () => {
      if (textRange.text === 'New' && !aborted) {
        aborted = true
        controller.abort()
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().editSlideText(0, '2', 'New', controller.signal),
    ).resolves.toBeUndefined()
    expect(textRange.text).toBe('New')
  })

  it('converges delayed text visibility even when cancellation races readback', async () => {
    const runScenario = async (abortDuringReadback: boolean) => {
      const controller = new AbortController()
      let visible = 'Old'
      let pending: string | undefined
      let readbacks = 0
      const textRange = {
        load: vi.fn(),
        get text() {
          return visible
        },
        set text(value: string) {
          pending = value
        },
      }
      const shape = { textFrame: { textRange } }
      const slide = {
        id: 's1',
        load: vi.fn(),
        shapes: { getItem: vi.fn(() => shape) },
      }
      const slides = {
        getCount: vi.fn(() => ({ value: 1 })),
        getItemAt: vi.fn(() => slide),
      }
      const sync = vi.fn().mockImplementation(async () => {
        if (pending === 'New') {
          readbacks += 1
          if (abortDuringReadback && readbacks === 2) controller.abort()
          if (readbacks >= 3) visible = pending
        } else if (pending === 'Old') visible = pending
      })
      Object.assign(globalThis, {
        Office: {
          context: {
            host: 'PowerPoint',
            requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
          },
        },
        PowerPoint: {
          run: (callback: (context: unknown) => unknown) =>
            callback({ presentation: { slides }, sync }),
        },
      })
      const result = new BrowserPowerPointAdapter().editSlideText(0, '2', 'New', controller.signal)
      if (abortDuringReadback) {
        await expect(result).resolves.toBeUndefined()
        expect(visible).toBe('New')
      } else {
        await expect(result).resolves.toBeUndefined()
        expect(visible).toBe('New')
      }
    }

    await runScenario(false)
    await runScenario(true)
  })

  it('reports an attributable declarative prefix as uncertain without restoring it', async () => {
    const textRange = { text: 'Old', load: vi.fn() }
    const textShape = { id: '2', load: vi.fn(), textFrame: { textRange } }
    const geometryShape = {
      id: '3',
      load: vi.fn(),
      left: 10,
      top: 20,
      width: 200,
      height: 40,
    }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: {
        getItem: vi.fn((id: string) => (id === '2' ? textShape : geometryShape)),
      },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (textRange.text === 'New' && geometryShape.left === 30 && !rejected) {
        rejected = true
        geometryShape.left = 10
        throw new Error('host rejected after committed prefix')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'New' },
        {
          op: 'set_shape_geometry',
          slide_index: 0,
          shape_id: '3',
          left: 30,
          top: 20,
          width: 200,
          height: 40,
        },
      ]),
    ).rejects.toThrow('office_state_uncertain')
    expect(textRange.text).toBe('New')
    expect(geometryShape.left).toBe(10)
  })

  it('does not restore a declarative batch over a concurrent third state', async () => {
    const textRange = { text: 'Old', load: vi.fn() }
    const shape = { id: '2', load: vi.fn(), textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (textRange.text === 'New' && !rejected) {
        rejected = true
        textRange.text = 'User edit'
        throw new Error('host rejected after concurrent edit')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'New' },
      ]),
    ).rejects.toThrow('office_concurrent_change')
    expect(textRange.text).toBe('User edit')
  })

  it('rejects repeated writes to the same shape property before Office dispatch', async () => {
    const run = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: { run },
    })

    await expect(
      new BrowserPowerPointAdapter().executeDeclarative([
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'First' },
        { op: 'set_shape_text', slide_index: 0, shape_id: '2', text: 'Second' },
      ]),
    ).rejects.toThrow('invalid_tool_input')
    expect(run).not.toHaveBeenCalled()
  })

  it.each([
    { mode: 'committed', expected: 'copy' },
    { mode: 'cancelled', expected: 'cancelled' },
    { mode: 'third-state', expected: 'office_concurrent_change' },
  ])('reconciles duplicate-slide sync rejection in $mode state', async ({ mode, expected }) => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const slidePackage = await packageZip.generateAsync({ type: 'base64' })
    const controller = new AbortController()
    const shape = (text: string) => ({
      id: '2',
      name: 'Title',
      type: 'TextBox',
      left: 10,
      top: 20,
      width: 200,
      height: 40,
      textFrame: { hasText: true, load: vi.fn(), textRange: { text, load: vi.fn() } },
    })
    const slides = {
      items: [] as Array<Record<string, unknown>>,
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
    }
    const source = {
      id: 'source',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [shape('Stable')] },
      exportAsBase64: vi.fn(() => ({ value: slidePackage })),
    }
    slides.items.push(source)
    const insertSlidesFromBase64 = vi.fn(() => {
      const inserted = {
        id: 'copy',
        load: vi.fn(),
        shapes: {
          load: vi.fn(),
          items: [shape(mode === 'third-state' ? 'User slide' : 'Stable')],
        },
        exportAsBase64: vi.fn(() => ({ value: slidePackage })),
        delete: vi.fn(() => {
          const index = slides.items.indexOf(inserted)
          if (index >= 0) slides.items.splice(index, 1)
        }),
      }
      slides.items.splice(1, 0, inserted)
    })
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (slides.items.length === 2 && !rejected) {
        rejected = true
        if (mode === 'cancelled') controller.abort()
        throw new Error('host rejected after insert')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })

    const result = new BrowserPowerPointAdapter().duplicateSlide(0, controller.signal)
    if (mode === 'committed' || mode === 'cancelled')
      await expect(result).resolves.toEqual({ slideId: 'copy' })
    else await expect(result).rejects.toThrow(expected)
    expect(slides.items).toHaveLength(2)
  })

  it('waits for duplicate collection convergence and proves package ownership after sync rejection', async () => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>')
    const base64 = await packageZip.generateAsync({ type: 'base64' })
    const source = { id: 'source', load: vi.fn(), exportAsBase64: vi.fn(() => ({ value: base64 })) }
    const slides = {
      getCount: vi.fn(() => ({ value: 2 })),
      getItemAt: vi.fn(() => source),
    }
    const insertSlidesFromBase64 = vi.fn()
    let rejected = false
    const sync = vi.fn().mockImplementation(async () => {
      if (insertSlidesFromBase64.mock.calls.length > 0 && !rejected) {
        rejected = true
        throw new Error('host rejected after committed insert')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })
    const subject = new BrowserPowerPointAdapter()
    vi.spyOn(subject, 'snapshotSlide')
      .mockResolvedValueOnce({ slideId: 'source', fingerprint: 'source:semantic' })
      .mockResolvedValueOnce({ slideId: 'existing', fingerprint: 'existing:other' })
      .mockResolvedValueOnce({ slideId: 'source', fingerprint: 'source:semantic' })
      .mockResolvedValueOnce({ slideId: 'source', fingerprint: 'source:semantic' })
      .mockResolvedValueOnce({ slideId: 'existing', fingerprint: 'existing:other' })
      .mockResolvedValueOnce({ slideId: 'existing', fingerprint: 'existing:other' })
      .mockResolvedValueOnce({ slideId: 'copy', fingerprint: 'copy:semantic' })
      .mockResolvedValue({ slideId: 'source', fingerprint: 'source:semantic' })
    vi.spyOn(subject, 'exportSlidePackage')
      .mockResolvedValueOnce({ slideId: 'source', base64, fingerprint: 'volatile-source' })
      .mockResolvedValue({ slideId: 'copy', base64, fingerprint: 'volatile-copy' })

    await expect(subject.duplicateSlide(0)).resolves.toEqual({ slideId: 'copy' })
    expect(subject.snapshotSlide).toHaveBeenCalledTimes(8)
    expect(subject.exportSlidePackage).toHaveBeenCalledWith(1)
  })

  it('revalidates the duplicate source immediately before insertion', async () => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const base64 = await packageZip.generateAsync({ type: 'base64' })
    const source = { id: 'source', load: vi.fn() }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => source),
    }
    const insertSlidesFromBase64 = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync: vi.fn() }),
      },
    })
    const subject = new BrowserPowerPointAdapter()
    vi.spyOn(subject, 'snapshotSlide')
      .mockResolvedValueOnce({ slideId: 'source', fingerprint: 'source:before' })
      .mockRejectedValueOnce(new Error('invalid_tool_input'))
      .mockResolvedValue({ slideId: 'source', fingerprint: 'source:user-edit' })
    vi.spyOn(subject, 'exportSlidePackage').mockResolvedValue({
      slideId: 'source',
      base64,
      fingerprint: 'volatile',
    })

    await expect(subject.duplicateSlide(0)).rejects.toThrow('office_concurrent_change')
    expect(insertSlidesFromBase64).not.toHaveBeenCalled()
  })

  it('removes an owned duplicate when the source changes after insertion', async () => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const slidePackage = await packageZip.generateAsync({ type: 'base64' })
    const shape = (text: string) => ({
      id: '2',
      name: 'Title',
      type: 'TextBox',
      left: 10,
      top: 20,
      width: 200,
      height: 40,
      textFrame: { hasText: true, load: vi.fn(), textRange: { text, load: vi.fn() } },
    })
    const sourceShape = shape('Stable')
    const slides = {
      items: [] as Array<Record<string, unknown>>,
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
    }
    const source = {
      id: 'source',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [sourceShape] },
      exportAsBase64: vi.fn(() => ({ value: slidePackage })),
    }
    slides.items.push(source)
    const insertSlidesFromBase64 = vi.fn(() => {
      const copy = {
        id: 'copy',
        load: vi.fn(),
        shapes: { load: vi.fn(), items: [shape('Stable')] },
        exportAsBase64: vi.fn(() => ({ value: slidePackage })),
        delete: vi.fn(() => slides.items.splice(slides.items.indexOf(copy), 1)),
      }
      slides.items.splice(1, 0, copy)
      ;(sourceShape.textFrame as { textRange: { text: string } }).textRange.text = 'User edit'
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync: vi.fn() }),
      },
    })

    await expect(new BrowserPowerPointAdapter().duplicateSlide(0)).rejects.toThrow(
      'office_concurrent_change',
    )
    expect(slides.items).toHaveLength(2)
    expect(slides.items[0]).toBe(source)
  })

  it('keeps duplicate validation stable when PowerPoint exports volatile slide packages', async () => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const textRange = { text: 'Stable title', load: vi.fn() }
    const textFrame = { hasText: true, load: vi.fn(), textRange }
    const shape = {
      id: '2',
      name: 'Title 1',
      type: 'Placeholder',
      left: 120,
      top: 88.4,
      width: 720,
      height: 188,
      textFrame,
    }
    let exportSequence = 0
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [shape] },
      exportAsBase64: vi.fn(() => ({ value: `volatile-package-${exportSequence++}` })),
    }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides }, sync }),
      },
    })

    const subject = new BrowserPowerPointAdapter()
    const first = await subject.snapshotSlide(0)
    const second = await subject.snapshotSlide(0)

    expect(second).toEqual(first)
    textRange.text = 'Changed title'
    await expect(subject.snapshotSlide(0)).resolves.not.toEqual(first)
  })

  it('rejects empty or oversized duplicate exports before insertion', async () => {
    const sync = vi.fn().mockResolvedValue(undefined)
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { load: vi.fn(), items: [] },
      exportAsBase64: vi.fn(() => ({ value: '' })),
    }
    const slides = {
      items: [slide],
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    const insertSlidesFromBase64 = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })
    const subject = new BrowserPowerPointAdapter()
    await expect(subject.duplicateSlide(0)).rejects.toThrow('office_write_failed')
    slide.exportAsBase64.mockReturnValueOnce({ value: 'x'.repeat(8 * 1024 * 1024 + 1) })
    await expect(subject.duplicateSlide(0)).rejects.toThrow('office_write_failed')
    const controller = new AbortController()
    sync.mockClear()
    sync.mockImplementation(async () => {
      if (sync.mock.calls.length === 3) controller.abort()
    })
    slide.exportAsBase64.mockReturnValueOnce({ value: 'ppt' })
    await expect(subject.duplicateSlide(0, controller.signal)).rejects.toThrow('cancelled')
    expect(insertSlidesFromBase64).not.toHaveBeenCalled()
  })

  it('cancels package replacement before queuing its irreversible insert/delete batch', async () => {
    const controller = new AbortController()
    const sync = vi.fn().mockImplementation(async () => {
      if (sync.mock.calls.length === 2) controller.abort()
    })
    const remove = vi.fn()
    const slide = { id: 's1', load: vi.fn(), delete: remove }
    const slides = {
      getCount: vi.fn(() => ({ value: 1 })),
      getItemAt: vi.fn(() => slide),
    }
    const insertSlidesFromBase64 = vi.fn()
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })
    await expect(
      new BrowserPowerPointAdapter().replaceSlidePackage(
        0,
        'ppt',
        false,
        undefined,
        controller.signal,
      ),
    ).rejects.toThrow('cancelled')
    expect(insertSlidesFromBase64).not.toHaveBeenCalled()
    expect(remove).not.toHaveBeenCalled()
  })

  it('rejects a stale chart package preimage inside the host adapter before insert/delete', async () => {
    const zip = new JSZip()
    zip.file('ppt/slides/slide1.xml', '<p:sld/>')
    const original = await zip.generateAsync({ type: 'base64' })
    const digest = await presentationPackageDigest(original)
    const remove = vi.fn(),
      insertSlidesFromBase64 = vi.fn()
    const slide = {
      id: 's1',
      load: vi.fn(),
      delete: remove,
      exportAsBase64: vi.fn(() => ({ value: original })),
    }
    const slides = { getCount: vi.fn(() => ({ value: 1 })), getItemAt: vi.fn(() => slide) }
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync: vi.fn() }),
      },
    })
    await expect(
      new BrowserPowerPointAdapter().replaceSlidePackage(0, original, false, undefined, undefined, {
        slideId: 's1',
        packageDigest: '0'.repeat(64),
      }),
    ).rejects.toThrow('proposal_stale')
    await expect(
      new BrowserPowerPointAdapter().replaceSlidePackage(0, original, false, undefined, undefined, {
        slideId: 'wrong',
        packageDigest: digest,
      }),
    ).rejects.toThrow('proposal_stale')
    expect(insertSlidesFromBase64).not.toHaveBeenCalled()
    expect(remove).not.toHaveBeenCalled()
  })

  it('never restores a package replacement over an unowned third state', async () => {
    const originalZip = new JSZip()
    originalZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>')
    const originalPackage = await originalZip.generateAsync({ type: 'base64' })
    const expected = await editPowerPointPackage(originalPackage, 'slide', [
      {
        path: 'ppt/slides/slide1.xml',
        xml: '<p:sld xmlns:p="urn:p"><p:cSld><p:sp/></p:cSld></p:sld>',
      },
    ])
    const thirdZip = new JSZip()
    thirdZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"><p:user/></p:sld>')
    const thirdPackage = await thirdZip.generateAsync({ type: 'base64' })
    const slides = {
      items: [] as Array<Record<string, unknown>>,
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
    }
    const makeSlide = (id: string, exported: string) => {
      const item = {
        id,
        load: vi.fn(),
        exportAsBase64: vi.fn(() => ({ value: exported })),
        delete: vi.fn(() => {
          slides.items = slides.items.filter((slide) => slide !== item)
        }),
      }
      return item
    }
    const original = makeSlide('source', originalPackage)
    const third = makeSlide('user-slide', thirdPackage)
    slides.items = [original]
    let batchQueued = false
    let failed = false
    const insertSlidesFromBase64 = vi.fn((value: string) => {
      batchQueued = true
      slides.items.splice(0, 0, makeSlide('imported', value))
    })
    const sync = vi.fn().mockImplementation(async () => {
      if (batchQueued && !failed) {
        failed = true
        slides.items = [third]
        throw new Error('host rejected after third-party replacement')
      }
    })
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({ presentation: { slides, insertSlidesFromBase64 }, sync }),
      },
    })

    await expect(
      new BrowserPowerPointAdapter().replaceSlidePackage(0, expected.base64, false, expected),
    ).rejects.toThrow(/office_(concurrent_change|state_uncertain)/)
    expect(slides.items).toEqual([third])
    expect(insertSlidesFromBase64).toHaveBeenCalledOnce()
  })

  it.each([
    'sync-failure',
    'ignored-package-import',
    'ignored-layout-recovery',
    'wrong-restored-slide',
  ])('never performs an unowned package or layout restore after %s', async (mode) => {
    const packageZip = new JSZip()
    packageZip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="urn:p"/>')
    const originalPackage = await packageZip.generateAsync({ type: 'base64' })
    const expected = await editPowerPointPackage(originalPackage, 'slide', [
      { path: 'ppt/slides/slide1.xml', xml: '<p:sld xmlns:p="urn:p"><p:cSld/></p:sld>' },
    ])
    const oldLayout = { id: 'old-layout', name: '', load: vi.fn() }
    const newLayout = { id: 'new-layout', name: '', load: vi.fn() }
    const wrongLayout = { id: 'wrong-layout', name: '', load: vi.fn() }
    const slides: {
      items: any[]
      getCount: ReturnType<typeof vi.fn>
      getItemAt: ReturnType<typeof vi.fn>
      load: ReturnType<typeof vi.fn>
    } = {
      items: [],
      getCount: vi.fn(() => ({ value: slides.items.length })),
      getItemAt: vi.fn((index: number) => slides.items[index]),
      load: vi.fn(),
    }
    const makeSlide = (id: string, layout: any, exported = 'original') => {
      const item: any = {
        id,
        layout,
        load: vi.fn(),
        exportAsBase64: vi.fn(() => ({ value: exported })),
        applyLayout: vi.fn((next) => {
          item.layout = next
        }),
      }
      item.delete = vi.fn(() => {
        slides.items = slides.items.filter((slide) => slide !== item)
      })
      return item
    }
    const original = makeSlide('s1', oldLayout, originalPackage)
    const sibling = makeSlide('s-other', oldLayout)
    slides.items = [original, sibling]
    const oldMaster = { layouts: { items: [oldLayout], load: vi.fn() } }
    const newMaster = { layouts: { items: [newLayout], load: vi.fn() } }
    const masters = { items: [oldMaster, newMaster], load: vi.fn() }
    let propagationStarted = false
    sibling.applyLayout.mockImplementation((next: unknown) => {
      if (mode === 'ignored-layout-recovery') {
        if (next === newLayout) sibling.layout = wrongLayout
      } else {
        sibling.layout = next
      }
      propagationStarted = true
    })
    let failed = false
    const sync = vi.fn().mockImplementation(async () => {
      if (mode !== 'ignored-layout-recovery' && propagationStarted && !failed) {
        failed = true
        throw new Error('host failure')
      }
    })
    const insertSlidesFromBase64 = vi.fn(
      (base64: string, _options: { formatting: string; targetSlideId?: string }) => {
        const inserted = makeSlide(
          base64 === expected.base64 ? 's2' : 's1-restored',
          base64 === expected.base64 ? newLayout : oldLayout,
          mode === 'ignored-package-import' && base64 === expected.base64
            ? originalPackage
            : mode === 'wrong-restored-slide' && base64 === originalPackage
              ? expected.base64
              : base64,
        )
        slides.items.splice(0, 0, inserted)
      },
    )
    Object.assign(globalThis, {
      Office: {
        context: {
          host: 'PowerPoint',
          requirements: { isSetSupported: vi.fn().mockReturnValue(true) },
        },
      },
      PowerPoint: {
        run: (callback: (context: unknown) => unknown) =>
          callback({
            presentation: { slides, slideMasters: masters, insertSlidesFromBase64 },
            sync,
          }),
      },
    })
    const failure = await new BrowserPowerPointAdapter()
      .replaceSlidePackage(0, expected.base64, true, expected)
      .catch((error: unknown) => error)
    expect(failure).toMatchObject({ message: 'office_state_uncertain' })
    if (mode === 'ignored-package-import')
      expect(failure).toMatchObject({
        debugInfo: { errorLocation: 'PowerPoint.replaceSlidePackage.packageImportVerify' },
      })
    if (mode === 'ignored-layout-recovery')
      expect(failure).toMatchObject({
        debugInfo: { errorLocation: 'PowerPoint.replaceSlidePackage.layoutApplyVerify' },
      })
    expect(insertSlidesFromBase64).toHaveBeenCalledOnce()
    expect(insertSlidesFromBase64.mock.calls[0]?.[1]).toMatchObject({
      formatting: 'KeepSourceFormatting',
    })
    expect(slides.items).toHaveLength(2)
    expect(slides.items[0].id).toBe('s2')
  })

  it('treats a text-only change as stale even when slide geometry is unchanged', async () => {
    const fake = adapter()
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: fake, proposals })
    await skill.executeTool(
      call('edit_slide_text', { slide_index: 0, shape_id: '2', text: 'Replacement' }),
    )
    ;(fake.readSlideText as ReturnType<typeof vi.fn>).mockResolvedValueOnce({
      slideId: 'slide-1',
      shapeId: '2',
      text: 'Changed elsewhere',
      paragraphs: ['Changed elsewhere'],
    })
    await expect(proposals.confirm(proposals.pending()!.id)).rejects.toThrow('proposal_stale')
    expect(fake.editSlideText).not.toHaveBeenCalled()
  })
})
