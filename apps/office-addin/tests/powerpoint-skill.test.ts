import { afterEach, describe, expect, it, vi } from 'vitest'
import JSZip from 'jszip'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  BrowserPowerPointAdapter,
  type PowerPointAdapter,
} from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'
import { editPowerPointPackage } from '../src/skills/powerpoint/powerpoint-package.js'

const png = 'iVBORw0KGgoAAAA='

function adapter(overrides: Partial<PowerPointAdapter> = {}): PowerPointAdapter {
  return {
    screenshotSlide: vi.fn().mockResolvedValue({ mime: 'image/png', base64: png }),
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
  it('exposes exactly the ten documented host tools with exact schemas', () => {
    const skill = createPowerPointSkill({
      adapter: adapter(),
      proposals: createStructuredProposalController(),
    })
    expect(skill.tools.map((tool) => tool.name)).toEqual([
      'screenshot_slide',
      'list_slide_shapes',
      'read_slide_text',
      'verify_slides',
      'execute_office_js',
      'edit_slide_text',
      'edit_slide_xml',
      'edit_slide_chart',
      'edit_slide_master',
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
    expect(skill.tools.find((tool) => tool.name === 'execute_office_js')?.description).toContain(
      '{"version":1,"operations":[{"op":"add_text_box","slide_index":0,"name":"Status","text":"PASS","left":300,"top":450,"width":360,"height":50}]}',
    )
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
    await expect(skill.executeTool(call('verify_slides', { nope: true }))).resolves.toMatchObject({
      output: 'invalid_tool_input',
      isError: true,
    })
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
    expect(fake.verifySlides).toHaveBeenCalledTimes(2)
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
        'edit_slide_master',
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

  it('maps adapter internals to stable errors and validates screenshots', async () => {
    const skill = createPowerPointSkill({
      adapter: adapter({
        listSlideShapes: vi.fn().mockRejectedValue(new Error('secret')),
        screenshotSlide: vi.fn().mockResolvedValue({ mime: 'image/png', base64: 'bad!' }),
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
      exportAsBase64: vi.fn(() => ({ value: 'ppt' })),
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
    expect(insertSlidesFromBase64).toHaveBeenCalledWith('ppt', { targetSlideId: 's1' })

    const controller = new AbortController()
    controller.abort()
    await expect(subject.editSlideText(0, '2', 'No', controller.signal)).rejects.toThrow(
      'cancelled',
    )
    expect(textRange.text).toBe('New')
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

  it.each([
    { mode: 'sync-failure', expectedError: 'office_write_failed' },
    { mode: 'ignored-layout-recovery', expectedError: 'office_recovery_failed' },
    { mode: 'wrong-restored-slide', expectedError: 'office_recovery_failed' },
  ])('proves package and layout recovery after $mode', async ({ mode, expectedError }) => {
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
    const insertSlidesFromBase64 = vi.fn((base64: string) => {
      const inserted = makeSlide(
        base64 === expected.base64 ? 's2' : 's1-restored',
        base64 === expected.base64 ? newLayout : oldLayout,
        mode === 'wrong-restored-slide' && base64 === originalPackage ? expected.base64 : base64,
      )
      slides.items.splice(0, 0, inserted)
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
          callback({
            presentation: { slides, slideMasters: masters, insertSlidesFromBase64 },
            sync,
          }),
      },
    })
    await expect(
      new BrowserPowerPointAdapter().replaceSlidePackage(0, expected.base64, true, expected),
    ).rejects.toThrow(expectedError)
    if (mode === 'sync-failure') expect(sibling.layout).toBe(oldLayout)
    expect(slides.items).toHaveLength(2)
    expect(slides.items[0].id).toBe('s1-restored')
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
