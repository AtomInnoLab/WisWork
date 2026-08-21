import { afterEach, describe, expect, it, vi } from 'vitest'
import { createStructuredProposalController } from '../src/agent/proposal-controller.js'
import {
  BrowserPowerPointAdapter,
  type PowerPointAdapter,
} from '../src/skills/powerpoint/browser-powerpoint-adapter.js'
import { createPowerPointSkill } from '../src/skills/powerpoint/powerpoint-skill.js'

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
      output: expect.stringContaining('"visualAvailableToModel":false'),
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

  it('keeps unaudited raw-code and OOXML tools visible but fails closed without proposals', async () => {
    const proposals = createStructuredProposalController()
    const skill = createPowerPointSkill({ adapter: adapter(), proposals })
    for (const [name, input] of [
      ['execute_office_js', { code: 'return context.presentation' }],
      ['edit_slide_xml', { slide_index: 0, code: 'zip.file("x")' }],
      ['edit_slide_chart', { slide_index: 0, code: 'markDirty()' }],
      ['edit_slide_master', { code: 'markDirty()' }],
    ] as const) {
      await expect(skill.executeTool(call(name, input))).resolves.toMatchObject({
        output: 'office_api_unsupported',
        isError: true,
        mutated: false,
      })
      expect(proposals.pending()).toBeUndefined()
    }
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
    const shape = { id: '2', textFrame: { textRange } }
    const slide = {
      id: 's1',
      load: vi.fn(),
      shapes: { getItem: vi.fn(() => shape) },
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
