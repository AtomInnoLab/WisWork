import { afterEach, expect, it, vi } from 'vitest'
import { createOfficeHostRuntime } from '../src/agent/host-runtime'
import { BrowserPresentationBaselineAdapter } from '../src/skills/powerpoint/browser-presentation-baseline-adapter'
import { BrowserPowerPointAdapter } from '../src/skills/powerpoint/browser-powerpoint-adapter'
afterEach(() => {
  vi.restoreAllMocks()
  vi.unstubAllGlobals()
})
it('offers existing-deck baselines while the PC is offline and without generated artifacts', async () => {
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readContext').mockResolvedValue({
    slideIds: ['native'],
    selectedSlideIds: ['native'],
    selectedShapeIds: [],
  })
  vi.spyOn(BrowserPresentationBaselineAdapter.prototype, 'readPage').mockResolvedValue({
    slideId: 'native',
    shapes: [],
  })
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectSlideMasters').mockRejectedValue(
    new Error('office_api_unsupported'),
  )
  const request = vi.fn(async () => {
    throw new Error('offline')
  })
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => false,
      documentId: async () => 'doc',
      request,
      lastProject: () => undefined,
      rememberProject: async () => {},
    },
  })
  expect(runtime.skill.tools.map((t) => t.name)).toContain('read_presentation_baseline')
  const response = await runtime.skill.executeTool({
    id: 'read',
    name: 'read_presentation_baseline',
    input: { scope: 'current' },
  })
  expect(response.isError, response.output).not.toBe(true)
  expect(JSON.parse(response.output).pages).toEqual([{ slideId: 'native', shapes: [] }])
  expect(request).not.toHaveBeenCalled()
  runtime.clearSession()
  expect(
    (
      await runtime.skill.executeTool({
        id: 'check',
        name: 'check_presentation_baseline',
        input: { baseline_id: JSON.parse(response.output).baselineId },
      })
    ).output,
  ).toBe('presentation_baseline_missing')
})
it('registers baseline reading in a local PowerPoint runtime independently of the presentation service', () => {
  const runtime = createOfficeHostRuntime('powerpoint')
  expect(runtime.skill.tools.map((t) => t.name)).toContain('read_presentation_baseline')
  expect(createOfficeHostRuntime('word').skill.tools.map((t) => t.name)).not.toContain(
    'read_presentation_baseline',
  )
})
it('detects manual title-placeholder edits through the real adapter and runtime', async () => {
  const load = vi.fn()
  const range = {
    text: 'Original title',
    load,
    font: { name: 'Arial', size: 24, color: '#000000', load },
  }
  const frame = { isNullObject: false, load, textRange: range }
  const shape = {
    id: 'title',
    name: 'Title',
    type: 'Placeholder',
    left: 0,
    top: 0,
    width: 100,
    height: 30,
    getTextFrameOrNullObject: () => frame,
  }
  const slide = {
    id: 'slide',
    load,
    shapes: { items: [shape], load },
    slideMaster: { id: 'master', load },
    layout: { id: 'layout', load },
  }
  const context = {
    sync: async () => {},
    presentation: {
      slides: { items: [slide], load, getItem: () => slide },
      getSelectedSlides: () => ({ items: [slide], load }),
      getSelectedShapes: () => ({ items: [shape], load }),
      pageSetup: { slideWidth: 960, slideHeight: 540, load },
    },
  }
  vi.stubGlobal('Office', {
    context: { host: 'PowerPoint', requirements: { isSetSupported: () => true } },
  })
  vi.stubGlobal('PowerPoint', {
    run: async (callback: (value: typeof context) => unknown) => callback(context),
  })
  vi.spyOn(BrowserPowerPointAdapter.prototype, 'inspectSlideMasters').mockRejectedValue(
    new Error('office_api_unsupported'),
  )
  const runtime = createOfficeHostRuntime('powerpoint', {
    presentation: {
      available: () => false,
      documentId: async () => 'doc',
      request: async () => {
        throw new Error('offline')
      },
      lastProject: () => undefined,
      rememberProject: async () => {},
    },
  })
  const read = await runtime.skill.executeTool({
    id: 'read',
    name: 'read_presentation_baseline',
    input: { scope: 'selected' },
  })
  expect(read.isError, read.output).not.toBe(true)
  const baseline = JSON.parse(read.output)
  expect(baseline.pages[0].shapes[0].text).toBe('Original title')
  range.text = 'User edited title'
  const checked = await runtime.skill.executeTool({
    id: 'check',
    name: 'check_presentation_baseline',
    input: { baseline_id: baseline.baselineId },
  })
  expect(JSON.parse(checked.output)).toMatchObject({ unchanged: false, changedSlideIds: ['slide'] })
})
