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
    presentation: { available: () => false, documentId: async () => 'doc', request },
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
