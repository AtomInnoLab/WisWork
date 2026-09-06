import { expect, it, vi } from 'vitest'
import { boundedScreenshot } from '../src/renderer/ai/bounded-screenshot'

it('re-renders large covers without truncating PNG data', async () => {
  const render = vi.fn(async (ratio: number) => (ratio === 1 ? 'a'.repeat(1_100_000) : 'small-png'))
  expect(await boundedScreenshot(render)).toEqual({ mime: 'image/png', base64: 'small-png' })
  expect(render.mock.calls.map(([ratio]) => ratio)).toEqual([1, 0.75])
})
it('reports unavailable if a readable resolution cannot fit', async () => {
  expect(await boundedScreenshot(async () => 'a'.repeat(1_100_000))).toBeNull()
})
