import { spawnSync } from 'node:child_process'
import { expect, it } from 'vitest'
import { compilePresentationDeck } from '@wiswork/pptx-engine/presentation-compiler'
import { benchmarkPlannedDeck } from '../../../packages/pptx-engine/tests/fixtures/presentation-plan'
import { convertSinglePagePackageToPng } from '../src/main/presentation-page-render'

const sofficeAvailable = spawnSync('soffice', ['--version'], { timeout: 5_000 }).status === 0

it.skipIf(!sofficeAvailable)(
  'converts a real one-page PPTX to PNG for the fallback renderer',
  async () => {
    const deck = benchmarkPlannedDeck()
    const { bytes } = await compilePresentationDeck({ ...deck, slides: [deck.slides[0]!] })
    const png = Buffer.from(
      await convertSinglePagePackageToPng(bytes, new AbortController().signal),
    )
    expect(png.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
    expect(png.readUInt32BE(16)).toBeGreaterThan(0)
    expect(png.readUInt32BE(20)).toBeGreaterThan(0)
  },
)

it('rejects cancellation before starting a converter', async () => {
  const controller = new AbortController()
  controller.abort()
  await expect(
    convertSinglePagePackageToPng(new Uint8Array([1]), controller.signal),
  ).rejects.toThrow('aborted')
})

it.skipIf(!sofficeAvailable)('stops a running converter after cancellation', async () => {
  const deck = benchmarkPlannedDeck()
  const { bytes } = await compilePresentationDeck({ ...deck, slides: [deck.slides[0]!] })
  const controller = new AbortController()
  const rendering = convertSinglePagePackageToPng(bytes, controller.signal)
  setTimeout(() => controller.abort(), 50)
  await expect(rendering).rejects.toThrow('aborted')
})
