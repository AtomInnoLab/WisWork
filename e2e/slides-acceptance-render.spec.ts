import { expect, test } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  addElement,
  createBlankPptx,
  duplicateSlide,
  openPptx,
  savePptx,
} from '@wiswork/pptx-engine'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

async function eightPageFixture(): Promise<string> {
  const opened = await openPptx(await createBlankPptx())
  for (let index = 1; index < 8; index += 1) {
    if (!duplicateSlide(opened, index - 1)) throw new Error('failed to create fixture slide')
  }
  for (const [index, slide] of opened.deck.slides.entries()) {
    addElement(slide, {
      kind: 'textbox',
      offset: { x: 900_000, y: 500_000, cx: 5_000_000, cy: 700_000 },
      paragraphs: [{ runs: [{ text: `Page ${index + 1}`, color: '#2457A7', bold: true }] }],
    })
  }
  const directory = await mkdtemp(join(tmpdir(), 'wiswork-slides-acceptance-'))
  const path = join(directory, 'eight-page-authority.pptx')
  await writeFile(path, await savePptx(opened))
  return path
}

test('production Slides renderer captures authority-bound affected and reference PNGs', async () => {
  const fixture = await eightPageFixture()
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'slides-acceptance-render',
    openFile: fixture,
  })
  try {
    const editor = await waitForPageWithUrl(launched.app, 'slides/out')
    await editor.waitForSelector('.stage-wrap canvas', { timeout: 20_000 })
    const result = await editor.evaluate(async () => {
      const render = (
        window as typeof window & {
          __wisworkSlidesRenderAcceptance?: (slides: number[]) => Promise<{
            documentToken: string
            sessionToken: string
            revision: string
            leaseToken: string
            pngs: string[]
          }>
        }
      ).__wisworkSlidesRenderAcceptance
      if (!render) throw new Error('acceptance render harness unavailable')
      return render([6, 7, 8, 6])
    })

    expect(result.documentToken).toBeTruthy()
    expect(result.sessionToken).toBeTruthy()
    expect(result.leaseToken).toBeTruthy()
    expect(result.revision).toMatch(/^sha256:[a-f0-9]{64}$/)
    expect(result.pngs).toHaveLength(4)
    for (const encoded of result.pngs) {
      const png = Buffer.from(encoded, 'base64')
      expect(png.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
      expect(png.byteLength).toBeGreaterThan(1_000)
    }
  } finally {
    await closeAndSaveVideo(launched, 'slides-acceptance-render')
  }
})
