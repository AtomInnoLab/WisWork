import { expect, test } from '@playwright/test'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import {
  addElement,
  createBlankPptx,
  duplicateSlide,
  fingerprintSlide,
  openPptx,
  savePptx,
} from '@wiswork/pptx-engine'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

async function eightPageFixture(): Promise<{ path: string; untouched: string[] }> {
  const opened = await openPptx(await createBlankPptx())
  for (let index = 1; index < 8; index += 1) {
    if (!duplicateSlide(opened, index - 1)) throw new Error('failed to create fixture slide')
  }
  for (const [index, slide] of opened.deck.slides.entries()) {
    const number = index + 1
    for (const [role, y] of [
      ['title', number === 6 ? 48 : 60],
      ['body', 150],
      ['emphasis', 300],
    ] as const)
      addElement(slide, {
        kind: 'textbox',
        offset: {
          x: (role === 'title' ? (number === 6 ? 72 : 80) : 90) * 12_700,
          y: y * 12_700,
          cx: (role === 'title' ? (number === 6 ? 816 : 790) : 700) * 12_700,
          cy: (role === 'title' ? (number === 6 ? 54 : 60) : 80) * 12_700,
        },
        paragraphs: [{ runs: [{ text: `${role}-${number}`, color: '#000000' }] }],
      })
  }
  const untouched = await Promise.all(
    opened.deck.slides.slice(0, 5).map((slide) => fingerprintSlide(opened, slide)),
  )
  const directory = await mkdtemp(join(tmpdir(), 'wiswork-slides-acceptance-'))
  const path = join(directory, 'eight-page-authority.pptx')
  await writeFile(path, await savePptx(opened))
  return { path, untouched }
}

test('production Slides renderer captures authority-bound affected and reference PNGs', async () => {
  const fixture = await eightPageFixture()
  const launched = await launchShell({
    onboardingSeen: true,
    videoDir: 'slides-acceptance-render',
    openFile: fixture.path,
  })
  try {
    const editor = await waitForPageWithUrl(launched.app, 'slides/out')
    await editor.waitForSelector('.stage-wrap canvas', { timeout: 20_000 })
    const execution = await editor.evaluate(async () => {
      const execute = (
        window as typeof window & {
          __wisworkSlidesRunAcceptanceAgent?: () => Promise<{
            text: string
            status?: string
            passedCheckIds?: string[]
            mutationReceiptIds?: string[]
            documentToken: string
            sessionToken: string
            revision: string
            leaseToken: string
            pngs: string[]
          }>
        }
      ).__wisworkSlidesRunAcceptanceAgent
      if (!execute) throw new Error('acceptance execution harness unavailable')
      return execute()
    })
    expect(execution.status).toBe('verified')
    expect(execution.text).toContain('Verified')
    await editor.locator('.ai-rail').click()
    await expect(editor.getByText('Verified', { exact: true })).toBeVisible()
    expect(execution.passedCheckIds).toHaveLength(17)
    expect(execution.mutationReceiptIds).toEqual([
      'golden-6-title',
      'golden-6-body',
      'golden-6-emphasis',
      'golden-7-title',
      'golden-7-body',
      'golden-7-emphasis',
      'golden-7-geometry',
      'golden-8-title',
      'golden-8-body',
      'golden-8-emphasis',
      'golden-8-geometry',
    ])
    const saved = await editor.evaluate(() => window.slidesApi.save())
    expect(saved.ok).toBe(true)
    expect(execution.documentToken).toBeTruthy()
    expect(execution.sessionToken).toBeTruthy()
    expect(execution.leaseToken).toBeTruthy()
    expect(execution.revision).toMatch(/^sha256:[a-f0-9]{64}$/)
    expect(execution.pngs).toHaveLength(4)
    for (const encoded of execution.pngs) {
      const png = Buffer.from(encoded, 'base64')
      expect(png.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]))
      expect(png.byteLength).toBeGreaterThan(1_000)
    }
    const reopened = await openPptx(
      await import('node:fs/promises').then(({ readFile }) => readFile(fixture.path)),
    )
    await expect(
      Promise.all(
        reopened.deck.slides.slice(0, 5).map((slide) => fingerprintSlide(reopened, slide)),
      ),
    ).resolves.toEqual(fixture.untouched)
  } finally {
    await closeAndSaveVideo(launched, 'slides-acceptance-render')
  }
})
